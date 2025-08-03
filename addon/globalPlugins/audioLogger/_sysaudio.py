"""Handles system audio capturing."""

from ctypes import POINTER, addressof, byref, c_uint16, cast, memmove, sizeof, windll
import threading
import time
import typing
from comtypes import COMError, CLSCTX_ALL, IUnknown
import wave
from pycaw.api.mmdeviceapi import IMMDevice, IMMDeviceCollection, IMMDeviceEnumerator, IMMEndpoint
from pycaw.api.audioclient import IAudioClient, WAVEFORMATEX as pycaw_WAVEFORMATEX
from pycaw.utils import AudioUtilities
from extensionPoints import Action
from ._wasapi import (
	AudioClientShareMode,
	AudioClientStreamFlags,
	AudioDeviceState,
	EDataFlow,
	ERole,
	WAVEFORMATEX,
	IAudioCaptureClient,
)
import config
from nvwave import WavePlayer
from ._wavestorage import WaveStorage
import speech
import inputCore
from logHandler import log


REFTIMES_PER_MS = 10_000
NS_PER_REFTIME = 100


def getNVDAOutputDevice() -> IMMDevice:
	devEnum: IMMDeviceEnumerator = AudioUtilities.GetDeviceEnumerator()
	try:
		devId = config.conf["audio"]["outputDevice"]
		if devId == WavePlayer.DEFAULT_DEVICE_KEY:
			return devEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eConsole)
		try:
			dev = devEnum.GetDevice(devId)
			state = dev.GetState()
			dataflow = dev.QueryInterface(IMMEndpoint).GetDataFlow()
			if state == AudioDeviceState.Active and dataflow == EDataFlow.eRender:
				return dev
		except COMError:
			pass
	except KeyError:
		pass
	try:
		devName = config.conf["speech"]["outputDevice"]
		devColl: IMMDeviceCollection = devEnum.EnumAudioEndpoints(EDataFlow.eRender, AudioDeviceState.Active)
		for i in range(devColl.GetCount()):
			dev = devColl.Item(i)
			if AudioUtilities.CreateDevice(dev).FriendlyName == devName:
				return dev
	except COMError:
		pass
	return devEnum.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eConsole)


class SystemAudioRecorder:
	def __init__(self, max_duration_sec: int, pre_speech: Action):
		self._max_duration_sec = max_duration_sec
		self._pre_speech = pre_speech
		self._wavestorage: WaveStorage | None = None

	def _openDevice(self) -> tuple[IAudioClient, IAudioCaptureClient]:
		dev = getNVDAOutputDevice()
		audClient = typing.cast(IUnknown, dev.Activate(IAudioClient._iid_, CLSCTX_ALL, None))
		audClient = audClient.QueryInterface(IAudioClient)

		if self._wavestorage is None:
			# Initialize wave format
			fmt = WAVEFORMATEX()
			pfmt = audClient.GetMixFormat()
			memmove(byref(fmt), pfmt, sizeof(WAVEFORMATEX))
			windll.ole32.CoTaskMemFree(pfmt)

			# Set wave format to 16-bit
			fmt.wFormatTag = wave.WAVE_FORMAT_PCM
			fmt.cbSize = 0
			fmt.wBitsPerSample = 16
			fmt.nBlockAlign = fmt.nChannels * fmt.wBitsPerSample // 8
			fmt.nAvgBytesPerSec = fmt.nBlockAlign * fmt.nSamplesPerSec

			self._format = fmt
			self._wavestorage = WaveStorage(
				fmt.nChannels, 2, fmt.nSamplesPerSec, fmt.nSamplesPerSec * self._max_duration_sec
			)

		audClient.Initialize(
			AudioClientShareMode.SHARED,
			AudioClientStreamFlags.LOOPBACK
			| AudioClientStreamFlags.AUTOCONVERTPCM
			| AudioClientStreamFlags.SRC_DEFAULT_QUALITY,
			REFTIMES_PER_MS * 1000,
			0,
			cast(byref(self._format), POINTER(pycaw_WAVEFORMATEX)),
			None,
		)

		capClient = typing.cast(IUnknown, audClient.GetService(IAudioCaptureClient._iid_))
		capClient = capClient.QueryInterface(IAudioCaptureClient)

		return audClient, capClient

	def _add_speech_marker(self, number: int, speechSequence: speech.SpeechSequence):
		if self._wavestorage:
			self._wavestorage.addmarker_at_time(
				time.perf_counter_ns(),
				f"#{number}: " + "".join(i for i in speechSequence if isinstance(i, str))
			)

	def _add_gesture_marker(self, gesture: inputCore.InputGesture) -> typing.Literal[True]:
		if self._wavestorage and not gesture.isModifier:
			self._wavestorage.addmarker_at_time(time.perf_counter_ns(), gesture.identifiers[0])
		return True

	def _captureThread(self):
		inputCore.decide_executeGesture.register(self._add_gesture_marker)
		self._pre_speech.register(self._add_speech_marker)
		audClient = None
		try:
			audClient, capClient = self._openDevice()
			assert self._wavestorage is not None
			audClient.Start()
			while not self._isStopped:
				try:
					data, frameCount, flags, devPos, qpcPos = capClient.GetBuffer()
					if frameCount == 0:
						capClient.ReleaseBuffer(0)
						time.sleep(0.2)
						continue
					try:
						arrayLen: int = frameCount * self._format.nChannels
						arrayType = c_uint16 * arrayLen
						dataArray = arrayType.from_address(addressof(data.contents))
						self._wavestorage.write(dataArray, qpcPos * NS_PER_REFTIME)
					finally:
						capClient.ReleaseBuffer(frameCount)
				except COMError as ex:
					# the device might be changed
					# try opening the device again
					audClient, capClient = self._openDevice()
					audClient.Start()
					log.info(f"Recording resumed after interruption: {ex!r}")
		except Exception:
			log.error("Recording stopped due to unrecoverable error", exc_info=True)
		finally:
			inputCore.decide_executeGesture.unregister(self._add_gesture_marker)
			self._pre_speech.unregister(self._add_speech_marker)
			if audClient:
				audClient.Stop()

	def start(self):
		self._isStopped = False
		self._thread = threading.Thread(name="SystemAudioCaptureThread", target=self._captureThread)
		self._thread.start()

	def stop(self):
		self._isStopped = True
		self._thread.join()

	def savetofile(self, file: str):
		if not self._wavestorage:
			raise ValueError("No wave storage")
		self._wavestorage.savetofile(file)
		log.info(f"Recorded system audio saved to {file}")

	def reset(self):
		self._wavestorage = None
