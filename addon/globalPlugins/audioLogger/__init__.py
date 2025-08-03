from extensionPoints import Action
import globalPluginHandler
from scriptHandler import script
from logHandler import log
import speech
from . import _sysaudio, _nvdaaudio
import synthDriverHandler
import os
from comtypes import GUID
from ctypes.wintypes import DWORD, HANDLE, LPWSTR
from ctypes import POINTER, byref, oledll, windll
from datetime import datetime
import shutil
import tones

oledll.shell32.SHGetKnownFolderPath.argtypes = (POINTER(GUID), DWORD, HANDLE, POINTER(LPWSTR))

try:
	_nvda_pre_speech = synthDriverHandler.pre_synthSpeak
except AttributeError:
	_nvda_pre_speech = speech.extensions.pre_speech


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	def __init__(self):
		super(GlobalPlugin, self).__init__()
		self._speechNum = 0
		self._pre_speech = Action()
		self._isRecording = False

	def terminate(self):
		if self._isRecording:
			self.stopCapture()

	@script("Start audio logging", gesture="kb:NVDA+Alt+R")
	def script_startCapture(self, gesture):
		self.startCapture()

	@script("Stop audio logging", gesture="kb:NVDA+Alt+T")
	def script_stopCapture(self, gesture):
		self.stopCapture()

	def _speechHandler(self, speechSequence: speech.SpeechSequence):
		self._speechNum += 1
		log.io(f"[AudioLogger] Speech #{self._speechNum}")
		self._pre_speech.notify(number=self._speechNum, speechSequence=speechSequence)

	def startCapture(self):
		if self._isRecording:
			return
		self._gestureNum = 0
		self._speechNum = 0
		_nvda_pre_speech.register(self._speechHandler)
		self._sysRecorder = _sysaudio.SystemAudioRecorder(60, self._pre_speech)
		self._sysRecorder.start()
		log.info("Recording started")
		_nvdaaudio.start(80, self._pre_speech)
		self._isRecording = True
		tones.beep(750, 100)

	def stopCapture(self):
		if not self._isRecording:
			return
		_nvda_pre_speech.unregister(self._speechHandler)
		_nvdaaudio.stop()
		self._sysRecorder.stop()
		self._isRecording = False
		log.info("Recording stopped")
		tones.beep(500, 100)
		self.saveFiles()

	def saveFiles(self):
		try:
			FOLDERID_Documents = GUID("{FDD39AD0-238F-46AF-ADB4-6C85480369C7}")
			ppath = LPWSTR()
			oledll.shell32.SHGetKnownFolderPath(byref(FOLDERID_Documents), 0, None, byref(ppath))
			myDocuments = ppath.value
			windll.ole32.CoTaskMemFree(ppath)
			audioLoggerRoot = os.path.join(myDocuments, "AudioLogger")
			subdirName = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
			recordSubdir = os.path.join(audioLoggerRoot, subdirName)
			os.makedirs(recordSubdir)
			sysAudioPath = os.path.join(recordSubdir, "SystemAudio.wav")
			self._sysRecorder.savetofile(sysAudioPath)
			del self._sysRecorder
			_nvdaaudio.savetodir(recordSubdir)
			shutil.make_archive(recordSubdir, "zip", recordSubdir)
			shutil.rmtree(recordSubdir)
			log.info("File saved")
		except Exception:
			log.error("Cannot save audio files", exc_info=True)
