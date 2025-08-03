"Handles NVDA WavePlayer audio capturing."

from ctypes import _Pointer, addressof, c_ubyte, c_void_p
import threading
import time
import typing
from typing_extensions import Buffer
from extensionPoints import Action
from nvwave import AudioPurpose, WavePlayer
from ._wavestorage import WaveStorage
import speech
import os
from logHandler import log


_original_feed = WavePlayer.feed
_original_sync = WavePlayer.sync
_original_stop = WavePlayer.stop

_player_num = 1
_player_num_lock = threading.Lock()
_max_duration_sec: int = 30
_pre_speech: Action

NS_PER_MILLISEC = 1_000_000


# When a WavePlayer is fed data the first time,
# a unique name will be assigned to it as the wave file name.
# When a player has been idle for more than the allowed duration,
# the item and recorded audio for the player will be removed.
class _PlayerInfo:
	def __init__(self, player: WavePlayer):
		global _player_num
		with _player_num_lock:
			self.fileName = f"WavePlayer {_player_num}.wav"
			_player_num += 1
		self.lastActivePerfCounter = time.perf_counter()
		self.waveStorage = WaveStorage(
			player.channels,
			player.bitsPerSample // 8,
			player.samplesPerSec,
			player.samplesPerSec * _max_duration_sec,
		)
		self.lock = threading.Lock()


_players: dict[WavePlayer, _PlayerInfo] = dict()


def _player_feed(
	self: WavePlayer,
	data: Buffer | _Pointer | None,
	size: int | None = None,
	onDone: typing.Callable | None = None,
	*args,
	**kwargs,
) -> None:
	if self._purpose != AudioPurpose.SPEECH:
		_original_feed(self, data, size, onDone, *args, **kwargs)
		return

	if size is not None:
		arrType = c_ubyte * size
		if data is None:
			d = None
		elif isinstance(data, c_void_p):  # c_void_p pointer
			d = arrType.from_address(data.value) if data.value else None
		elif isinstance(data, _Pointer):  # other ctypes pointer
			d = arrType.from_address(addressof(data.contents))
		else:
			d = data
		if d is not None:
			d = memoryview(d).cast("B")[:size]
	else:
		if data is None:
			d = None
		else:
			d = memoryview(data).cast("B")
	if not d:
		_original_feed(self, data, size, onDone, *args, **kwargs)
		return

	if self in _players:
		player = _players[self]
	else:
		player = _PlayerInfo(self)
		_players[self] = player

	framesize = self.channels * self.bitsPerSample // 8
	maxblocksize = framesize * self.samplesPerSec // 5  # send 200ms max per block

	if size is None:
		size = len(d)

	for i in range(0, size, maxblocksize):
		block = d[i : i + maxblocksize]
		_original_feed(self, block.tobytes(), len(block), onDone, *args, **kwargs)
		with player.lock:
			player.waveStorage.write(block)
			player.lastActivePerfCounter = time.perf_counter()

	now = time.perf_counter()
	for k, v in list(_players.items()):
		if v.lastActivePerfCounter - now > _max_duration_sec:
			del _players[k]


def _player_sync(self: WavePlayer, *args, **kwargs):
	_original_sync(self, *args, **kwargs)
	if self in _players:
		player = _players[self]
		with player.lock:
			player.waveStorage.addmarker("Sync")


def _player_stop(self: WavePlayer, *args, **kwargs):
	_original_stop(self, *args, **kwargs)
	if self in _players:
		player = _players[self]
		with player.lock:
			player.waveStorage.addmarker("Stop")


def _add_speech_marker(number: int, speechSequence: speech.SpeechSequence):
	for info in _players.values():
		info.waveStorage.addmarker(f"#{number}: " + "".join(i for i in speechSequence if isinstance(i, str)))


def start(max_duration_sec: int, pre_speech: Action):
	global _player_num, _max_duration_sec, _pre_speech
	_player_num = 1
	WavePlayer.feed = _player_feed
	WavePlayer.sync = _player_sync
	WavePlayer.stop = _player_stop
	_max_duration_sec = max_duration_sec
	_pre_speech = pre_speech
	pre_speech.register(_add_speech_marker)


def stop():
	WavePlayer.feed = _original_feed
	WavePlayer.sync = _original_sync
	WavePlayer.stop = _original_stop
	_pre_speech.unregister(_add_speech_marker)


def savetodir(dirpath: str):
	# save files
	for info in _players.values():
		info.waveStorage.savetofile(os.path.join(dirpath, info.fileName))
	log.info(f"Recorded wave player audio files saved to {dirpath}")


def reset():
	_players.clear()
