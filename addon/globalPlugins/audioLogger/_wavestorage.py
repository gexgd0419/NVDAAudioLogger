import wave
import os
import struct
from typing_extensions import Buffer


NS_PER_SEC = 1_000_000_000


class WaveStorage:
	"""Storage for wave data and time markers.
	When the maximum duration is reached, the oldest data will be overwritten to keep the last recorded data."""

	def __init__(self, nchannels: int, sampwidth: int, framerate: int, maxframes: int | None):
		self._nchannels = nchannels
		self._sampwidth = sampwidth
		self._framerate = framerate
		self._framesize = sampwidth * nchannels
		self._maxframes = maxframes
		self._maxsize = maxframes * self._framesize if maxframes is not None else None  # max size in bytes
		self._buffer = bytearray()  # circular buffer
		self._writeidx = 0  # where the oldest data is at
		self._nsampleswritten = 0
		self._perf_counter_ns: int | None = None  # perf counter at the end of buffer
		self._markers: list[tuple[int, str]] = []  # (samplepos, label)

	def _remove_old_markers(self):
		if self._maxframes is not None:
			firstframe = self._nsampleswritten - self._maxframes
			# log.info(f"Removing old markers, first frame={firstframe}, current markers: {self._markers}")
			self._markers = [(pos, text) for pos, text in self._markers if pos >= firstframe]

	def addmarker(self, text: str) -> None:
		"""Add a time marker at the current write position."""
		self._markers.append((self._nsampleswritten, text))
		self._remove_old_markers()

	def addmarker_at_time(self, perf_counter_ns: int, text: str) -> None:
		"""Add a time marker at the specific time."""
		if self._perf_counter_ns is None:
			self.addmarker(text)
			return
		time_delta_ns = perf_counter_ns - self._perf_counter_ns
		sample_offset = time_delta_ns * self._framerate // NS_PER_SEC
		self._markers.append((self._nsampleswritten + sample_offset, text))
		self._remove_old_markers()

	def write(self, data: Buffer, perf_counter_ns: int | None = None) -> None:
		"""Write to wave storage buffer.
		When the maximum size is reached, the oldest wave data will be overwritten.

		:param data: Binary wave data to be written. Can be a bytes-like object or ctypes array.
		:param perf_counter_ns: The time that the beginning of this data chunk corresponds to.
			If omitted, the time will not be recorded."""
		data = memoryview(data).cast("B")  # memoryview in bytes
		size = len(data)
		if size % self._framesize != 0:
			# Wave data size not aligned to frame boundary
			self.addmarker("!WAVE MISALIGNED")
			data = data[: -(size % self._framesize)]  # remove misaligned tail
			size = len(data)
		if self._maxsize is None:
			self._buffer.extend(data)
		elif size >= self._maxsize:
			# Overwrite the entire buffer
			self._buffer[:] = data[-self._maxsize :]
			self._writeidx = 0
		else:
			space = self._maxsize - len(self._buffer)
			if space > 0:  # there's still space for growing
				if space >= size:
					self._buffer.extend(data)
				else:
					remaining = size - space
					self._buffer.extend(data[:space])
					self._buffer[:remaining] = data[space:]
					self._writeidx = remaining
			else:  # buffer has reached maxsize, overwrite oldest data from _writeidx
				if self._writeidx + size <= self._maxsize:
					self._buffer[self._writeidx : self._writeidx + size] = data
					self._writeidx += size
				else:
					space = self._maxsize - self._writeidx
					remaining = size - space
					self._buffer[self._writeidx : self._maxsize] = data[:space]
					self._buffer[:remaining] = data[space:]
					self._writeidx = remaining
		self._nsampleswritten += size // self._framesize
		if perf_counter_ns is not None:
			duration_ns = size // self._framesize * NS_PER_SEC // self._framerate
			self._perf_counter_ns = perf_counter_ns + duration_ns

	def getbytes(self) -> bytes:
		if self._maxsize is None or len(self._buffer) < self._maxsize:
			return bytes(self._buffer)
		else:
			return bytes(self._buffer[self._writeidx : self._maxsize] + self._buffer[: self._writeidx])

	def getwavecuedata(self) -> bytes:
		"""Generate time marker data as a RIFF cue chunk and a LIST-adtl chunk.
		Can be appended to a wave file."""
		self._remove_old_markers()
		firstframe = max(self._nsampleswritten - self._maxframes, 0) if self._maxframes is not None else 0
		cue_points: list[bytes] = []
		label_chunks: list[bytes] = []
		for id, (pos, text) in enumerate(self._markers, start=1):
			cue_points.append(
				struct.pack(
					"<II4sIII",
					id,  # cue ID
					pos,  # sample position in playlist
					b"data",  # chunk ID that contains the cue ("data" chunk)
					0,  # data chunk start, 0 when without playlists
					0,  # block start
					pos - firstframe,  # sample offset
				)
			)
			text_bytes = text.encode("utf-8") + b"\0"
			if len(text_bytes) % 2 != 0:  # Align each chunk to WORD
				text_bytes += b"\0"
			label_chunks.append(
				struct.pack(
					"<4sII",
					b"labl",
					4 + len(text_bytes),  # size (4-byte cue ID + label text)
					id,
				)
				+ text_bytes
			)
		cue_data = b"".join(cue_points)
		cue_chunk = (
			struct.pack(
				"<4sII",
				b"cue ",
				4 + len(cue_data),  # size (4-byte cue count + data)
				len(cue_points),
			)
			+ cue_data
		)
		label_data = b"".join(label_chunks)
		list_chunk = (
			struct.pack(
				"<4sI4s",
				b"LIST",
				4 + len(label_data),  # size (4-byte "adtl" + data)
				b"adtl",  # list type
			)
			+ label_data
		)
		return cue_chunk + list_chunk

	def savetofile(self, file: str) -> None:
		with open(file, "wb") as f:
			with wave.open(f, "wb") as w:
				w.setnchannels(self._nchannels)
				w.setsampwidth(self._sampwidth)
				w.setframerate(self._framerate)
				w.writeframes(self.getbytes())
			f.seek(0, os.SEEK_END)
			cuedata = self.getwavecuedata()
			f.write(cuedata)
			filesize = f.tell()
			f.seek(4, os.SEEK_SET)
			f.write(struct.pack("<I", filesize - 8))
