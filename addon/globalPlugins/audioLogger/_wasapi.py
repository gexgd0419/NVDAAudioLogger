"""Provides WASAPI definitions that are missing or incorrect in pycaw."""

from ctypes import HRESULT, POINTER, Structure, c_uint64
from ctypes.wintypes import BYTE, DWORD, WORD
from enum import IntEnum
from comtypes import COMMETHOD, GUID, IUnknown


class WAVEFORMATEX(Structure):
	_fields_ = [
		("wFormatTag", WORD),
		("nChannels", WORD),
		("nSamplesPerSec", DWORD),
		("nAvgBytesPerSec", DWORD),
		("nBlockAlign", WORD),
		("wBitsPerSample", WORD),
		("cbSize", WORD),
	]


class ERole(IntEnum):
	eConsole = 0
	eMultimedia = 1
	eCommunications = 2


class EDataFlow(IntEnum):
	eRender = 0
	eCapture = 1
	eAll = 2


class AudioDeviceState(IntEnum):
	Active = 0x1
	Disabled = 0x2
	NotPresent = 0x4
	Unplugged = 0x8


class AudioClientShareMode(IntEnum):
	SHARED = 0
	EXCLUSIVE = 1


class AudioClientStreamFlags(IntEnum):
	LOOPBACK = 0x00020000
	SRC_DEFAULT_QUALITY = 0x08000000
	AUTOCONVERTPCM = 0x80000000


class IAudioCaptureClient(IUnknown):
	_iid_ = GUID("{C8ADBD64-E71E-48a0-A4DE-185C395CD317}")
	_methods_ = [
		COMMETHOD(
			[],
			HRESULT,
			"GetBuffer",
			(["out"], POINTER(POINTER(BYTE)), "ppData"),
			(["out"], POINTER(DWORD), "pNumFramesToRead"),
			(["out"], POINTER(DWORD), "pdwFlags"),
			(["out"], POINTER(c_uint64), "pu64DevicePosition"),
			(["out"], POINTER(c_uint64), "pu64QPCPosition"),
		),
		COMMETHOD(
			[],
			HRESULT,
			"ReleaseBuffer",
			(["in"], DWORD, "NumFramesRead"),
		),
		COMMETHOD(
			[],
			HRESULT,
			"GetNextPacketSize",
			(["out"], POINTER(DWORD), "pNumFramesInNextPacket"),
		),
	]
