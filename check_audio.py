import time
import struct
import ctypes
from ctypes import cast, POINTER, c_void_p
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioClient

WAVEFORMATEX_SIZE = 18  # bytes for the base WAVEFORMATEX

def _get_ptr_address(ptr):
    try:
        return ctypes.addressof(ptr.contents)
    except Exception:
        # fallback: cast to void*
        return ctypes.cast(ptr, c_void_p).value

def parse_waveformat(ptr):
    addr = _get_ptr_address(ptr)
    if not addr:
        raise RuntimeError("GetMixFormat returned NULL")

    # read the base WAVEFORMATEX (18 bytes)
    base = ctypes.string_at(addr, WAVEFORMATEX_SIZE)
    wFormatTag, nChannels, nSamplesPerSec, nAvgBytesPerSec, nBlockAlign, wBitsPerSample, cbSize = \
        struct.unpack('<HHIIHHH', base)

    bits = wBitsPerSample

    if wFormatTag == 0xFFFE and cbSize >= 22:
        ext_bytes = ctypes.string_at(addr + WAVEFORMATEX_SIZE, 22)
        # wValidBitsPerSample: WORD, dwChannelMask: DWORD, SubFormat: 16 bytes
        wValidBitsPerSample, dwChannelMask, subformat = struct.unpack('<HI16s', ext_bytes)
        if wValidBitsPerSample != 0:
            bits = wValidBitsPerSample

    try:
        ptr_val = ctypes.cast(ptr, c_void_p).value
        if ptr_val:
            ctypes.windll.ole32.CoTaskMemFree(ptr_val)
    except Exception:
        pass

    return {
        'format_tag': wFormatTag,
        'format_tag_str': {1: 'PCM', 3: 'IEEE_FLOAT', 0xFFFE: 'EXTENSIBLE'}.get(wFormatTag, hex(wFormatTag)),
        'channels': nChannels,
        'sample_rate': nSamplesPerSec,
        'bits_per_sample': bits,
        'avg_bytes_per_sec': nAvgBytesPerSec,
        'block_align': nBlockAlign,
        'cbSize': cbSize,
    }

def get_default_output_format():
    device = AudioUtilities.GetSpeakers()  # default render device
    interface = device.Activate(IAudioClient._iid_, CLSCTX_ALL, None)
    audio_client = cast(interface, POINTER(IAudioClient))
    mix_ptr = audio_client.GetMixFormat()
    return parse_waveformat(mix_ptr)

if __name__ == '__main__':
    prev = None
    try:
        while True:
            try:
                fmt = get_default_output_format()
                if fmt != prev:
                    print(f"Sample Rate: {fmt['sample_rate']} Hz | Bit Depth: {fmt['bits_per_sample']} | "
                          f"Channels: {fmt['channels']} | Format: {fmt['format_tag_str']} | "
                          "Ctrl+C to stop.")
                    prev = fmt
            except Exception as e:
                print("Error reading format:", e)
            time.sleep(1.0)
    except KeyboardInterrupt:
        print("\nStopped.")