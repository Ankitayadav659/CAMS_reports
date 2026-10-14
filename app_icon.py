"""
App icon drawn in code (no image files or Pillow needed): a blue rounded square with three rising white bars.

  python app_icon.py      -> (re)writes assets/app.ico
"""
import os
import struct
import zlib

HERE = os.path.dirname(os.path.abspath(__file__))
ICO_PATH = os.path.join(HERE, 'assets', 'app.ico')
SIZES = (16, 24, 32, 48, 64, 128, 256)

_TOP, _BOTTOM = (0x3B, 0x82, 0xF6), (0x1E, 0x40, 0xAF)          # gradient, top-left -> bottom-right
_BARS = ((0.23, 0.52), (0.43, 0.38), (0.63, 0.24))             # (left x, top y) of each bar, unit square
_BAR_W, _BAR_BOTTOM = 0.15, 0.76


def _inside_square(x, y, r=0.22):
    cx = min(max(x, r), 1 - r)
    cy = min(max(y, r), 1 - r)
    return (x - cx) ** 2 + (y - cy) ** 2 <= r * r


def _inside_bar(x, y):
    return any(bx <= x <= bx + _BAR_W and top <= y <= _BAR_BOTTOM for bx, top in _BARS)


def pixels(size):
    """RGBA bytes, row by row, 4x4 super-sampled for smooth edges."""
    ss = 4
    out = bytearray()
    for py in range(size):
        out.append(0)                                          # PNG filter byte for this row
        for px in range(size):
            bg = white = 0
            for sy in range(ss):
                for sx in range(ss):
                    x, y = (px + (sx + 0.5) / ss) / size, (py + (sy + 0.5) / ss) / size
                    if _inside_square(x, y):
                        bg += 1
                        white += _inside_bar(x, y)
            if not bg:
                out += b'\0\0\0\0'
                continue
            t = (px + py) / (2 * size)
            base = [a + (b - a) * t for a, b in zip(_TOP, _BOTTOM)]
            w = white / bg
            out += bytes(int(c + (255 - c) * w) for c in base) + bytes([int(255 * bg / ss / ss)])
    return bytes(out)


def png(size):
    def chunk(kind, data):
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data) & 0xFFFFFFFF)
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', size, size, 8, 6, 0, 0, 0)) +
            chunk(b'IDAT', zlib.compress(pixels(size), 9)) + chunk(b'IEND', b''))


def write_ico(path=ICO_PATH):
    images = [png(s) for s in SIZES]
    head = struct.pack('<HHH', 0, 1, len(images))
    offset = 6 + 16 * len(images)
    entries = b''
    for s, data in zip(SIZES, images):
        entries += struct.pack('<BBBBHHII', s % 256, s % 256, 0, 0, 1, 32, len(data), offset)
        offset += len(data)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'wb') as f:
        f.write(head + entries + b''.join(images))
    return path


def ensure_ico():
    """Path of assets/app.ico, created on first use; None if it cannot be written."""
    if os.path.isfile(ICO_PATH):
        return ICO_PATH
    try:
        return write_ico()
    except OSError:
        return None


if __name__ == '__main__':
    print(write_ico())
