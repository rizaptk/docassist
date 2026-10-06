"""Generate DocAssist icons (16/32/80px PNG) tanpa dependensi: biru + huruf DA bitmap 5x7."""
import struct, zlib, os

FONT = {
    "D": ["11110", "10001", "10001", "10001", "10001", "10001", "11110"],
    "A": ["01110", "10001", "10001", "11111", "10001", "10001", "10001"],
}
BG, FG, BORDER = (37, 99, 235), (255, 255, 255), (30, 58, 138)

def png(path, size):
    scale = max(1, size // 24)
    gw, gh = 5 * 2 + 1, 7
    ox = (size - gw * scale) // 2
    oy = (size - gh * scale) // 2
    px = bytearray()
    for y in range(size):
        px.append(0)
        for x in range(size):
            edge = x == 0 or y == 0 or x == size - 1 or y == size - 1
            c = BORDER if edge else BG
            gx, gy = (x - ox) // scale, (y - oy) // scale
            if 0 <= gx < gw and 0 <= gy < gh and scale > 0 and (x - ox) % scale >= 0 and (y - oy) % scale >= 0:
                gl = gx if gx < 5 else gx - 6
                ch = "D" if gx < 5 else "A"
                if 0 <= gl < 5 and FONT[ch][gy][gl] == "1":
                    c = FG
            px += bytes(c)
    raw = zlib.compress(bytes(px), 9)
    def chunk(t, d):
        h = struct.pack(">I4s", len(d), t) + d
        return h + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    out = (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0))
           + chunk(b"IDAT", raw) + chunk(b"IEND", b""))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as f:
        f.write(out)

if __name__ == "__main__":
    base = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets")
    for s in (16, 32, 80):
        png(os.path.join(base, f"icon-{s}.png"), s)
    print("icons OK: assets/icon-16,32,80.png")
