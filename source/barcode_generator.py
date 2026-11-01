# -*- coding: utf-8 -*-
from pathlib import Path
import sys, zipfile, xml.etree.ElementTree as ET, struct, zlib, binascii

XLSX_NAME = "barcode_generator.xlsx"
OUT_DIR = "out"

NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

L = {"0":"0001101","1":"0011001","2":"0010011","3":"0111101","4":"0100011","5":"0110001","6":"0101111","7":"0111011","8":"0110111","9":"0001011"}
G = {"0":"0100111","1":"0110011","2":"0011011","3":"0100001","4":"0011101","5":"0111001","6":"0000101","7":"0010001","8":"0001001","9":"0010111"}
R = {"0":"1110010","1":"1100110","2":"1101100","3":"1000010","4":"1011100","5":"1001110","6":"1010000","7":"1000100","8":"1001000","9":"1110100"}
PARITY = {"0":"LLLLLL","1":"LLGLGG","2":"LLGGLG","3":"LLGGGL","4":"LGLLGG","5":"LGGLLG","6":"LGGGLL","7":"LGLGLG","8":"LGLGGL","9":"LGGLGL"}

def app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def checksum12(s12):
    nums = [int(c) for c in s12]
    total = sum(nums[::2]) + 3 * sum(nums[1::2])
    return str((10 - total % 10) % 10)

def build_ean(prefix2, brigade3, item7):
    payload = f"{prefix2:02d}{brigade3:03d}{item7:07d}"
    if len(payload) != 12:
        raise ValueError("Internal payload must contain exactly 12 digits")
    return payload + checksum12(payload)

def encode_ean13(ean):
    bits = "101"
    parity = PARITY[ean[0]]
    for digit, p in zip(ean[1:7], parity):
        bits += L[digit] if p == "L" else G[digit]
    bits += "01010"
    for digit in ean[7:]:
        bits += R[digit]
    return bits + "101"

def read_shared_strings(z):
    name = "xl/sharedStrings.xml"
    if name not in z.namelist():
        return []
    root = ET.fromstring(z.read(name))
    return ["".join((t.text or "") for t in si.iter(NS_MAIN + "t")) for si in root.findall(NS_MAIN + "si")]

def sheet_paths(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    relmap = {r.attrib["Id"]: r.attrib["Target"] for r in rels.findall(NS_PKG_REL + "Relationship")}
    out = {}
    for s in wb.find(NS_MAIN + "sheets").findall(NS_MAIN + "sheet"):
        target = relmap[s.attrib[NS_REL + "id"]].replace("\\", "/")
        if target.startswith("/"):
            target = target[1:]
        if not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")
        out[s.attrib["name"]] = target
    return out

def read_sheet(z, path, shared):
    root = ET.fromstring(z.read(path))
    cells = {}
    for c in root.iter(NS_MAIN + "c"):
        ref = c.attrib.get("r", "")
        typ = c.attrib.get("t")
        value = None
        if typ == "inlineStr":
            isel = c.find(NS_MAIN + "is")
            if isel is not None:
                value = "".join((t.text or "") for t in isel.iter(NS_MAIN + "t"))
        else:
            v = c.find(NS_MAIN + "v")
            if v is not None:
                raw = v.text or ""
                if typ == "s":
                    value = shared[int(raw)]
                else:
                    try:
                        value = int(raw) if raw.isdigit() else float(raw)
                    except Exception:
                        value = raw
        if ref:
            cells[ref] = value
    return cells

def read_config(xlsx):
    with zipfile.ZipFile(xlsx, "r") as z:
        shared = read_shared_strings(z)
        paths = sheet_paths(z)
        s = read_sheet(z, paths["SETTINGS"], shared)
        f = read_sheet(z, paths["FORMATS"], shared)

    prefix = int(s.get("B4", 20))
    brigade = int(s.get("B5", 0))
    item = int(s.get("B6", 1))
    count = int(s.get("B7", 10))
    dpi = int(s.get("B8", 600))
    overwrite = int(s.get("B9", 1))
    show_text = int(s.get("B10", 1))

    if not (0 <= prefix <= 99):
        raise ValueError("Internal prefix must be 00..99")
    if not (0 <= brigade <= 999):
        raise ValueError("Shift code must be 000..999")
    if not (0 <= item <= 9999999):
        raise ValueError("Item number must be 0000000..9999999")
    if not (1 <= count <= 50):
        raise ValueError("Quantity must be 1..50")
    if item + count - 1 > 9999999:
        raise ValueError("Item-number range exceeds 9999999")
    if overwrite not in (0, 1):
        raise ValueError("Overwrite must be 0 or 1")
    if show_text not in (0, 1):
        raise ValueError("Digits-below-barcode must be 0 or 1")

    formats = []
    for row in range(2, 500):
        active = f.get(f"A{row}")
        width = f.get(f"B{row}")
        height = f.get(f"C{row}")
        mode = f.get(f"D{row}")
        if active is None and width is None and height is None and mode is None:
            if row > 40:
                break
            continue
        if int(active or 0) == 0:
            continue
        mode = int(mode)
        if mode not in (0, 1, 2):
            raise ValueError(f"FORMATS row {row}: Mode must be 0, 1 or 2")
        formats.append((float(width), float(height), mode))
    return prefix, brigade, item, count, dpi, overwrite, show_text, formats

def mm_to_px(mm, dpi):
    return max(1, int(round(float(mm) / 25.4 * dpi)))

def set_rect(buf, W, H, x1, y1, x2, y2, value=0):
    x1 = max(0, int(x1))
    y1 = max(0, int(y1))
    x2 = min(W - 1, int(x2))
    y2 = min(H - 1, int(y2))
    if x2 < x1 or y2 < y1:
        return
    for y in range(y1, y2 + 1):
        a = y * W + x1
        b = y * W + x2 + 1
        buf[a:b] = bytes([value]) * (x2 - x1 + 1)

def png_chunk(kind, data):
    return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", binascii.crc32(kind + data) & 0xffffffff)

def save_png_gray(path, W, H, buf, dpi):
    rows = bytearray()
    for y in range(H):
        rows.append(0)
        rows.extend(buf[y * W:(y + 1) * W])
    ppm = int(round(dpi / 0.0254))
    data = bytearray(b"\x89PNG\r\n\x1a\n")
    data += png_chunk(b"IHDR", struct.pack(">IIBBBBB", W, H, 8, 0, 0, 0, 0))
    data += png_chunk(b"pHYs", struct.pack(">IIB", ppm, ppm, 1))
    data += png_chunk(b"IDAT", zlib.compress(bytes(rows), 9))
    data += png_chunk(b"IEND", b"")
    path.write_bytes(data)

def clean_num(v):
    return str(int(v)) if float(v).is_integer() else str(v).rstrip("0").rstrip(".").replace(".", "_")

def draw_text_winapi(buf, W, H, text, y, h_area):
    import ctypes
    from ctypes import wintypes

    BI_RGB = 0
    DIB_RGB_COLORS = 0
    DT_CENTER = 1
    DT_VCENTER = 4
    DT_SINGLELINE = 0x20
    DT_NOPREFIX = 0x800
    TRANSPARENT = 1
    DEFAULT_CHARSET = 1
    OUT_OUTLINE_PRECIS = 8
    CLIP_DEFAULT_PRECIS = 0
    ANTIALIASED_QUALITY = 4
    VARIABLE_PITCH = 2
    FF_SWISS = 32
    FW_NORMAL = 400

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", wintypes.DWORD), ("biWidth", ctypes.c_long), ("biHeight", ctypes.c_long),
            ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD),
            ("biCompression", wintypes.DWORD), ("biSizeImage", wintypes.DWORD),
            ("biXPelsPerMeter", ctypes.c_long), ("biYPelsPerMeter", ctypes.c_long),
            ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)
        ]

    class BITMAPINFO(ctypes.Structure):
        _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 3)]

    class RECT(ctypes.Structure):
        _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

    gdi32 = ctypes.windll.gdi32
    user32 = ctypes.windll.user32

    bmi = BITMAPINFO()
    bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth = W
    bmi.bmiHeader.biHeight = -h_area
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32
    bmi.bmiHeader.biCompression = BI_RGB
    bmi.bmiHeader.biSizeImage = W * h_area * 4

    ppv = ctypes.c_void_p()
    hdc = gdi32.CreateCompatibleDC(0)
    if not hdc:
        raise RuntimeError("Windows GDI CreateCompatibleDC failed")
    hbm = gdi32.CreateDIBSection(hdc, ctypes.byref(bmi), DIB_RGB_COLORS, ctypes.byref(ppv), 0, 0)
    if not hbm:
        gdi32.DeleteDC(hdc)
        raise RuntimeError("Windows GDI CreateDIBSection failed")

    oldbmp = gdi32.SelectObject(hdc, hbm)
    rect = RECT(0, 0, W, h_area)
    user32.FillRect(hdc, ctypes.byref(rect), gdi32.GetStockObject(0))

    font_h = max(10, int(h_area * 0.72))
    hfont = gdi32.CreateFontW(
        -font_h, 0, 0, 0, FW_NORMAL, 0, 0, 0,
        DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS,
        ANTIALIASED_QUALITY, VARIABLE_PITCH | FF_SWISS, "Arial"
    )
    oldfont = gdi32.SelectObject(hdc, hfont)
    gdi32.SetBkMode(hdc, TRANSPARENT)
    gdi32.SetTextColor(hdc, 0x000000)
    user32.DrawTextW(hdc, text, -1, ctypes.byref(rect), DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX)

    raw = ctypes.string_at(ppv, W * h_area * 4)
    for yy in range(h_area):
        ro = yy * W * 4
        bo = (y + yy) * W
        for xx in range(W):
            i = ro + xx * 4
            if raw[i] < 220 or raw[i + 1] < 220 or raw[i + 2] < 220:
                buf[bo + xx] = 0

    gdi32.SelectObject(hdc, oldfont)
    gdi32.DeleteObject(hfont)
    gdi32.SelectObject(hdc, oldbmp)
    gdi32.DeleteObject(hbm)
    gdi32.DeleteDC(hdc)

def make_barcode(ean, width_mm, height_mm, dpi, frame, show_text):
    W = mm_to_px(width_mm, dpi)
    H = mm_to_px(height_mm, dpi)
    buf = bytearray([255]) * (W * H)
    bits = encode_ean13(ean)

    module = max(1, W // 113)
    x0 = module * 11 + max(0, W - module * 113) // 2

    top = max(2, int(H * 0.035))
    bottom = max(2, int(H * 0.035))
    text_area = int(H * 0.20) if show_text else 0
    bars_bottom = H - bottom - text_area

    if bars_bottom <= top + 3:
        raise ValueError(f"Format {width_mm}x{height_mm} mm is too short")

    for i, bit in enumerate(bits):
        if bit == "1":
            x1 = x0 + i * module
            set_rect(buf, W, H, x1, top, x1 + module - 1, bars_bottom - 1, 0)

    if show_text:
        guard_extra = min(max(1, text_area // 3), max(1, H - bars_bottom - bottom - 1))
        for a, b in ((0, 2), (45, 49), (92, 94)):
            for i in range(a, b + 1):
                if bits[i] == "1":
                    x1 = x0 + i * module
                    set_rect(buf, W, H, x1, top, x1 + module - 1, bars_bottom + guard_extra, 0)

        text_y = bars_bottom + max(1, guard_extra + 1)
        text_h = max(8, H - bottom - text_y)
        draw_text_winapi(buf, W, H, ean, text_y, text_h)

    if frame:
        frame_px = max(1, int(round(0.30 / 25.4 * dpi)))
        for k in range(frame_px):
            set_rect(buf, W, H, k, k, W - 1 - k, k, 0)
            set_rect(buf, W, H, k, H - 1 - k, W - 1 - k, H - 1 - k, 0)
            set_rect(buf, W, H, k, k, k, H - 1 - k, 0)
            set_rect(buf, W, H, W - 1 - k, k, W - 1 - k, H - 1 - k, 0)

    return W, H, buf

def main():
    here = app_dir()
    xlsx = here / XLSX_NAME
    if not xlsx.exists():
        raise FileNotFoundError(f"{XLSX_NAME} not found next to the program")

    prefix, brigade, item, count, dpi, overwrite, show_text, formats = read_config(xlsx)
    eans = [build_ean(prefix, brigade, item + i) for i in range(count)]

    out_root = here / OUT_DIR
    out_root.mkdir(exist_ok=True)
    made = 0
    skipped = 0

    for width_mm, height_mm, mode in formats:
        wt = clean_num(width_mm)
        ht = clean_num(height_mm)
        folder = out_root / f"{wt}x{ht}_mm"
        folder.mkdir(parents=True, exist_ok=True)

        variants = []
        if mode in (0, 2):
            variants.append(("noframe", False))
        if mode in (1, 2):
            variants.append(("frame", True))

        for ean in eans:
            for suffix, fr in variants:
                path = folder / f"{ean}_{wt}x{ht}_{suffix}.png"
                if path.exists() and overwrite == 0:
                    skipped += 1
                    continue
                W, H, buf = make_barcode(ean, width_mm, height_mm, dpi, fr, show_text)
                save_png_gray(path, W, H, buf, dpi)
                made += 1

    print("DONE")
    print("First EAN :", eans[0])
    print("Last EAN  :", eans[-1])
    print("Count     :", len(eans))
    print("PNG made  :", made)
    print("Skipped   :", skipped)
    print("Output    :", out_root)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("ERROR:", e)
        sys.exit(1)
