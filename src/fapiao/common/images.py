# -*- coding: utf-8 -*-
"""Excel 内嵌图片：xlsx drawing XML / xls BIFF+Escher 解析、写入、媒体去重。

业务无关，可独立测试。
"""
from __future__ import annotations

import hashlib
import logging
import re
import struct
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor

from ..config import DETAIL_ROW_HEIGHT, INVOICE_DETAIL_START_ROW
from .strings import replace_with_retry

log = logging.getLogger(__name__)

_IMAGE_NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
_EMU_PER_PX = 9525


class RawImage(XLImage):
    """直接把 JPEG/PNG 字节写入 xlsx，避免 openpyxl 依赖 Pillow。"""

    _id = 1
    _path = "/xl/media/image{0}.{1}"

    def __init__(self, data, fmt="jpeg", width=72, height=33):
        self.ref = data
        self.format = fmt
        self.width = width
        self.height = height
        self.anchor = None

    def _data(self):
        return self.ref

    @property
    def path(self):
        return self._path.format(self._id, self.format)


def read_drawing_images(path: Path):
    """从 xlsx 包中读取第一张 drawing 的图片锚点（1 基 Excel 坐标）。"""
    result = []
    try:
        z = zipfile.ZipFile(path)
    except Exception:  # noqa: BLE001  非 zip 包
        return result

    drawing_names = [n for n in z.namelist()
                     if re.match(r"xl/drawings/drawing\d+\.xml$", n)]
    if not drawing_names:
        return result
    drawing_name = sorted(drawing_names)[0]
    rels_name = drawing_name.replace("drawings/", "drawings/_rels/") + ".rels"

    rels = {}
    if rels_name in z.namelist():
        root = ET.fromstring(z.read(rels_name))
        for rel in root:
            rels[rel.get("Id")] = rel.get("Target")

    root = ET.fromstring(z.read(drawing_name))
    anchors = (list(root.findall("xdr:twoCellAnchor", _IMAGE_NS))
               + list(root.findall("xdr:oneCellAnchor", _IMAGE_NS)))
    for anchor in anchors:
        frm = anchor.find("xdr:from", _IMAGE_NS)
        to = anchor.find("xdr:to", _IMAGE_NS)
        if frm is None:
            continue
        blip = anchor.find(".//a:blip", _IMAGE_NS)
        if blip is None:
            continue
        rid = blip.get("{%s}embed" % _IMAGE_NS["r"])
        target = rels.get(rid)
        if not target:
            continue
        if target.startswith("../"):
            target = "xl/" + target[3:]
        elif target.startswith("/"):
            target = target.lstrip("/")
        try:
            data = z.read(target)
        except KeyError:
            continue

        ext = anchor.find(".//a:ext", _IMAGE_NS)
        width = height = None
        if ext is not None:
            width = int(ext.get("cx")) / _EMU_PER_PX
            height = int(ext.get("cy")) / _EMU_PER_PX

        result.append({
            "row": int(frm.find("xdr:row", _IMAGE_NS).text) + 1,
            "col": int(frm.find("xdr:col", _IMAGE_NS).text) + 1,
            "to_row": int(to.find("xdr:row", _IMAGE_NS).text) + 1 if to is not None else None,
            "to_col": int(to.find("xdr:col", _IMAGE_NS).text) + 1 if to is not None else None,
            "data": data,
            "fmt": target.rsplit(".", 1)[-1].lower(),
            "width": width,
            "height": height,
        })
    return result


# ---------- .xls（BIFF8/Escher）内嵌图片 ----------
# xlrd 能读单元格但读不到浮动图片；这里直接解析 Workbook 流里的
# MSODRAWINGGROUP（图片流）和 MSODRAWING（锚点），只依赖 xlrd 的 OLE 读取器。
_XLS_ESCHER_CONTAINERS = {0xF000, 0xF001, 0xF002, 0xF003, 0xF004, 0xF005}
_XLS_IMAGE_SIGNATURES = (
    (b"\xff\xd8\xff", "jpg"),
    (b"\x89PNG\r\n\x1a\n", "png"),
)


def _xls_merged_records(stream: bytes):
    """把 BIFF 记录切出来，并合并 MSODRAWING 后面的 CONTINUE 记录。"""
    pos = 0
    records = []
    while pos + 4 <= len(stream):
        typ, ln = struct.unpack_from("<HH", stream, pos)
        records.append((typ, stream[pos + 4:pos + 4 + ln]))
        pos += 4 + ln

    merged = []
    i = 0
    while i < len(records):
        typ, data = records[i]
        if typ in (0x00EB, 0x00EC):  # MSODRAWINGGROUP / MSODRAWING
            buf = bytearray(data)
            i += 1
            while i < len(records) and records[i][0] == 0x003C:  # CONTINUE
                buf.extend(records[i][1])
                i += 1
            merged.append((typ, bytes(buf)))
            continue
        merged.append((typ, data))
        i += 1
    return merged


def _xls_parse_escher(buf: bytes, start=0, end=None, depth=0, out=None):
    """递归解析 Escher；对"超长容器"按实际剩余字节继续解析。"""
    if out is None:
        out = []
    if end is None:
        end = len(buf)

    pos = start
    while pos + 8 <= end:
        _ver, typ, ln = struct.unpack_from("<HHI", buf, pos)
        data_start = pos + 8
        data_end = data_start + ln
        if typ in _XLS_ESCHER_CONTAINERS:
            out.append((depth, typ, buf[data_start:min(data_end, end)]))
            if data_end > end:
                _xls_parse_escher(buf, data_start, end, depth + 1, out)
                break
            _xls_parse_escher(buf, data_start, data_end, depth + 1, out)
            pos = data_end
        else:
            if data_end > end:
                break
            out.append((depth, typ, buf[data_start:data_end]))
            pos = data_end
    return out


def _xls_fopt_picture_index(data: bytes):
    """从 FOPT（0xF00B）的属性表里取 pib（图片在 BStore 中的 1 基下标）。"""
    pos = 0
    while pos + 6 <= len(data):
        opid, value = struct.unpack_from("<HI", data, pos)
        if (opid & 0x3FFF) == 0x0104:  # pib
            return value
        pos += 6
    return None


def _xls_client_anchor(data: bytes):
    """ClientAnchor 的 9 个 short：colL, dxL, rowL, dyL, colR, dxR, rowR, dyR, flags。"""
    if len(data) < 18:
        return None
    return struct.unpack_from("<9h", data, 0)


def _xls_image_from_blip(blob: bytes):
    """FBSE 里存的是 BLIP；按 JPEG/PNG 文件头截出真正的图片字节。"""
    for sig, fmt in _XLS_IMAGE_SIGNATURES:
        i = blob.find(sig)
        if i < 0:
            continue
        data = blob[i:]
        if fmt == "jpg":
            j = data.rfind(b"\xff\xd9")
            if j >= 0:
                data = data[:j + 2]
        else:
            j = data.rfind(b"IEND")
            if j >= 0:
                data = data[:j + 8]
        return fmt, data
    return None, None


def _xls_image_size(data: bytes, fmt: str):
    """读 JPEG SOF / PNG IHDR 的像素尺寸，用于按比例缩放。"""
    try:
        if fmt == "png":
            return struct.unpack(">II", data[16:24])
        if fmt == "jpg":
            i = 2
            while i + 9 < len(data):
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                              0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                    h, w = struct.unpack(">HH", data[i + 5:i + 9])
                    return w, h
                i += 2 + struct.unpack(">H", data[i + 2:i + 4])[0]
    except Exception:  # noqa: BLE001  尺寸读不出就不缩放
        log.debug("读取图片尺寸失败", exc_info=True)
    return None


def read_xls_images(path: Path):
    """读取 .xls 中的浮动图片（row/col 1 基，size=(宽,高) 像素）。"""
    try:
        from xlrd.compdoc import CompDoc
    except ImportError:
        return []

    try:
        raw = Path(path).read_bytes()
        comp = CompDoc(raw)
        stream = bytes(comp.get_named_stream("Workbook"))
    except Exception:  # noqa: BLE001  不是 BIFF8 / 无 Workbook 流
        log.debug("解析 xls 图片失败：%s", path, exc_info=True)
        return []

    images = []
    shapes = []
    for typ, data in _xls_merged_records(stream):
        if typ == 0x00EB:  # MSODRAWINGGROUP
            for _depth, rec, payload in _xls_parse_escher(data):
                if rec != 0xF007:  # FBSE
                    continue
                name_len = payload[33] if len(payload) > 33 else 0
                blip = payload[34 + name_len:]
                fmt, img = _xls_image_from_blip(blip)
                if img:
                    images.append((fmt, img))
        elif typ == 0x00EC:  # MSODRAWING
            current = {}
            for _depth, rec, payload in _xls_parse_escher(data):
                if rec == 0xF00B:  # FOPT
                    pib = _xls_fopt_picture_index(payload)
                    if pib is not None:
                        current["pib"] = pib
                elif rec == 0xF010:  # ClientAnchor
                    anchor = _xls_client_anchor(payload)
                    if anchor:
                        current["anchor"] = anchor
                elif rec == 0xF004:  # SpContainer
                    if current:
                        shapes.append(current)
                    current = {}
            if current:
                shapes.append(current)

    result = []
    for shape in shapes:
        pib = shape.get("pib")
        anchor = shape.get("anchor")
        if not pib or not anchor or pib > len(images):
            continue
        fmt, data = images[pib - 1]
        result.append({
            "row": anchor[3] + 1,
            "col": anchor[1] + 1,
            "to_row": anchor[7] + 1,
            "to_col": anchor[5] + 1,
            "data": data,
            "fmt": fmt,
            "size": _xls_image_size(data, fmt),
        })
    return result


def read_shiyizhong_images(path: Path):
    """时易中图片：.xls 解析 BIFF，.xlsx 沿用 drawing XML。"""
    suffix = Path(path).suffix.lower()
    if suffix == ".xls":
        return read_xls_images(path)
    if suffix in (".xlsx", ".xlsm"):
        out = []
        for item in read_drawing_images(path):
            w, h = item.get("width"), item.get("height")
            out.append({
                "row": item["row"],
                "col": item["col"],
                "to_row": item.get("to_row"),
                "to_col": item.get("to_col"),
                "data": item["data"],
                "fmt": "png" if item.get("fmt") == "png" else "jpg",
                "size": (w, h) if w and h else None,
            })
        return out
    return []


def insert_shiyizhong_images(ws, input_path: Path, detail_start: int, data_count: int) -> int:
    """把源发票 Q 列的浮动图片写进目标 Z{明细行}，等比缩放并居中。"""
    if data_count <= 0:
        return 0

    # 目标单元格像素尺寸：行高已知，列宽尽量从模板读取
    row_h_px = int(DETAIL_ROW_HEIGHT * 96 / 72)  # 54 磅 → 72 px
    try:
        col_w_chars = ws.column_dimensions["Z"].width
    except (KeyError, AttributeError):
        col_w_chars = None
    if not col_w_chars:
        col_w_chars = 8.43
    col_w_px = int(col_w_chars * 7 + 5)  # Calibri 11：1 字符宽 ≈ 7px + 5px padding

    images = {}
    for info in read_shiyizhong_images(input_path):
        if info["col"] != 17:  # Q
            continue
        images.setdefault(info["row"], []).append(info)

    used = 0
    for i in range(data_count):
        source_row = detail_start + 1 + i
        target_row = INVOICE_DETAIL_START_ROW + i
        candidates = images.get(source_row)
        if not candidates:
            continue
        info = candidates[0]

        size = info.get("size")
        if size and size[0] and size[1]:
            w, h = size
            scale = min((col_w_px - 4) / w, (row_h_px - 4) / h, 1.0)
            img_w = max(1, int(w * scale))
            img_h = max(1, int(h * scale))
        else:
            img_w = max(1, col_w_px - 4)
            img_h = max(1, row_h_px - 4)

        off_x = max(0, (col_w_px - img_w) // 2)
        off_y = max(0, (row_h_px - img_h) // 2)

        ws.cell(target_row, 26).value = None
        img = RawImage(info["data"], "png" if info["fmt"] == "png" else "jpeg",
                       img_w, img_h)
        # from / to 都落在 Z 列同一格（0 基 col=25），靠偏移量控制位置与大小
        img.anchor = TwoCellAnchor(
            _from=AnchorMarker(col=25, row=target_row - 1,
                               colOff=off_x * _EMU_PER_PX,
                               rowOff=off_y * _EMU_PER_PX),
            to=AnchorMarker(col=25, row=target_row - 1,
                            colOff=(off_x + img_w) * _EMU_PER_PX,
                            rowOff=(off_y + img_h) * _EMU_PER_PX),
        )
        ws.add_image(img)
        used += 1
    return used


def dedupe_media(path: Path) -> None:
    """合并内容相同的图片，避免同一张图片被重复写入多次。"""
    path = Path(path)
    tmp = None
    try:
        with zipfile.ZipFile(path) as zin:
            names = zin.namelist()
            media = [n for n in names if n.startswith("xl/media/")]
            if not media:
                return

            hash_to_name = {}
            target_map = {}
            for n in media:
                digest = hashlib.sha1(zin.read(n)).hexdigest()
                hash_to_name.setdefault(digest, n)
                target_map[n] = hash_to_name[digest]
            if len(set(target_map.values())) == len(media):
                return

            rels_name = "xl/drawings/_rels/drawing1.xml.rels"
            if rels_name not in names:
                return

            ET.register_namespace(
                "", "http://schemas.openxmlformats.org/package/2006/relationships"
            )
            root = ET.fromstring(zin.read(rels_name))
            for rel in root:
                target = rel.get("Target", "")
                if target.startswith("/xl/media/"):
                    full = target.lstrip("/")
                    canon = target_map.get(full, full)
                    rel.set("Target", "/" + canon)
                elif target.startswith("../media/"):
                    full = "xl/media/" + target.rsplit("/", 1)[-1]
                    canon = target_map.get(full, full)
                    rel.set("Target", "../media/" + canon.rsplit("/", 1)[-1])
            new_rels = ET.tostring(root, encoding="utf-8", xml_declaration=True)

            tmp = path.with_name(path.stem + ".dedupe" + path.suffix)
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for info in zin.infolist():
                    name = info.filename
                    if name in media and target_map[name] != name:
                        continue
                    if name == rels_name:
                        zout.writestr(info, new_rels)
                    else:
                        zout.writestr(info, zin.read(name))

        if tmp is not None:
            replace_with_retry(tmp, path)
    except OSError as e:
        # 去重只是减小体积，失败不能影响正常输出
        log.warning("%s: 图片去重失败，已保留未去重文件（%s）", path.name, e)
        if tmp is not None and tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
