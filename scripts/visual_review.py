#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""企业官网资料包 · 视觉核对产物生成

为每家企业生成"看图材料"和空结论载体，供 Codex（图像查看）或人工核对：

    <run>/review/视觉核对图/<企业>/0.总览.png      四类速览
    <run>/review/视觉核对图/<企业>/<分类>.png       分类拼版（单页≤12 张；超过则 <分类>_p1.png、_p2.png…）
    <run>/review/图片核对表.xlsx                    带缩略图 + 结论下拉
    <run>/视觉核对.json                             结论载体，供 gates.py 读取

本脚本只负责出材料和占位，不代替人/模型判断图片内容。
Codex 执行时用图像查看工具打开拼版，把结论写回 视觉核对.json，再跑 gates.py。
"""
import argparse
import datetime
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

DIRS = {
    "factory": "1.企业工厂图",
    "product": "2.企业产品图",
    "logo": "3.企业logo",
    "cert": "4.资质证书",
}
LABELS = {
    "1.企业工厂图": "工厂/车间/厂房/设备",
    "2.企业产品图": "产品/商品/样品",
    "3.企业logo": "品牌 logo/标志",
    "4.资质证书": "证书/资质/荣誉/认证",
}
FONT_CANDIDATES = (
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\msyhbd.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttc",
    r"C:\Windows\Fonts\arial.ttf",
)


def read_json(path):
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def write_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def sha256_file(path):
    """图片/拼版内容指纹；文件不存在或读取失败时返回空串。"""
    try:
        h = hashlib.sha256()
        with open(path, "rb") as fh:
            for chunk in iter(lambda: fh.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()
    except OSError:
        return ""



ANCHOR_STOP = {
    "公司", "企业", "产品", "主要", "业务", "简介", "工厂", "车间",
    "图片", "照片", "其他", "更多", "首页", "关于", "联系", "详情",
}


def clean_anchor(text):
    s = re.sub(r"\s+", " ", str(text or "")).strip()
    s = re.sub(r"^[\-—·•\d]+[\.、,，:：]?\s*", "", s)
    return s.strip(" .、,，;；")


def clean_inline_anchor(text):
    s = re.sub(r"[（(【\[].*?[）)】\]]", " ", str(text or ""))
    s = re.sub(r"\s+", " ", s).strip()
    return clean_anchor(s)


def anchor_tokens(archive):
    """从企业档案抽取文字锚点：已证实的名称/产品/型号/品类等词。

    用于看图判断时的参照——图片内容要能挂到这些词上才算贴合本企业，
    挂不上就不该直接判符合。锚点只取结构化字段（名称/产品/品类），
    不拆简介全文，避免把上游噪声带成假锚点。锚点是判断依据，不是自动结论。
    """
    name = clean_anchor(archive.get("名称"))
    raw = []
    for key in ("核心产品", "主要业务", "产业", "规模"):
        raw.append(archive.get(key))
    for prod in archive.get("products") or []:
        raw.append(clean_inline_anchor(prod))
    details = archive.get("product_details")
    if isinstance(details, dict):
        for key in details:
            raw.append(clean_inline_anchor(key))
    anchors, seen = [], set()
    if name:
        anchors.append(name)
        seen.add(name)
    for item in raw:
        s = clean_inline_anchor(item)
        if not s or len(s) < 2 or len(s) > 24 or s in ANCHOR_STOP or s in seen:
            continue
        seen.add(s)
        anchors.append(s)
        if len(anchors) >= 24:
            break
    return anchors


def match_anchors(*texts, anchors):
    """在图名/alt/URL 里找出能挂靠的文字锚点，作为判断线索。"""
    blob = " ".join(str(t or "") for t in texts).lower()
    hit = []
    for a in anchors or []:
        key = a.lower().strip()
        if len(key) >= 2 and key in blob and a not in hit:
            hit.append(a)
    return hit


def company_dir_name(name):
    """与 local_pipeline.safe_filename 保持一致。"""
    s = re.sub(r"\s+", " ", str(name or "")).strip()
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", s)
    s = s.strip(" .")
    return s[:120] or "company"


def font(size):
    from PIL import ImageFont
    for cand in FONT_CANDIDATES:
        if os.path.isfile(cand):
            try:
                return ImageFont.truetype(cand, size)
            except Exception:
                continue
    return ImageFont.load_default()


def wrap(draw, text, fnt, max_w, max_lines=2):
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    if not text:
        return []
    lines, cur, truncated = [], "", False
    for ch in text:
        if draw.textlength(cur + ch, font=fnt) <= max_w:
            cur += ch
        else:
            lines.append(cur)
            cur = ch
            if len(lines) >= max_lines:
                truncated = True
                break
    if not truncated and cur and len(lines) < max_lines:
        lines.append(cur)
    if truncated and lines:
        lines[-1] = lines[-1][:-1] + "\u2026"
    return lines


def load_thumb(path, box):
    from PIL import Image, ImageDraw
    try:
        im = Image.open(str(path))
        im = im.convert("RGBA")
        # 透明 logo 常见为白色或蓝色单色素材，直接叠白底会“看不见”。
        # 用浅色棋盘底展示透明通道，同时不影响无透明通道的普通图片。
        tile = 12
        bg = Image.new("RGB", im.size, (245, 246, 248))
        d = ImageDraw.Draw(bg)
        for yy in range(0, im.height, tile):
            for xx in range(0, im.width, tile):
                if ((xx // tile) + (yy // tile)) % 2:
                    d.rectangle([xx, yy, min(xx + tile - 1, im.width - 1),
                                 min(yy + tile - 1, im.height - 1)],
                                fill=(221, 225, 232))
        bg = bg.convert("RGBA")
        im = Image.alpha_composite(bg, im).convert("RGB")
        im.thumbnail(box, Image.LANCZOS)
        return im
    except Exception:
        return None


def placeholder(box, text, fnt):
    from PIL import Image, ImageDraw
    im = Image.new("RGB", box, (240, 241, 245))
    d = ImageDraw.Draw(im)
    d.rectangle([0, 0, box[0] - 1, box[1] - 1], outline=(200, 204, 214), width=2)
    lines = wrap(d, text, fnt, box[0] - 30, max_lines=3)
    y = box[1] // 2 - 10 * len(lines)
    for line in lines:
        w = d.textlength(line, font=fnt)
        d.text(((box[0] - w) / 2, y), line, font=fnt, fill=(110, 114, 124))
        y += 22
    d.text((12, 12), "无法预览", font=fnt, fill=(170, 60, 60))
    return im


MAX_CELLS_PER_PAGE = 12
MAX_PAGE_BYTES = 1_000_000


def page_index(path):
    """拼版页码，用于把 _p2 排在 _p10 前面；单页文件返回 1。"""
    m = re.search(r"_p(\d+)$", Path(path).stem)
    return int(m.group(1)) if m else 1


def seq_range(page):
    """该页图片在所属分类内的序号范围，与图片核对表的「序号」列一致。"""
    if not page:
        return 0, 0
    return (page[0].get("idx", 1), page[-1].get("idx", len(page)))


def page_note(page, index, total, grand_total):
    """分页时的页眉说明；单页不写，避免和总览混淆。"""
    if total <= 1:
        return ""
    first, last = seq_range(page)
    span = str(first) if first == last else f"{first}-{last}"
    return f"第 {index + 1} / {total} 页 · 序号 {span} · 共 {grand_total} 张"


def _render_montage_page(records, out_path, title, subtitle, page_note=""):
    from PIL import Image, ImageDraw
    cols, cell_w, cell_h = 4, 380, 430
    thumb_w, thumb_h, head = 340, 300, 116
    rows = max(1, (len(records) + cols - 1) // cols)
    width = cols * cell_w + 20
    height = head + rows * cell_h + 20
    canvas = Image.new("RGB", (width, height), (255, 255, 255))
    d = ImageDraw.Draw(canvas)
    tf, cf, sf = font(30), font(17), font(15)
    d.rectangle([0, 0, width, head], fill=(24, 64, 120))
    d.text((20, 16), title, font=tf, fill=(255, 255, 255))
    d.text((20, 68), subtitle, font=sf, fill=(214, 227, 245))
    if page_note:
        d.text((20, 90), page_note, font=sf, fill=(184, 205, 232))
    for i, rec in enumerate(records):
        row, col = divmod(i, cols)
        x, y = 10 + col * cell_w, head + row * cell_h
        d.rectangle([x + 6, y + 6, x + cell_w - 6, y + cell_h - 6], outline=(208, 213, 224), width=2)
        im = load_thumb(rec["path"], (thumb_w, thumb_h))
        if im is None:
            im = placeholder((thumb_w, thumb_h), rec["file"], cf)
        px = x + (cell_w - im.width) // 2
        py = y + 18 + (thumb_h - im.height) // 2
        canvas.paste(im, (px, py))
        d.rectangle([px, py, px + im.width, py + im.height], outline=(178, 183, 195), width=1)
        # 角标与标题都用「该分类内的序号」，与图片核对表的「序号」列一一对应，
        # 便于先用核对表初筛、再按序号只打开相关拼版页。
        no = rec.get("idx", i + 1)
        d.rectangle([x + 16, y + 16, x + 62, y + 50], fill=(204, 42, 42))
        d.text((x + 30, y + 21), str(no), font=cf, fill=(255, 255, 255))
        cap_y = y + 18 + thumb_h + 8
        d.text((x + 18, cap_y), f"{no:02d} {rec['file']}", font=cf, fill=(20, 20, 20))
        cap_y += 24
        for line in wrap(d, rec.get("alt") or rec.get("url", ""), sf, cell_w - 40, 2):
            d.text((x + 18, cap_y), line, font=sf, fill=(92, 97, 108))
            cap_y += 20
    out_path.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(str(out_path))
    canvas.close()
    return out_path


def build_montage(records, out_path, title, subtitle,
                  max_cells=MAX_CELLS_PER_PAGE, max_bytes=MAX_PAGE_BYTES):
    """生成拼版图；单页超过 max_cells 张或 max_bytes 字节时自动分页。

    返回每页信息列表（``path`` / ``first`` / ``last`` / ``count``）；
    ``first``、``last`` 是该页图片在本分类内的序号范围，与图片核对表的
    「序号」列一致。单页沿用 out_path；分页时写成
    ``<名称>_p1.png``、``<名称>_p2.png``…… 顺序与 records 一致。
    分页是为了让 Codex 逐页看图时不会把超大拼版读进会话（历史事故：
    146 张的产品拼版单张 13.6MB）。常规情况每页只渲染一次；只有某页
    仍超 max_bytes 时才整体重排重渲染。
    """
    records = list(records or [])
    if not records:
        return []
    out_path = Path(out_path)
    for stale in out_path.parent.glob(f"{out_path.stem}*.png"):
        stale.unlink()
    tmp_dir = out_path.parent / f"{out_path.stem}.paged"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    pages = [records[i:i + max_cells] for i in range(0, len(records), max_cells)]
    try:
        while True:
            for old in tmp_dir.glob(f"*{out_path.suffix}"):
                old.unlink()
            total = len(pages)
            page_files = []
            for i, page in enumerate(pages):
                tmp_file = tmp_dir / f"{i:03d}{out_path.suffix}"
                _render_montage_page(
                    page, tmp_file, title, subtitle,
                    page_note(page, i, total, len(records)))
                page_files.append(tmp_file)
            oversized = [i for i, tmp_file in enumerate(page_files)
                         if tmp_file.stat().st_size > max_bytes and len(pages[i]) > 1]
            if not oversized:
                break
            for i in reversed(oversized):
                page = pages[i]
                half = len(page) // 2
                pages[i:i + 1] = [page[:half], page[half:]]

        total = len(pages)
        result = []
        for i, tmp_file in enumerate(page_files):
            dst = out_path if total == 1 else out_path.with_name(
                f"{out_path.stem}_p{i + 1}{out_path.suffix}")
            if dst.exists():
                dst.unlink()
            shutil.move(str(tmp_file), str(dst))
            first, last = seq_range(pages[i])
            result.append({"path": dst, "first": first, "last": last,
                           "count": len(pages[i])})
        return result
    finally:
        shutil.rmtree(str(tmp_dir), ignore_errors=True)


def company_records(deliverable, name, archive):
    out = {}
    home = deliverable / company_dir_name(name)
    company_anchors = anchor_tokens(archive)
    for cat, folder in DIRS.items():
        items = []
        for i, it in enumerate(archive.get(cat) or [], 1):
            filename = str(it.get("file", ""))
            path = home / folder / filename
            alt = str(it.get("alt", ""))
            items.append({
                "idx": i,
                "folder": folder,
                "file": filename,
                "path": path,
                "alt": alt,
                "url": str(it.get("url", "")),
                "from": str(it.get("from", "")),
                "wh": str(it.get("wh", "")),
                "source": str(it.get("source", "")),
                "trust": "用户资料" if str(it.get("source", "")) == "用户资料" else "官网",
                "sha256": sha256_file(path),
                "anchors": company_anchors,
                "candidate_anchors": match_anchors(alt, filename, str(it.get("url", "")), anchors=company_anchors),
            })
        out[folder] = items
    return out


def make_review_xlsx(path, records_by_cat):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = Workbook()
    ws = wb.active
    ws.title = "图片核对"
    headers = ["分类", "序号", "文件", "缩略图", "尺寸", "图片URL", "来源页面", "alt",
               "锚点线索", "结论", "说明", "SHA256", "信任层级"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)

    tmp = Path(tempfile.mkdtemp(prefix="review_thumb_"))
    row = 2
    try:
        for folder, items in records_by_cat.items():
            for rec in items:
                ws.cell(row=row, column=1, value=folder)
                ws.cell(row=row, column=2, value=rec["idx"])
                ws.cell(row=row, column=3, value=rec["file"])
                ws.cell(row=row, column=5, value=rec["wh"])
                ws.cell(row=row, column=6, value=rec["url"])
                ws.cell(row=row, column=7, value=rec["from"])
                ws.cell(row=row, column=8, value=rec["alt"])
                ws.cell(row=row, column=9, value="、".join(rec.get("candidate_anchors") or []))
                ws.cell(row=row, column=12, value=rec["sha256"])
                ws.cell(row=row, column=13, value=rec.get("trust", "官网"))
                ws.row_dimensions[row].height = 92
                im = load_thumb(rec["path"], (120, 120))
                if im is not None:
                    thumb = tmp / f"t{row}.png"
                    im.save(str(thumb))
                    from openpyxl.drawing.image import Image as XLImage
                    xl = XLImage(str(thumb))
                    xl.width, xl.height = 116, 116
                    ws.add_image(xl, f"D{row}")
                else:
                    ws.cell(row=row, column=4, value="无法预览")
                row += 1
        last = max(2, row - 1)
        dv = DataValidation(type="list", formula1='"符合,不符,待核对"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"J2:J{last}")
        widths = {"A": 16, "B": 6, "C": 26, "D": 18, "E": 10, "F": 40,
                  "G": 40, "H": 24, "I": 26, "J": 10, "K": 30, "L": 68, "M": 12}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"

        ws2 = wb.create_sheet("文档与产品")
        ws2.append(["对象", "位置", "核对要点", "结论", "说明"])
        for c in range(1, 6):
            ws2.cell(row=1, column=c).font = Font(bold=True)
        rows = [
            ["企业介绍 docx", "5.企业介绍/*.docx", "中文简介是否为本企业事实、无导航套话；中英段是否对应", "", ""],
            ["产品清单 xlsx", "产品清单.xlsx", "每行是否为真实产品；中文名、英文名必填；产品详情非空时中英对应，来源无详情时留空", "", ""],
        ]
        for r in rows:
            ws2.append(r)
        dv2 = DataValidation(type="list", formula1='"符合,不符,待核对"', allow_blank=True)
        ws2.add_data_validation(dv2)
        dv2.add("D2:D3")
        for col, w in {"A": 16, "B": 22, "C": 46, "D": 10, "E": 34}.items():
            ws2.column_dimensions[col].width = w
        ws2.freeze_panes = "A2"
        for row_cells in ws.iter_rows():
            for cell in row_cells:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(str(path))
    finally:
        shutil.rmtree(str(tmp), ignore_errors=True)


def build_template(companies_records, existing, montage_hashes=None):
    out = {}
    for name, records_by_cat in companies_records.items():
        prev = existing.get(name) if isinstance(existing.get(name), dict) else {}
        prev_imgs = prev.get("图片") if isinstance(prev.get("图片"), dict) else {}
        imgs = {}
        for folder, items in records_by_cat.items():
            for rec in items:
                key = f"{folder}/{rec['file']}"
                pi = prev_imgs.get(key) if isinstance(prev_imgs.get(key), dict) else {}
                current_sha = rec.get("sha256", "")
                keep = bool(current_sha and pi.get("sha256") == current_sha)
                imgs[key] = {
                    "结论": pi.get("结论", "") if keep else "",
                    "说明": pi.get("说明", "") if keep else "",
                    "依据": pi.get("依据", "") if keep else "",
                    "trust": rec.get("trust", "官网"),
                    "必核": rec.get("trust", "官网") != "用户资料",
                    "候选锚点": rec.get("candidate_anchors") or [],
                    "sha256": current_sha,
                    "reviewed_at": pi.get("reviewed_at", "") if keep else "",
                }
        doc = prev.get("文档") if isinstance(prev.get("文档"), dict) else {}
        prod = prev.get("产品清单") if isinstance(prev.get("产品清单"), dict) else {}
        out[name] = {
            "图片": imgs,
            "文档": {"结论": doc.get("结论", ""), "说明": doc.get("说明", "")},
            "产品清单": {"结论": prod.get("结论", ""), "说明": prod.get("说明", "")},
        }
    prev_info = existing.get("_核对信息") if isinstance(existing.get("_核对信息"), dict) else {}
    anchor_map = {}
    for name, records_by_cat in companies_records.items():
        for items in records_by_cat.values():
            for rec in items:
                if rec.get("anchors"):
                    anchor_map[name] = rec["anchors"]
                    break
            if name in anchor_map:
                break
    out["_核对信息"] = {
        "核对人": prev_info.get("核对人", ""),
        "核对时间": prev_info.get("核对时间", ""),
        "写入条数": prev_info.get("写入条数", 0),
        "拼版哈希": montage_hashes or {},
        "文字锚点": anchor_map,
    }
    return out


def apply_verdicts(json_path, verdicts_path, reviewer="Codex"):
    """把 Codex（或人工）看图后给出的结论合并进 视觉核对.json。

    verdicts.json 既可写成 {"企业名": {"图片": {...}}}，
    也可写成 {"核对": {...}}。图片条目支持简写 "符合"。
    """
    if not verdicts_path.is_file():
        raise SystemExit(f"找不到核对结论文件：{verdicts_path}")
    data = read_json(json_path) if json_path.is_file() else {}
    payload = read_json(verdicts_path)
    if isinstance(payload, dict) and isinstance(payload.get("核对"), dict):
        payload = payload["核对"]
    applied = 0
    now = datetime.datetime.now().isoformat(timespec="seconds")
    errors = []
    sections = ("图片",)
    for name, entry in payload.items():
        if name.startswith("_") or not isinstance(entry, dict):
            continue
        dst = data.setdefault(name, {})
        for section in sections:
            block = entry.get(section)
            if not isinstance(block, dict):
                continue
            target = dst.setdefault(section, {})
            for key, val in block.items():
                if isinstance(val, str):
                    val = {"结论": val}
                if isinstance(val, dict):
                    clean = {k: v for k, v in val.items() if k in ("结论", "说明", "依据")}
                    if clean:
                        item = target.get(key)
                        if not isinstance(item, dict):
                            errors.append(f"{name} {key}: 当前图片清单中不存在该键")
                            continue
                        expected_sha = str(item.get("sha256") or "")
                        if not expected_sha:
                            errors.append(f"{name} {key}: 当前图片缺少 sha256，先重跑 visual_review.py")
                            continue
                        if val.get("sha256") and str(val.get("sha256")) != expected_sha:
                            errors.append(f"{name} {key}: 提交的 sha256 与当前图片不一致")
                            continue
                        item.update(clean)
                        item["sha256"] = expected_sha
                        item["reviewed_at"] = now
                        applied += 1
        for field in ("文档", "产品清单"):
            val = entry.get(field)
            if isinstance(val, str):
                val = {"结论": val}
            if isinstance(val, dict):
                clean = {k: v for k, v in val.items() if k in ("结论", "说明")}
                if clean:
                    item = dst.setdefault(field, {})
                    item.update(clean)
                    item["reviewed_at"] = now
                    applied += 1
    if errors:
        raise SystemExit("视觉结论未写入，请修正后重试：\n- " + "\n- ".join(errors[:20]))
    prev_info = data.get("_核对信息") if isinstance(data.get("_核对信息"), dict) else {}
    data["_核对信息"] = {
        "核对人": reviewer,
        "核对时间": now,
        "写入条数": applied,
        "拼版哈希": prev_info.get("拼版哈希", {}),
        "文字锚点": prev_info.get("文字锚点", {}),
    }
    write_json(json_path, data)
    return applied


def write_review_guide(path, records_map, review, json_path, gates_cmd, pages_map=None):
    """生成给 Codex 执行的看图核对指引（作为 agent 的必做清单）。"""
    pages_map = pages_map or {}
    guard = Path(__file__).with_name("session_guard.py")
    lines = [
        "# 视觉核对指引（Codex 必做）",
        "",
        "本步骤不能交给用户代做，也不能把结论留成“待核对”后直接发布。",
        "Codex 必须自己看图判断图片是否属于该企业、是否属于该分类、是否为清晰可用素材。",
        "必须在短线程内完成：每批 ≤10 家，一个会话只做「看图 → 回写结论 → 跑门禁 → 发布」，做完即止。",
        "开工前先查本会话 rollout 体积，超过 20MB 就另开会话再继续（拼版图会以 base64 留在会话历史里）：",
        "",
        "```powershell",
        f'python "{guard}"',
        "```",
        "",
        "## 大批量初筛（图片多时优先）",
        "",
        "- 图片多的企业先用同目录的 `图片核对表.xlsx` 初筛：每行一张图，对着 分类/序号/尺寸/图片URL/来源页面/alt/锚点线索 找可疑项。",
        "- 用下面的「序号范围」定位可疑序号所在的拼版页，只打开这些页做放大确认；每张图片都必须有独立结论，不能用类别结论代替。",
        "- 初筛只是先导，不能代替结论：初筛点名的行必须逐张给出结论，`视觉核对.json` 里不能留下未填的「待核对」。",
        "",
        "## 每家企业",
        "",
    ]
    for name, records_by_cat in records_map.items():
        base = review / "视觉核对图" / company_dir_name(name)
        total = sum(len(v) for v in records_by_cat.values())
        lines.append(f"### {name}（{total} 张）")
        for page_path in sorted(base.glob("0.总览*.png"), key=page_index):
            lines.append(f"- 总览：`{page_path}`")
        for folder in sorted(records_by_cat):
            items = records_by_cat[folder]
            if not items:
                continue
            pages = pages_map.get((name, folder)) or []
            user_n = sum(1 for rec in items if rec.get("trust") == "用户资料")
            site_n = len(items) - user_n
            trust_note = f"官网必核 {site_n} 张" + (
                f"；用户资料直通 {user_n} 张" if user_n else "")
            if pages:
                lines.append(f"- {folder}（{len(items)} 张，{len(pages)} 页；{trust_note}）：")
                for page in pages:
                    span = (str(page["first"]) if page["first"] == page["last"]
                            else f"{page['first']}-{page['last']}")
                    lines.append(f"    - 序号 {span}（{page['count']} 张）：`{page['path']}`")
            else:
                lines.append(f"- {folder}（{len(items)} 张；{trust_note}）：`{base / (folder + '.png')}`")
        lines.append("")
    lines += [
        "## 看图范围",
        "",
        "- 分类拼版单页最多 12 张、≤1MB；图片多时写成 `<分类>_p1.png`、`<分类>_p2.png`……按序号分页。初筛点名的页必须逐页打开；每张图片都必须逐张给出结论，不支持类别级继承。",
        "- 分级信任：官网抓取图（核对表「信任层级=官网」）必须逐张看图并回写结论，未回写会阻断发布。",
        "- 分级信任：用户资料图（核对表「信任层级=用户资料」）按资料优先直通，不要求逐张结论；但一旦标「不符」仍会阻断，需先处理。",
        "- 外域/聚合站图片由 `image_provenance` 门禁判红，不进入强制核对面；处理方式是删除该图并重抓，不要靠视觉结论放行。",
        "- `视觉核对.json` 中每张图片都带 `sha256`；回写时由 `--apply` 自动绑定当前文件指纹和 `reviewed_at`，不要手改。图片变化后旧结论会自动作废。",
        "",
        "## 判断口径",
        "",
        "三档结论：`符合` / `不符` / `待核对`。只有 `符合` 才能过门禁；`待核对` 与留空一样阻断发布，不是放行。",
        "判断时先看核对表的「锚点线索」：图片内容要能挂到本企业已证实的名称/产品/型号/品类上，才判「符合」；",
        "挂不上任何锚点又无法确认归属的，判「待核对」，不要硬判「符合」。",
        "",
        "- 企业logo：是否为该企业标识；不是则「不符」。",
        "- 企业工厂图：是否为厂区、车间、产线、办公/园区实景；不是则「不符」。",
        "- 企业产品图：是否为该企业产品；错图、宣传海报、无关配图、图库素材则「不符」。",
        "- 资质证书：是否为该企业资质/证书；模糊到不可辨认或有其他企业名称则「不符」。",
        "- 文档与产品清单：docx 是否中英双语且无乱码；产品清单行是否为真实产品、图片与产品是否对应。",
        "- 同一张图跨类别重复且语义不符，判「不符」并在说明里写明。",
        "- 拿不准时用「待核对」，并在说明里写清卡在哪一步（分辨率、归属、来源）；不要为了过门禁硬填「符合」。",
        "",
        "## 回写格式",
        "",
        "把结论写成 verdicts.json（简写也可）；`依据` 写清命中哪个锚点或来源，便于事后复核：",
        "",
        "```json",
        "{",
        '  "企业全称": {',
        '    "图片": {"3.企业logo/01_xxx.png": {"结论": "符合", "说明": "", "依据": "锚点：企业名"}},',
        '    "图片待核对示例": {"2.企业产品图/02_xxx.jpg": {"结论": "待核对", "说明": "图库素材疑似，无法确认归属"}},',
        '    "文档": {"结论": "符合"},',
        '    "产品清单": {"结论": "符合"}',
        "  }",
        "}",
        "```",
        "",
        "## 应用结论",
        "",
        "```powershell",
        f'python "{Path(__file__).resolve()}" --json "{json_path}" --apply "<verdicts.json>"',
        gates_cmd,
        "```",
        "",
        "任一条「不符」都必须先修数据或补抓再重跑；「待核对」同样阻断发布，拿不准就如实标注并说明原因，不要用「符合」蒙混过关。",
        "",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description="生成企业官网资料包的视觉核对材料")
    ap.add_argument("--deliverable", default="", help="build/<run_id>/deliverable")
    ap.add_argument("--raw", default="", help="build/<run_id>/raw")
    ap.add_argument("--out", default="", help="review 目录，默认 <run>/review")
    ap.add_argument("--json", dest="json_path", default="", help="视觉核对.json，默认 <run>/视觉核对.json")
    ap.add_argument("--companies", default="", help="只处理指定企业，逗号分隔")
    ap.add_argument("--apply", default="", help="把 Codex/人工看图结论 verdicts.json 合并进 视觉核对.json")
    ap.add_argument("--reviewer", default="Codex", help="核对人名称，默认 Codex")
    a = ap.parse_args()

    if a.apply:
        if not a.json_path:
            raise SystemExit("--apply 需要同时给 --json <视觉核对.json>")
        json_path = Path(a.json_path).expanduser().resolve()
        n = apply_verdicts(json_path, Path(a.apply).expanduser().resolve(), a.reviewer)
        print(f"已写入视觉核对结论 {n} 条：{json_path}")
        return 0

    if not a.deliverable or not a.raw:
        raise SystemExit("生成核对材料需要 --deliverable 和 --raw；只回写结论请用 --apply")
    deliverable = Path(a.deliverable).expanduser().resolve()
    raw = Path(a.raw).expanduser().resolve()
    if not deliverable.is_dir():
        raise SystemExit(f"交付目录不存在：{deliverable}")
    if not raw.is_dir():
        raise SystemExit(f"档案目录不存在：{raw}")
    stage = deliverable.parent
    review = Path(a.out).expanduser().resolve() if a.out else stage / "review"
    json_path = Path(a.json_path).expanduser().resolve() if a.json_path else stage / "视觉核对.json"
    only = {s.strip() for s in a.companies.split(",") if s.strip()}

    archives = {}
    for fp in sorted(raw.glob("*.json")):
        try:
            d = read_json(fp)
        except Exception as exc:
            print(f"  ! 跳过无法解析的档案 {fp}: {exc}", file=sys.stderr)
            continue
        if d.get("名称") and (not only or d["名称"] in only):
            archives[d["名称"]] = d
    if not archives:
        raise SystemExit("没有可处理的企业档案")

    existing = read_json(json_path) if json_path.is_file() else {}
    records_map, counts, pages_map = {}, {}, {}
    for name, archive in archives.items():
        records_by_cat = company_records(deliverable, name, archive)
        records_map[name] = records_by_cat
        base = review / "视觉核对图" / company_dir_name(name)
        flat = []
        for folder, items in records_by_cat.items():
            counts[(name, folder)] = len(items)
            flat += items
            if items:
                hint = " / ".join((items[0].get("anchors") or [])[:6])
                subtitle = (f"{LABELS.get(folder, '')} · {len(items)} 张 · "
                            f"Codex 逐张确认是否属于本企业与本分类")
                if hint:
                    subtitle += f" · 锚点参照：{hint}"
                pages_map[(name, folder)] = build_montage(
                    items, base / f"{folder}.png",
                    f"{name} · {folder}",
                    subtitle,
                )
        if flat:
            picks = []
            for folder, items in records_by_cat.items():
                for rec in items[:3]:
                    r2 = dict(rec)
                    r2["alt"] = f"[{folder}] " + (rec.get("alt") or rec.get("url", ""))
                    picks.append(r2)
            build_montage(
                picks, base / "0.总览.png",
                f"{name} · 图片总览",
                f"共 {len(flat)} 张 · 每类最多 3 张速览 · 细看请打开分类拼版",
            )
        make_review_xlsx(review / company_dir_name(name) / "图片核对表.xlsx", records_by_cat)

    montage_hashes = {}
    montage_root = review / "视觉核对图"
    if montage_root.is_dir():
        for fp in sorted(montage_root.rglob("*.png")):
            montage_hashes[str(fp.relative_to(review)).replace("\\", "/")] = sha256_file(fp)
    template = build_template(records_map, existing, montage_hashes)
    write_json(json_path, {
        "_说明": ("结论三档：符合/不符/待核对（空等同待核对，只有 符合 放行）。"
                 "每张图带 候选锚点 供挂靠参照，并可用 依据 字段写清命中哪个锚点/来源。"
                 "Codex 必须逐张看图后回写，不支持类别级继承。"),
        **template,
    })

    gate_script = Path(__file__).with_name("gates.py")
    en_file = stage / "en.json"
    summary_file = stage / "汇总.xlsx"
    gates_cmd = (f'python "{gate_script}" --deliverable "{deliverable}" --raw "{raw}" '
                 f'--en "{en_file}" --summary "{summary_file}" '
                 f'--visual "{json_path}" --require-visual')
    guide = review / "核对指引.md"
    write_review_guide(guide, records_map, review, json_path, gates_cmd, pages_map)

    print(f"企业 {len(archives)} 家")
    for name in archives:
        total = sum(counts.get((name, folder), 0) for folder in DIRS.values())
        print(f"  {name}: 图片 {total} 张")
        for folder in DIRS.values():
            print(f"    {folder}: {counts.get((name, folder), 0)} 张")
    print(f"拼版：{review / '视觉核对图'}")
    print(f"核对指引：{guide}")
    print(f"核对表：{review / '<企业>' / '图片核对表.xlsx'}")
    print(f"结论文件：{json_path}")
    print("下一步：Codex 打开拼版逐张核对 -> 写 verdicts.json -> --apply -> gates.py --require-visual")
    return 0


if __name__ == "__main__":
    sys.exit(main())
