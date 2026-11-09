# -*- coding: utf-8 -*-
"""
电子签名工具 (Sign Tool)
--------------------------------
功能：上传签名照片 -> 自动去除白色背景生成透明签名 -> 插入 Word (.docx) 文档
支持：浮动自由定位 / 内嵌跟随文字，可指定"文档末尾"或"某段文字之后"插入
依赖：仅 Pillow + Python 标准库（直接读写 docx 包，不依赖 python-docx）
"""

import os
import re
import sys
import html
import zipfile
import tempfile
import traceback

APP_NAME = "电子签名工具"
APP_VERSION = "1.0.0"
EMU_PER_CM = 360000  # 1 cm = 360000 EMU

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
IMG_REL_TYPE = R_NS + "/image"


# =====================================================================
# 一、签名图片处理
# =====================================================================
def process_signature(src_path, threshold=215, ink="keep", width_cm=4.0, dpi=300):
    """把签名照片处理成透明底 PNG（返回 PIL.Image，RGBA）

    threshold: 白底阈值 0-254，越大保留越多墨迹
    ink: 'keep' 保留原墨色 / 'black' 强制黑色
    """
    from PIL import Image, ImageChops

    im = Image.open(src_path)
    im = im.convert("RGBA")
    if im.width < 4 or im.height < 4:
        raise ValueError("图片尺寸过小")

    gray = im.convert("L")
    thr = max(1, min(254, int(threshold)))
    # 亮度 -> 透明度：越接近白越透明，越黑越不透明
    lut = [0 if i >= thr else int(255 * (thr - i) / thr) for i in range(256)]
    alpha = gray.point(lut)

    # 原图若自带透明通道，取两者更透明的一方
    orig_alpha = im.split()[3]
    alpha = ImageChops.darker(alpha, orig_alpha)

    if ink == "black":
        base = Image.new("RGB", im.size, (0, 0, 0))
    else:
        base = im.convert("RGB")

    out = base.convert("RGBA")
    out.putalpha(alpha)

    # 裁掉四周空白
    bbox = alpha.getbbox()
    if bbox:
        pad = 2
        l, t, r, b = bbox
        out = out.crop((max(0, l - pad), max(0, t - pad),
                        min(im.width, r + pad), min(im.height, b + pad)))

    # 按目标宽度缩放（dpi 300 足够清晰）
    target_w = max(30, int(float(width_cm) / 2.54 * dpi))
    ratio = target_w / float(out.width)
    target_h = max(10, int(round(out.height * ratio)))
    out = out.resize((target_w, target_h), Image.LANCZOS)
    return out


def save_png(img, path):
    img.save(path, "PNG", optimize=True)
    return path


# =====================================================================
# 二、docx 包读写
# =====================================================================
def _read_zip(path):
    with zipfile.ZipFile(path, "r") as z:
        return [(i, z.read(i.filename)) for i in z.infolist()]


def _write_zip(path, items):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            z.writestr(zi, data)


def _decode(b):
    return b.decode("utf-8")


def _encode(s):
    return s.encode("utf-8")


def _drawing_run(rid, cx, cy, doc_id, floating, x_emu, y_emu, opacity=100):
    """构造 <w:r><w:drawing>... 片段（浮动或内嵌）"""
    alpha_xml = ""
    if int(opacity) < 100:
        amt = max(5, min(100, int(opacity))) * 1000
        alpha_xml = '<a:alphaModFix amt="%d"/>' % amt

    graphic = (
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic><pic:nvPicPr><pic:cNvPr id="%d" name="signature.png"/>'
        '<pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="%s">%s</a:blip>'
        '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic>'
    ) % (doc_id, rid, alpha_xml, cx, cy)

    ns = (' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
          ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
          ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
          ' xmlns:r="%s"' % R_NS)

    if floating:
        body = (
            '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"'
            ' relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="page"><wp:posOffset>%d</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="page"><wp:posOffset>%d</wp:posOffset></wp:positionV>'
            '<wp:extent cx="%d" cy="%d"/>'
            '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:wrapNone/>'
            '<wp:docPr id="%d" name="电子签名"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '%s</wp:anchor>'
        ) % (x_emu, y_emu, cx, cy, doc_id, graphic)
    else:
        body = (
            '<wp:inline distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%d" cy="%d"/>'
            '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:docPr id="%d" name="电子签名"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '%s</wp:inline>'
        ) % (cx, cy, doc_id, graphic)

    return '<w:r><w:drawing%s>%s</w:drawing></w:r>' % (ns, body)


_PARAGRAPH_RE = re.compile(r"<w:p\b[^>]*>.*?</w:p>", re.S)


def _para_text(xml_frag):
    txt = re.sub(r"<[^>]+>", "", xml_frag)
    return html.unescape(txt)


def _find_para_after_keyword(doc_xml, keyword):
    """返回 (插入位置, 段落起始位置)；未找到返回 None"""
    for m in _PARAGRAPH_RE.finditer(doc_xml):
        if keyword and keyword in _para_text(m.group(0)):
            return m.end()
    return None


def insert_signature(docx_in, docx_out, png_path, mode="float", width_cm=4.0,
                     x_cm=10.0, y_cm=20.0, anchor="end", keyword="",
                     align="right", opacity=100):
    """把签名图片插入 docx

    mode: 'float' 浮动自由定位 / 'inline' 内嵌跟随文字
    anchor: 'end' 文档末尾 / 'keyword' 指定文字所在段落之后
    """
    if not os.path.isfile(docx_in):
        raise FileNotFoundError("找不到 Word 文件：%s" % docx_in)
    if not docx_in.lower().endswith(".docx"):
        raise ValueError("仅支持 .docx 格式，请先用 Word 另存为 .docx")

    items = _read_zip(docx_in)
    names = [i.filename for i, _ in items]
    if "word/document.xml" not in names:
        raise ValueError("不是有效的 Word 文档（缺少 word/document.xml）")

    doc_xml = _decode(dict((i.filename, d) for i, d in items)["word/document.xml"])
    rels_name = "word/_rels/document.xml.rels"
    if rels_name in names:
        rels_xml = _decode(dict((i.filename, d) for i, d in items)[rels_name])
    else:
        rels_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="%s"></Relationships>' % PKG_REL_NS)

    # 1) 媒体文件名（避免重名）
    idx = 1
    while ("word/media/signature_%d.png" % idx) in names:
        idx += 1
    media_name = "word/media/signature_%d.png" % idx

    # 2) 新关系 ID
    used = set(int(n) for n in re.findall(r'Id="rId(\d+)"', rels_xml))
    rid = "rId%d" % (max(used) + 1 if used else 1)
    rel_entry = ('<Relationship Id="%s" Type="%s" Target="media/signature_%d.png"/>'
                 % (rid, IMG_REL_TYPE, idx))
    if "</Relationships>" in rels_xml:
        rels_xml = rels_xml.replace("</Relationships>", rel_entry + "</Relationships>")
    else:
        rels_xml = rels_xml.rstrip()[:-1] + rel_entry + "</Relationships>"

    # 3) 内容类型登记 png
    ct_name = "[Content_Types].xml"
    if ct_name in names:
        ct_xml = _decode(dict((i.filename, d) for i, d in items)[ct_name])
        if not re.search(r'Extension="png"', ct_xml, re.I):
            add = '<Default Extension="png" ContentType="image/png"/>'
            ct_xml = ct_xml.replace("</Types>", add + "</Types>")
    else:
        ct_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  '<Types xmlns="%s"><Default Extension="png" ContentType="image/png"/>'
                  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                  '<Default Extension="xml" ContentType="application/xml"/>'
                  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                  '</Types>' % CT_NS)

    # 4) 尺寸
    from PIL import Image
    with Image.open(png_path) as im:
        px_w, px_h = im.size
    cx = int(round(float(width_cm) * EMU_PER_CM))
    cy = int(round(cx * px_h / float(px_w)))
    x_emu = int(round(float(x_cm) * EMU_PER_CM))
    y_emu = int(round(float(y_cm) * EMU_PER_CM))

    # 5) docPr id 必须唯一
    ids = [int(v) for v in re.findall(r'<wp:docPr[^>]*\bid="(\d+)"', doc_xml)]
    doc_id = (max(ids) + 1) if ids else 1

    floating = (mode == "float")
    run_xml = _drawing_run(rid, cx, cy, doc_id, floating, x_emu, y_emu, opacity)

    # 6) 锚点段落
    align_val = {"left": "left", "center": "center", "right": "right"}.get(align, "right")
    tiny_ppr = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>' \
               '<w:rPr><w:sz w:val="2"/></w:rPr></w:pPr>'
    if floating:
        new_para = '<w:p>%s%s</w:p>' % (tiny_ppr, run_xml)
    else:
        new_para = ('<w:p><w:pPr><w:jc w:val="%s"/></w:pPr>%s</w:p>'
                    % (align_val, run_xml))

    if anchor == "keyword" and keyword.strip():
        pos = _find_para_after_keyword(doc_xml, keyword.strip())
        if pos is None:
            raise ValueError('文档中没有找到包含“%s”的段落，请检查关键字' % keyword.strip())
        doc_xml = doc_xml[:pos] + new_para + doc_xml[pos:]
    else:
        # 文档末尾（插在 body 级 sectPr 之前，保持页面设置）
        end = doc_xml.rfind("</w:body>")
        if end < 0:
            raise ValueError("文档结构异常（缺少 w:body）")
        sect = doc_xml.rfind("<w:sectPr", 0, end)
        pos = sect if sect > 0 else end
        doc_xml = doc_xml[:pos] + new_para + doc_xml[pos:]

    # 7) 写回
    replace = {
        "word/document.xml": _encode(doc_xml),
        rels_name: _encode(rels_xml),
        ct_name: _encode(ct_xml),
    }
    out_items = []
    for info, data in items:
        if info.filename in replace:
            out_items.append((info, replace[info.filename]))
            del replace[info.filename]
        else:
            out_items.append((info, data))
    for name, data in replace.items():
        zi = zipfile.ZipInfo(name, date_time=(2024, 1, 1, 0, 0, 0))
        out_items.append((zi, data))
    out_items.append((zipfile.ZipInfo(media_name, date_time=(2024, 1, 1, 0, 0, 0)),
                      open(png_path, "rb").read()))

    tmp_out = docx_out + ".tmp"
    _write_zip(tmp_out, out_items)
    if os.path.exists(docx_out):
        os.remove(docx_out)
    os.replace(tmp_out, docx_out)
    return docx_out


# =====================================================================
# 三、图形界面
# =====================================================================
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from PIL import Image, ImageTk

    FONT = ("Microsoft YaHei UI", 10)
    root = tk.Tk()
    root.title("%s v%s" % (APP_NAME, APP_VERSION))
    root.geometry("880x560")
    root.minsize(820, 520)

    style = ttk.Style()
    try:
        style.theme_use("vista")
    except Exception:
        pass
    style.configure(".", font=FONT)
    style.configure("Title.TLabel", font=("Microsoft YaHei UI", 11, "bold"))
    style.configure("Big.TButton", font=("Microsoft YaHei UI", 11, "bold"), padding=(12, 8))

    state = {
        "img_path": None,
        "png_tmp": None,
        "preview": None,
        "docx": None,
        "out": None,
    }

    # ---------------- 顶部 ----------------
    header = ttk.Frame(root, padding=(14, 12, 14, 4))
    header.pack(fill="x")
    ttk.Label(header, text="上传签名照片，一键插入 Word 文档",
              style="Title.TLabel").pack(side="left")
    ttk.Label(header, text="v" + APP_VERSION, foreground="#888").pack(side="right")

    main = ttk.Frame(root, padding=(14, 6, 14, 6))
    main.pack(fill="both", expand=True)
    main.columnconfigure(0, weight=1)
    main.columnconfigure(1, weight=1)
    main.rowconfigure(0, weight=1)

    # ---------------- 左：签名 ----------------
    lf1 = ttk.LabelFrame(main, text=" ① 签名图片 ", padding=10)
    lf1.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

    btn_row = ttk.Frame(lf1)
    btn_row.pack(fill="x")
    ttk.Button(btn_row, text="选择签名照片…", command=lambda: choose_image()).pack(side="left")
    ttk.Button(btn_row, text="清除白底预览", command=lambda: do_preview()).pack(side="left", padx=6)

    canvas = tk.Canvas(lf1, width=340, height=190, bg="white",
                       highlightthickness=1, highlightbackground="#cccccc")
    canvas.pack(pady=8)
    canvas.create_text(170, 95, text="尚未选择图片\n支持 JPG / PNG / BMP",
                       fill="#999999", font=FONT, justify="center")

    def add_scale(parent, label, frm, to, init, fmt, cmd=None):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text=label, width=12).pack(side="left")
        val = tk.StringVar(value=fmt(init))
        sc = tk.Scale(row, from_=frm, to=to, orient="horizontal", showvalue=0,
                      resolution=1 if to > 20 else 0.1, length=170,
                      highlightthickness=0)
        sc.set(init)
        sc.pack(side="left", fill="x", expand=True)
        lab = ttk.Label(row, textvariable=val, width=7, anchor="e")
        lab.pack(side="left")

        def on_change(_=None):
            v = sc.get()
            val.set(fmt(v))
            if cmd:
                cmd()

        sc.configure(command=on_change)
        return sc

    thr_scale = add_scale(lf1, "去白底强度", 60, 254, 215, lambda v: "%d" % int(v))
    wid_scale = add_scale(lf1, "签名宽度(cm)", 1.0, 15.0, 4.0, lambda v: "%.1f" % float(v))

    ink_row = ttk.Frame(lf1)
    ink_row.pack(fill="x", pady=(6, 0))
    ttk.Label(ink_row, text="墨色", width=12).pack(side="left")
    ink_var = tk.StringVar(value="keep")
    ttk.Radiobutton(ink_row, text="保留原色", value="keep", variable=ink_var,
                    command=lambda: do_preview()).pack(side="left")
    ttk.Radiobutton(ink_row, text="纯黑", value="black", variable=ink_var,
                    command=lambda: do_preview()).pack(side="left", padx=8)

    # ---------------- 右：Word ----------------
    lf2 = ttk.LabelFrame(main, text=" ② Word 文档与插入方式 ", padding=10)
    lf2.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

    doc_row = ttk.Frame(lf2)
    doc_row.pack(fill="x")
    doc_entry = ttk.Entry(doc_row)
    doc_entry.pack(side="left", fill="x", expand=True)
    ttk.Button(doc_row, text="浏览…", width=7,
               command=lambda: choose_doc()).pack(side="left", padx=(6, 0))

    out_row = ttk.Frame(lf2)
    out_row.pack(fill="x", pady=6)
    ttk.Label(out_row, text="输出文件", width=8).pack(side="left")
    out_entry = ttk.Entry(out_row)
    out_entry.pack(side="left", fill="x", expand=True)

    ttk.Separator(lf2).pack(fill="x", pady=6)

    mode_row = ttk.Frame(lf2)
    mode_row.pack(fill="x")
    ttk.Label(mode_row, text="插入方式", width=8).pack(side="left")
    mode_box = ttk.Combobox(mode_row, state="readonly", width=22,
                            values=["浮动定位（可自由摆放）", "内嵌（跟随文字）"])
    mode_box.current(0)
    mode_box.pack(side="left")

    anchor_row = ttk.Frame(lf2)
    anchor_row.pack(fill="x", pady=6)
    ttk.Label(anchor_row, text="插入位置", width=8).pack(side="left")
    anchor_box = ttk.Combobox(anchor_row, state="readonly", width=22,
                              values=["文档末尾", "指定文字所在段落之后"])
    anchor_box.current(0)
    anchor_box.pack(side="left")

    kw_row = ttk.Frame(lf2)
    kw_row.pack(fill="x")
    ttk.Label(kw_row, text="关键字", width=8).pack(side="left")
    kw_entry = ttk.Entry(kw_row)
    kw_entry.pack(side="left", fill="x", expand=True)
    kw_entry.configure(state="disabled")

    ttk.Separator(lf2).pack(fill="x", pady=6)

    pos_frame = ttk.Frame(lf2)
    pos_frame.pack(fill="x")
    x_scale = add_scale(pos_frame, "距页左(cm)", 0, 21, 12.0, lambda v: "%.1f" % float(v))
    y_scale = add_scale(pos_frame, "距页顶(cm)", 0, 30, 22.0, lambda v: "%.1f" % float(v))

    al_row = ttk.Frame(lf2)
    al_row.pack(fill="x", pady=4)
    ttk.Label(al_row, text="对齐方式", width=8).pack(side="left")
    al_box = ttk.Combobox(al_row, state="readonly", width=10, values=["左", "居中", "右"])
    al_box.current(2)
    al_box.pack(side="left")

    ttk.Label(lf2, text="提示：浮动定位的坐标以所在页面的左上角为原点，可用 0.1cm 微调。",
              foreground="#777777", wraplength=360, justify="left").pack(fill="x", pady=(6, 0))

    def sync_mode(*_):
        floating = mode_box.current() == 0
        st = "normal" if floating else "disabled"
        for s in (x_scale, y_scale):
            s.configure(state=st)
        al_box.configure(state="readonly" if not floating else "disabled")

    def sync_anchor(*_):
        kw_entry.configure(state="normal" if anchor_box.current() == 1 else "disabled")

    mode_box.bind("<<ComboboxSelected>>", sync_mode)
    anchor_box.bind("<<ComboboxSelected>>", sync_anchor)
    sync_mode()
    sync_anchor()

    # ---------------- 底部 ----------------
    bottom = ttk.Frame(root, padding=(14, 4, 14, 12))
    bottom.pack(fill="x")
    status = ttk.Label(bottom, text="就绪", foreground="#555555")
    status.pack(side="left")

    def set_status(msg, color="#555555"):
        status.configure(text=msg, foreground=color)
        root.update_idletasks()

    # ---------------- 逻辑 ----------------
    def choose_image():
        p = filedialog.askopenfilename(
            title="选择签名照片",
            filetypes=[("图片", "*.jpg *.jpeg *.png *.bmp *.webp *.tif *.tiff"),
                       ("所有文件", "*.*")])
        if not p:
            return
        state["img_path"] = p
        set_status("已选择：%s" % os.path.basename(p))
        do_preview()

    def do_preview():
        if not state["img_path"]:
            return
        try:
            img = process_signature(state["img_path"],
                                    threshold=thr_scale.get(),
                                    ink=ink_var.get(),
                                    width_cm=float(wid_scale.get()))
        except Exception as e:
            messagebox.showerror(APP_NAME, "图片处理失败：\n%s" % e)
            return
        tmp = os.path.join(tempfile.gettempdir(), "sign_tool_preview.png")
        img.save(tmp, "PNG")
        state["png_tmp"] = tmp

        cw, ch = 340, 190
        w, h = img.size
        scale = min(cw / float(w), ch / float(h), 1.0)
        shown = img.resize((max(1, int(w * scale)), max(1, int(h * scale))))
        bg = Image.new("RGBA", shown.size, (255, 255, 255, 255))
        bg.alpha_composite(shown)
        state["preview"] = ImageTk.PhotoImage(bg.convert("RGB"))
        canvas.delete("all")
        canvas.create_image(cw // 2, ch // 2, image=state["preview"])
        set_status("预览已更新（%d×%d px）" % (img.width, img.height), "#1a7f37")

    def choose_doc():
        p = filedialog.askopenfilename(title="选择 Word 文档",
                                       filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if not p:
            return
        state["docx"] = p
        doc_entry.delete(0, "end")
        doc_entry.insert(0, p)
        base, ext = os.path.splitext(p)
        out = base + "_已签名" + ext
        state["out"] = out
        out_entry.delete(0, "end")
        out_entry.insert(0, out)
        set_status("已选择文档：%s" % os.path.basename(p))

    def do_insert():
        if not state["img_path"]:
            messagebox.showwarning(APP_NAME, "请先选择签名照片")
            return
        if not doc_entry.get().strip():
            messagebox.showwarning(APP_NAME, "请先选择要签名的 Word 文档")
            return
        docx = doc_entry.get().strip()
        out = out_entry.get().strip() or (os.path.splitext(docx)[0] + "_已签名.docx")
        if os.path.abspath(out) == os.path.abspath(docx):
            out = os.path.splitext(docx)[0] + "_已签名.docx"
            out_entry.delete(0, "end")
            out_entry.insert(0, out)

        # 生成最新签名图
        try:
            img = process_signature(state["img_path"],
                                    threshold=thr_scale.get(),
                                    ink=ink_var.get(),
                                    width_cm=float(wid_scale.get()))
            png = os.path.join(tempfile.gettempdir(), "sign_tool_sign.png")
            save_png(img, png)
        except Exception as e:
            messagebox.showerror(APP_NAME, "签名图片处理失败：\n%s" % e)
            return

        try:
            set_status("正在插入签名…")
            insert_signature(
                docx_in=docx, docx_out=out, png_path=png,
                mode="float" if mode_box.current() == 0 else "inline",
                width_cm=float(wid_scale.get()),
                x_cm=float(x_scale.get()),
                y_cm=float(y_scale.get()),
                anchor="keyword" if anchor_box.current() == 1 else "end",
                keyword=kw_entry.get().strip(),
                align=["left", "center", "right"][al_box.current()],
            )
        except Exception as e:
            messagebox.showerror(APP_NAME, "插入失败：\n%s" % e)
            set_status("插入失败", "#c00")
            return

        set_status("完成：%s" % os.path.basename(out), "#1a7f37")
        if messagebox.askyesno(APP_NAME, "签名已插入并保存到：\n%s\n\n现在打开文档查看吗？" % out):
            try:
                os.startfile(out)
            except Exception:
                pass

    ttk.Button(bottom, text="生成签名并插入 Word", style="Big.TButton",
               command=do_insert).pack(side="right")
    ttk.Button(bottom, text="退出", command=root.destroy).pack(side="right", padx=8)

    root.mainloop()


# =====================================================================
# 四、自检（打包后可用 电子签名工具.exe --selftest 验证）
# =====================================================================
def _make_test_docx(path, paragraphs):
    body = "".join(
        '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr>'
        '<w:rFonts w:ascii="SimSun" w:eastAsia="SimSun"/><w:sz w:val="24"/></w:rPr>'
        '<w:t xml:space="preserve">%s</w:t></w:r></w:p>' % html.escape(t) for t in paragraphs)
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="%s" xmlns:r="%s"><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"'
           ' w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
           '</w:body></w:document>' % (W_NS, R_NS, body))
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="%s"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '</Types>' % CT_NS)
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s/officeDocument"'
            ' Target="word/document.xml"/></Relationships>' % (PKG_REL_NS, R_NS))
    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="%s"></Relationships>' % PKG_REL_NS)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", doc_rels)


def _make_test_signature(path):
    from PIL import Image, ImageDraw
    im = Image.new("RGB", (900, 400), (252, 252, 250))
    d = ImageDraw.Draw(im)
    # 模拟手写签名：几条曲线
    pts = [(120, 300), (170, 180), (220, 250), (280, 140), (330, 280),
           (400, 200), (470, 260), (560, 150), (640, 240), (760, 190)]
    d.line(pts, fill=(20, 30, 90), width=9, joint="curve")
    d.line([(120, 320), (780, 300)], fill=(20, 30, 90), width=5)
    # 加一点纸张噪点
    for i in range(0, 900, 7):
        d.point((i, (i * 37) % 400), fill=(240, 240, 240))
    im.save(path, "JPEG", quality=92)
    return path


def selftest():
    log = []
    ok = True

    def check(name, cond, extra=""):
        nonlocal ok
        log.append(("[PASS] " if cond else "[FAIL] ") + name + (("  -> " + str(extra)) if extra else ""))
        if not cond:
            ok = False

    tmpdir = tempfile.mkdtemp(prefix="signtool_test_")
    try:
        # 1. 图片处理
        jpg = _make_test_signature(os.path.join(tmpdir, "sign.jpg"))
        img = process_signature(jpg, threshold=215, ink="keep", width_cm=4.0)
        check("签名图片生成透明底 PNG", img.mode == "RGBA" and img.width > 100, img.size)
        a = img.split()[3]
        check("四角为透明（白底已去除）",
              a.getpixel((0, 0)) == 0 and a.getpixel((img.width - 1, 0)) == 0,
              (a.getpixel((0, 0)), a.getpixel((img.width - 1, 0))))
        mx = a.getextrema()[1]
        check("墨迹区域不透明", mx > 200, mx)
        check("宽度符合 4cm@300dpi", abs(img.width - int(4 / 2.54 * 300)) <= 2, img.width)

        png = os.path.join(tmpdir, "sign.png")
        save_png(img, png)

        # 2. 文档：浮动定位
        src = os.path.join(tmpdir, "src.docx")
        _make_test_docx(src, ["劳动合同", "甲方：某某公司", "乙方：张三", "签字："])
        out1 = os.path.join(tmpdir, "out_float.docx")
        insert_signature(src, out1, png, mode="float", width_cm=4.0,
                         x_cm=10.0, y_cm=20.0, anchor="end")

        with zipfile.ZipFile(out1) as z:
            names = z.namelist()
            dx = z.read("word/document.xml").decode("utf-8")
            rx = z.read("word/_rels/document.xml.rels").decode("utf-8")
            ct = z.read("[Content_Types].xml").decode("utf-8")
            media = [n for n in names if n.startswith("word/media/")]
        check("媒体文件已写入", len(media) == 1, media)
        check("关系项已登记", 'Type="%s"' % IMG_REL_TYPE in rx and "media/signature_1.png" in rx)
        check("Content_Types 含 png", 'Extension="png"' in ct)
        check("存在浮动锚点 wp:anchor", "<wp:anchor" in dx)
        check("X 坐标 10cm 正确", "<wp:posOffset>%d</wp:posOffset>" % int(10 * EMU_PER_CM) in dx)
        check("Y 坐标 20cm 正确", "<wp:posOffset>%d</wp:posOffset>" % int(20 * EMU_PER_CM) in dx)
        check("插入在 sectPr 之前", dx.find("<w:drawing") < dx.rfind("<w:sectPr"))
        embed = re.findall(r'r:embed="(rId\d+)"', dx)
        check("关系 ID 与嵌入一致",
              len(embed) == 1 and ('Id="%s"' % embed[0]) in rx
              and 'Target="media/signature_1.png"' in rx, embed)

        # 3. 文档：内嵌 + 关键字
        out2 = os.path.join(tmpdir, "out_inline.docx")
        insert_signature(src, out2, png, mode="inline", width_cm=5.0,
                         anchor="keyword", keyword="签字", align="right")
        with zipfile.ZipFile(out2) as z:
            dx2 = z.read("word/document.xml").decode("utf-8")
        check("存在内嵌 wp:inline", "<wp:inline" in dx2 and "<wp:anchor" not in dx2)
        check("内嵌右对齐", '<w:jc w:val="right"/>' in dx2)
        para_list = _PARAGRAPH_RE.findall(dx2)
        idx_sig = next((i for i, p in enumerate(para_list) if "签字" in _para_text(p)), -1)
        check("插入在关键字段落之后",
              idx_sig >= 0 and idx_sig + 1 < len(para_list)
              and "<w:drawing" in para_list[idx_sig + 1])
        check("5cm 宽度正确", 'cx="%d"' % int(5 * EMU_PER_CM) in dx2)

        # 4. 关键字不存在时应报错
        try:
            insert_signature(src, os.path.join(tmpdir, "x.docx"), png,
                             anchor="keyword", keyword="不存在的词")
            check("关键字缺失时给出错误", False)
        except ValueError:
            check("关键字缺失时给出错误", True)

        # 5. 二次插入同一文档不冲突
        out3 = os.path.join(tmpdir, "out_twice.docx")
        insert_signature(out1, out3, png, mode="inline", width_cm=3.0, anchor="end")
        with zipfile.ZipFile(out3) as z:
            dx3 = z.read("word/document.xml").decode("utf-8")
            m3 = [n for n in z.namelist() if n.startswith("word/media/")]
        check("重复插入生成独立媒体文件", len(m3) == 2, m3)
        check("重复插入 docPr id 唯一",
              len(re.findall(r'<wp:docPr id="1"', dx3)) == 1)

        # 6. 产出文件可被解析
        check("输出为合法 zip 且体积正常", os.path.getsize(out1) > 3000, os.path.getsize(out1))

    except Exception:
        ok = False
        log.append("[FAIL] 异常：\n" + traceback.format_exc())
    finally:
        report = "%s %s\n\n" % (APP_NAME, APP_VERSION) + "\n".join(log) + \
                 "\n\n结果：%s\n" % ("全部通过" if ok else "存在失败")
        rp = os.path.join(tempfile.gettempdir(), "sign_tool_selftest.log")
        try:
            with open(rp, "w", encoding="utf-8") as f:
                f.write(report)
        except Exception:
            pass
        print(report)
        try:
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass
    return 0 if ok else 1


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    if "--selftest" in sys.argv:
        sys.exit(selftest())
    try:
        run_gui()
    except Exception:
        err = traceback.format_exc()
        try:
            with open(os.path.join(tempfile.gettempdir(), "sign_tool_error.log"), "w",
                      encoding="utf-8") as f:
                f.write(err)
        except Exception:
            pass
        try:
            import tkinter.messagebox as mb
            mb.showerror(APP_NAME, err)
        except Exception:
            pass


if __name__ == "__main__":
    main()
