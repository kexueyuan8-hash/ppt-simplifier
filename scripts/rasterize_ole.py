#!/usr/bin/env python3
"""把 pptx 里的公式 OLE 对象栅格化:提取内部 EMF/WMF 图形 → PNG → 插回原位置。

背景:很多课程 PPT 的公式是 Word.Picture.8 / MathType / Equation 3.0 这类
OLE 嵌入对象。公式编辑器没装的机器上经常渲染不了,甚至整份文件连 PowerPoint
都无法打开(加载 OLE server 失败)。把公式转成普通 PNG 后,任何渲染器
(PowerPoint / LibreOffice)都能正确显示——这是最泛化的公式处理方式。

做法:
1. 解析每个 ppt/embeddings/oleObjectN.bin(OLE2 复合文档,用 olefile)
2. 提取其中的 EMF/WMF 数据(Word.Picture 的 CONTENTS 流等);
   识别不了的(如 MTEF 二进制)保留 OLE 原样,交给渲染器
3. EMF/WMF 用 Windows 自带 GDI+ (System.Drawing) 转成 PNG
4. 把 slide XML 里的 <p:graphicFrame> 替换为等位置等尺寸的 <p:pic>
5. 重建 pptx 包

用法:
    python rasterize_ole.py --input in.pptx --output out.pptx
退出码 0 = 完成。
"""
import argparse
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

REL_TYPE_IMAGE = ("http://schemas.openxmlformats.org/"
                  "officeDocument/2006/relationships/image")
EMU_PER_IN = 914400


def find_emf_in_stream(data: bytes):
    """在字节流里找 EMF/WMF 数据。返回 (kind, bytes) 或 None。"""
    if data[:4] == b"\xd7\xcd\xc6\x9a":          # WMF 魔数
        return "wmf", data
    if data[:4] == b"\x01\x00\x00\x00" and b"EMF" in data[:24]:
        return "emf", data
    i = data.find(b"\x01\x00\x00\x00\x01\x00\x00\x00")
    if i >= 0 and b"EMF" in data[i + 8:i + 12]:
        return "emf", data[i:]
    return None


def extract_formula_image(ole_path: Path):
    """从 OLE2 复合文档提取公式图形数据。返回 (kind, bytes) 或 None。"""
    import olefile
    if not olefile.isOleFile(str(ole_path)):
        return None
    ole = olefile.OleFileIO(str(ole_path))
    try:
        for s in ole.listdir():
            name = "/".join(s).lower()
            if name in ("contents", "package", "equation native", "native",
                        "\x01native"):
                try:
                    data = ole.openstream(s).read()
                except Exception:
                    continue
                hit = find_emf_in_stream(data)
                if hit:
                    return hit
        for s in ole.listdir():
            try:
                data = ole.openstream(s).read()
            except Exception:
                continue
            hit = find_emf_in_stream(data)
            if hit:
                return hit
    finally:
        ole.close()
    return None


def emf_to_png(emf_bytes: bytes, png_out: Path):
    """用 Windows GDI+ (System.Drawing) 把 EMF/WMF 渲染成 PNG。"""
    with tempfile.TemporaryDirectory() as td:
        src = Path(td) / "formula.emf"
        src.write_bytes(emf_bytes)
        ps = """
Add-Type -AssemblyName System.Drawing
$img = [System.Drawing.Image]::FromFile('{src}')
$bmp = New-Object System.Drawing.Bitmap($img.Width, $img.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear([System.Drawing.Color]::White)
$g.DrawImage($img, 0, 0, $img.Width, $img.Height)
$g.Dispose()
$bmp.Save('{dst}', [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose(); $img.Dispose()
""".replace("{src}", str(src).replace("'", "''")) \
    .replace("{dst}", str(png_out).replace("'", "''"))
        ps1 = Path(td) / "conv.ps1"
        ps1.write_text(ps, encoding="utf-8-sig")
        r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                            "-File", str(ps1)], capture_output=True, text=True,
                           encoding="gbk", errors="replace", timeout=120)
        if r.returncode != 0 or not png_out.exists():
            raise RuntimeError((r.stderr or r.stdout)[-500:])
    return png_out


# ---------- pptx 包内 XML 操作(字符串级,针对 PowerPoint 标准结构) ----------

GRAPHIC_FRAME_RE = re.compile(r"<p:graphicFrame>.*?</p:graphicFrame>", re.S)


def build_pic_tag(frame_xml: str, img_rid: str) -> str:
    """由原 graphicFrame 生成等尺寸 p:pic 标签。找不到 xfrm 返回 None。"""
    xfrm_m = re.search(r"<a:xfrm>.*?</a:xfrm>", frame_xml, re.S)
    if not xfrm_m:
        return None
    cid_m = re.search(r'<p:cNvPr id="(\d+)"', frame_xml)
    cid = cid_m.group(1) if cid_m else "1"
    return (
        '<p:pic><p:nvPicPr><p:cNvPr id="{cid}" name="FormulaImg"/>'
        '<p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
        '<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch>'
        '</p:blipFill><p:spPr>{xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '</p:spPr></p:pic>'
    ).format(cid=cid, rid=img_rid, xfrm=xfrm_m.group(0))


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--work", help="公式 PNG 输出目录(默认 output 旁边)")
    args = ap.parse_args()
    src, out = Path(args.input), Path(args.output)
    work = Path(args.work) if args.work else out.parent / "formula_images"
    work.mkdir(parents=True, exist_ok=True)

    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    media_new = {}            # "ppt/media/xxx.png" -> bytes
    content_types = zin.read("[Content_Types].xml").decode("utf-8", errors="replace")
    if '<Default Extension="png"' not in content_types:
        content_types = content_types.replace(
            "</Types>", '<Default Extension="png" '
            'ContentType="image/png"/></Types>')
    stats = {"栅格化": 0, "无EMF跳过": 0}

    zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    try:
        for n in names:
            if n in ("[Content_Types].xml",) or n.startswith("ppt/slides/"):
                continue
            zout.writestr(n, zin.read(n))
        zout.writestr("[Content_Types].xml", content_types.encode("utf-8"))

        slide_re = re.compile(r"ppt/slides/slide(\d+)\.xml$")
        for n in sorted([x for x in names if slide_re.match(x)],
                        key=lambda x: int(slide_re.match(x).group(1))):
            num = slide_re.match(n).group(1)
            xml = zin.read(n).decode("utf-8", errors="replace")
            rels_name = f"ppt/slides/_rels/slide{num}.xml.rels"
            try:
                rels_xml = zin.read(rels_name).decode("utf-8", errors="replace")
            except KeyError:
                rels_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                            '<Relationships xmlns="http://schemas.openxmlformats.org/'
                            'package/2006/relationships"></Relationships>')

            rel_map = {}
            for rm in re.finditer(r'<Relationship Id="(rId\d+)"[^>]*Target="([^"]+)"',
                                  rels_xml):
                rel_map[rm.group(1)] = rm.group(2)

            if "<p:oleObj" not in xml:
                zout.writestr(n, xml.encode("utf-8"))
                if rels_name in names:
                    zout.writestr(rels_name, rels_xml.encode("utf-8"))
                continue

            new_xml, last, rid_no = "", 0, 2000
            new_rels_parts = []
            for fm in GRAPHIC_FRAME_RE.finditer(xml):
                new_xml += xml[last:fm.start()]
                frame = fm.group(0)
                ole_m = re.search(r'<p:oleObj\b[^>]*r:id="(rId\d+)"', frame)
                if not ole_m:
                    new_xml += frame
                    last = fm.end()
                    continue
                rid = ole_m.group(1)
                target = rel_map.get(rid, "")
                if not target.endswith(".bin"):
                    new_xml += frame
                    last = fm.end()
                    continue
                emb_key = target[2:].lstrip("/")  # "../embeddings/x.bin" -> "embeddings/x.bin"
                try:
                    emb = zin.read("ppt/" + emb_key)
                except KeyError:
                    new_xml += frame
                    last = fm.end()
                    continue
                with tempfile.TemporaryDirectory() as td:
                    olefile = Path(td) / "x.bin"
                    olefile.write_bytes(emb)
                    hit = extract_formula_image(olefile)
                if not hit:
                    stats["无EMF跳过"] += 1
                    new_xml += frame                    # 保留 OLE 原样
                    last = fm.end()
                    continue
                kind, raw = hit
                png_name = f"media/ole_s{num}_{rid}.png"
                png = work / png_name.replace("media/", "")
                try:
                    emf_to_png(raw, png)
                except Exception as e:
                    print(f"warn: 公式转换失败 slide{num} {rid}: {e}",
                          file=sys.stderr)
                    stats["无EMF跳过"] += 1
                    new_xml += frame
                    last = fm.end()
                    continue
                stats["栅格化"] += 1
                rid_no += 1
                img_rid = f"rId{rid_no}"
                media_new["ppt/media/" + png_name] = png.read_bytes()
                pic = build_pic_tag(frame, img_rid)
                if pic is None:
                    new_xml += frame
                    last = fm.end()
                    continue
                new_xml += pic
                new_rels_parts.append(
                    f'<Relationship Id="{img_rid}" Type="{REL_TYPE_IMAGE}" '
                    f'Target="../{png_name}"/>')
                last = fm.end()
            new_xml += xml[last:]
            if new_rels_parts:
                rels_xml = rels_xml.replace(
                    "</Relationships>",
                    "".join(new_rels_parts) + "</Relationships>")
            zout.writestr(n, new_xml.encode("utf-8"))
            zout.writestr(rels_name, rels_xml.encode("utf-8"))
    finally:
        zin.close()
    # 新 media 文件写入
    if media_new:
        with zipfile.ZipFile(out, "a", zipfile.ZIP_DEFLATED) as za:
            for k, v in media_new.items():
                za.writestr(k, v)

    print(f"完成: {stats['栅格化']} 个公式已转为图片,"
          f"{stats['无EMF跳过']} 个保留原样(交给渲染器)")
    print(f"输出 -> {out}")


if __name__ == "__main__":
    main()
