#!/usr/bin/env python3
"""提取老师 PPT 的内容,供「通俗化改写」使用。

输出(工作目录内):
- slides_text.json        每页文本(段落 / 表格),模型阅读用
- assets/pages/slide_001.png  每页高清渲染图(公式、波形图、电路图都在这里面,直接"看图"读公式)
- assets/images/           PPT 内嵌图片(按页编号另存)

用法:
    python extract.py --input 老师的课.pptx --work ./work

注意:PPT 里的公式常是 OLE/EQ 对象或图片,文本提取拿不到,务必配合
assets/pages/*.png 看图理解,不要只依赖 slides_text.json。
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

sys.path.insert(0, str(Path(__file__).parent))


def walk_shapes(shapes):
    """递归遍历所有形状(包括组合里的)。"""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk_shapes(shape.shapes)
        else:
            yield shape


def extract_text(pptx_path: str, out_dir: Path):
    prs = Presentation(pptx_path)
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        items = []
        for shape in walk_shapes(slide.shapes):
            if shape.has_text_frame and shape.text_frame.text.strip():
                items.append({"kind": "text", "content": shape.text_frame.text.strip()})
            elif shape.has_table:
                rows = [[c.text.strip() for c in r.cells] for r in shape.table.rows]
                items.append({"kind": "table", "content": rows})
        slides.append({"index": i, "items": items})
    (out_dir / "slides_text.json").write_text(
        json.dumps(slides, ensure_ascii=False, indent=2), encoding="utf-8")
    return len(slides)


def extract_images(pptx_path: str, out_dir: Path):
    image_dir = out_dir / "assets" / "images"
    image_dir.mkdir(parents=True, exist_ok=True)
    prs = Presentation(pptx_path)
    n = 0
    for i, slide in enumerate(prs.slides, 1):
        idx = 0
        for shape in walk_shapes(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                idx += 1
                try:
                    img = shape.image
                    fn = image_dir / f"slide_{i:03d}_img_{idx}.{img.ext}"
                    fn.write_bytes(img.blob)
                    n += 1
                except Exception as e:  # 个别图片可能损坏,不阻断主流程
                    print(f"warn: 提取图片失败 slide {i}#{idx}: {e}", file=sys.stderr)
    return n


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True, help="老师 PPT 路径")
    ap.add_argument("--work", required=True, help="工作目录(不存在会自动创建)")
    args = ap.parse_args()

    src = Path(args.input)
    work = Path(args.work)
    if not src.exists():
        sys.exit(f"文件不存在: {src}")
    if src.suffix.lower() not in (".pptx", ".ppt"):
        sys.exit("只支持 .pptx / .ppt")
    work.mkdir(parents=True, exist_ok=True)

    n_text = extract_text(str(src), work)
    n_img = extract_images(str(src), work)
    print(f"已从 {src.name} 提取文本: {n_text} 页")

    # 渲染每页 PNG。失败不致命:文本和图片仍可用,但公式只能靠素材图片或用户补充。
    try:
        subprocess.run([sys.executable, str(Path(__file__).parent / "render.py"),
                        "--input", str(src), "--out", str(work / "assets" / "pages")],
                       check=True)
    except subprocess.CalledProcessError as e:
        print(f"warn: 页面渲染失败({e}),继续使用文本+内嵌图片", file=sys.stderr)

    print(f"提取完毕 -> {work}")
    print("  阅读: slides_text.json ; 看图/公式: assets/pages/*.png ; 内嵌图: assets/images/")


if __name__ == "__main__":
    main()
