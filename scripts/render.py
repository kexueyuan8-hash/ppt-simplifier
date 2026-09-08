#!/usr/bin/env python3
"""把任意 pptx 渲染成每页一张高清 PNG。

Windows 上优先用 PowerPoint COM(已安装 Office 时),链路:
PowerPoint COM 导出 PDF (ExportAsFixedFormat) -> PyMuPDF 按 200 DPI 渲染。
无 PowerPoint 时降级尝试 LibreOffice (soffice --convert-to pdf)。

用法:
    python render.py --input course.pptx --out dir/pages
输出:
    dir/pages/slide_001.png, slide_002.png, ...
"""
import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

POWERSHELL = shutil.which("powershell") or "powershell"


def render_via_powerpoint(pptx_path: str, out_dir: str, dpi: int = 200):
    """用 PowerPoint COM 导出页面图片。

    优先用 Export 直接把每页导出为 PNG;失败再退回 ExportAsFixedFormat 导出 PDF。
    成功返回 (mode, 产物路径):
        ("png", png_dir) 已把每页 PNG 放到 out_dir
        ("pdf", pdf_path) PDF 已生成,调用方继续转 PNG
    失败返回 (None, None)。
    """
    abs_pptx = str(Path(pptx_path).resolve())
    # PowerShell 单引号字符串里,单引号需要双写转义
    safe_pptx = abs_pptx.replace("'", "''")
    png_folder = Path(out_dir) / "_ppt_export_tmp"
    pdf_path = Path(out_dir) / "_ppt_export.pdf"
    safe_png = str(png_folder.resolve()).replace("'", "''")
    safe_pdf = str(pdf_path.resolve()).replace("'", "''")
    ps = """
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
try {
    $folder = '{folder}'
    New-Item -ItemType Directory -Force -Path $folder | Out-Null
    $pres = $app.Presentations.Open('{pptx}', $true, $false, $false)
    $saved = $false
    try {
        $pres.Export($folder, 'PNG', [int]{w}, [int]{h})
        if (Test-Path $folder) { $saved = $true }
    } catch {
        Write-Output ("EXPORT_FALLBACK: " + $_.Exception.Message)
    }
    if (-not $saved) {
        $pres.ExportAsFixedFormat('{pdf}', [int]2, [int]1, [int]1, [int]1, [int]2, $true, $true, [int]1, $true, $false, $false)
    }
    $pres.Close()
} finally {
    $app.Quit()
}
""".replace("{pptx}", safe_pptx).replace("{folder}", safe_png) \
      .replace("{pdf}", safe_pdf).replace("{w}", str(dpi * 10)) \
      .replace("{h}", str(int(dpi * 10 * 9 / 16)))
    with tempfile.TemporaryDirectory() as td:
        ps1 = Path(td) / "export.ps1"
        ps1.write_text(ps, encoding="utf-8-sig")
        try:
            r = subprocess.run(
                [POWERSHELL, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps1)],
                capture_output=True, text=True, encoding="gbk", errors="replace", timeout=300,
            )
        except subprocess.TimeoutExpired:
            print("render_via_powerpoint: timed out", file=sys.stderr)
            return None, None
    if r.returncode != 0:
        err = (r.stderr or "")[-2000:]
        print("render_via_powerpoint failed:", err, file=sys.stderr)
        return None, None
    pngs = sorted(png_folder.glob("*.png"),
                  key=lambda p: int("".join(ch for ch in p.stem if ch.isdigit()) or 0)) \
        if png_folder.exists() else []
    if pngs:
        out = Path(out_dir)
        out.mkdir(parents=True, exist_ok=True)
        for i, p in enumerate(pngs, 1):
            p.replace(out / f"slide_{i:03d}.png")  # replace 可覆盖已存在文件
        shutil.rmtree(png_folder, ignore_errors=True)
        return "png", str(out)
    if pdf_path.exists():
        return "pdf", str(pdf_path)
    return None, None


def render_via_libreoffice(pptx_path: str, pdf_dir: str) -> bool:
    """用 LibreOffice 导出 PDF(未装 PowerPoint 时的备选)。

    LibreOffice 的 PDF 文件名与源文件同名,所以用 pdf_dir 目录 + 源文件主名来定位。
    """
    soffice = shutil.which("soffice")
    if not soffice:
        for p in [r"C:\Program Files\LibreOffice\program\soffice.exe",
                  r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]:
            if Path(p).exists():
                soffice = p
                break
    if not soffice:
        return False
    src = Path(pptx_path).resolve()
    out_dir = Path(pdf_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    r = subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir),
                        str(src)],
                       capture_output=True, text=True, encoding="utf-8", errors="replace",
                       timeout=600)
    pdf = out_dir / (src.stem + ".pdf")
    return r.returncode == 0 and pdf.exists()


def pdf_to_png(pdf_path: str, out_dir: str, dpi: int = 200):
    import fitz  # PyMuPDF
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf_path)
    n = len(doc)
    for i, page in enumerate(doc, 1):
        pix = page.get_pixmap(dpi=dpi)
        pix.save(str(out / f"slide_{i:03d}.png"))
    doc.close()
    return n


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True, help="pptx 路径")
    ap.add_argument("--out", required=True, help="输出 PNG 目录")
    ap.add_argument("--dpi", type=int, default=200)
    args = ap.parse_args()

    pptx_path = Path(args.input)
    out_dir = Path(args.out)
    if not pptx_path.exists():
        sys.exit(f"文件不存在: {pptx_path}")

    # 优先:PowerPoint COM 直接导出每页 PNG
    mode, path = render_via_powerpoint(str(pptx_path), str(out_dir), args.dpi)
    if mode == "png":
        n = len(list(out_dir.glob("slide_*.png")))
        print(f"已渲染 {n} 页 -> {out_dir}")
        return
    if mode == "pdf":
        n = pdf_to_png(path, str(out_dir), args.dpi)
        print(f"已渲染 {n} 页 -> {out_dir}")
        return

    # 备选:LibreOffice 转 PDF 再渲染(PDF 文件名与源文件同名)
    with tempfile.TemporaryDirectory() as td:
        if render_via_libreoffice(str(pptx_path), td):
            pdf = Path(td) / f"{Path(pptx_path).stem}.pdf"
            n = pdf_to_png(str(pdf), str(out_dir), args.dpi)
            print(f"已渲染 {n} 页 -> {out_dir}")
            return
    sys.exit("渲染失败: 需要 PowerPoint 或 LibreOffice 之一。\n"
             "也可以手动: 用 PowerPoint「另存为 PDF」,再把 PDF 交给本脚本:"
             "python render.py --input out.pdf --out dir/pages")


if __name__ == "__main__":
    main()
