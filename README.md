# ppt-simplifier 通俗版课件生成器

把老师的教学课件(尤其是电子信息类专业课:信号与系统、模电、数电、通信原理、电磁场、概率论等)改写成自己看得懂的「学霸讲解版」PPT。

## 核心特性

- **例题、公式、图 100% 无损保留**:每页课件渲染成高清图,例题页整页复用原页截图,不转写、不重画
- **白话讲解**:每页一个知识点,先「一句话版本」再讲大白话要点,专业术语配生活类比
- **作业题重点联动**:把老师布置的作业题与课件考点逐题映射,相关页面标注「呼应作业 x.x.x」并打重点标签(必考 / 掌握 / 了解)
- **自动渲染降级链**:PowerPoint COM → LibreOffice → 页面高清 PNG;公式 OLE(MathType/Equation/Word.Picture)即使编辑器缺失也能渲染
- **布局自动检查**:文本框溢出/越界/重叠自动检测,交付前无需人工排雷

## 安装

```bash
pip install python-pptx pymupdf pillow matplotlib olefile rapidocr_onnxruntime
```

渲染器二选一:Microsoft PowerPoint(Windows)或 LibreOffice(免费)。

## 用法

1. 把本目录放到 ZCode 的 skills 目录:
   - Windows: `%USERPROFILE%\.zcode\skills\ppt-simplifier\`
2. 告诉 ZCode:

   > 这是老师课件 `信号与系统.pptx`,这是作业题 `作业.jpg`,帮我做份通俗版

3. 流水线自动执行:
   - `scripts/extract.py` — 提取文本 + 渲染每页高清图
   - 模型根据提取结果写 `outline.json`(通俗版大纲)
   - `scripts/build.py` — 生成 PPTX + 渲染检查图 + 布局检查
   - `scripts/rasterize_ole.py` — 备选:把 OLE 公式栅格化为图片

## 流程产物

- `slides_text.json` — 每页文本
- `assets/pages/*.png` — 每页 200 DPI 渲染图(公式都在图里)
- `outline.json` — 通俗版大纲(概念页 / 公式页 / 例题页 / 小结页)
- 输出「通俗版.pptx」+ `check/*.png` 检查图
