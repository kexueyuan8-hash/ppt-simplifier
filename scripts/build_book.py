"""米白讲义版：按可编辑大纲生成；原图按字节保留，可从交付PPT恢复。
Usage: python build_book.py --outline outline.json --work work
"""
import sys, json, argparse, math, zipfile, posixpath
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python_libs'))
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from PIL import ImageFont
import xml.etree.ElementTree as E
BG='FAF6EC'; INK='322D27'; BROWN='66513D'; MUTED='75634E'; LINE='D8CBBB'
FONT='Microsoft YaHei'; W=13.333; H=7.5
def color(h):return RGBColor.from_string(h)
def text(slide,x,y,w,h,content,size=20,bold=False,ink=INK):
    box=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
    tf=box.text_frame; tf.word_wrap=True
    tf.margin_left=tf.margin_right=0; tf.margin_top=tf.margin_bottom=0
    for i,line in enumerate(content.split('\n')):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.space_after=Pt(10); p.line_spacing=1.25
        r=p.add_run(); r.text=line; r.font.name=FONT; r.font.size=Pt(size);r.font.bold=bold;r.font.color.rgb=color(ink)
        ea=OxmlElement('a:ea');ea.set('typeface',FONT);r._r.get_or_add_rPr().append(ea)
    return box
def rule(s,x,y,w):
    line=s.shapes.add_connector(1, Inches(x), Inches(y), Inches(x+w), Inches(y))
    line.line.color.rgb=color(LINE);line.line.width=Pt(.7)
def source_bytes(d,work,outline_path):
    p=work/d['source']
    if p.exists():return p.read_bytes()
    spec=d['embedded_source']; ppt=outline_path.parent.parent/spec['deck']
    with zipfile.ZipFile(ppt) as z:
        return z.read(spec['media'])
def make(outline_path,work):
    import io
    data=json.loads(outline_path.read_text('utf-8')); prs=Presentation()
    prs.slide_width=Inches(W);prs.slide_height=Inches(H)
    for i,d in enumerate(data['slides'],1):
        s=prs.slides.add_slide(prs.slide_layouts[6]); s.background.fill.solid();s.background.fill.fore_color.rgb=color(BG)
        if d['type']=='cover':
            text(s,.8,1.2,11.5,.5,'电子技术基础 · 模拟部分',18,ink=MUTED)
            text(s,.8,2.15,11.5,1.5,data['title'],42,True)
            text(s,.85,4.1,10.8,1.2,'概念讲清楚，例题一步一步算\n米白讲义版',22,ink=MUTED)
            rule(s,.8,5.8,11.7)
            text(s,.85,6.1,11.2,.5,'适合已掌握电压、电流、电阻基础的自学者',17,ink=MUTED)
            continue
        label=d.get('source_label',''); tag=d.get('exam','掌握')
        text(s,.6,.25,11.8,.3,f"第{data['chapter']}章　／　{tag}　／　{label}",11,ink=MUTED)
        title=d['title'];text(s,.6,.75,12.05,.75,title,30 if len(title)<26 else 27,True)
        rule(s,.6,1.6,12.1)
        if d['type']=='source':
            blob=source_bytes(d,work,outline_path)
            from PIL import Image
            im=Image.open(io.BytesIO(blob));iw,ih=im.size
            scale=min(8.65/iw,5.48/ih);pw,ph=iw*scale,ih*scale
            s.shapes.add_picture(io.BytesIO(blob),Inches(.6+(8.65-pw)/2), Inches(1.77+(5.48-ph)/2),Inches(pw),Inches(ph))
            text(s,9.65,1.93,3.05,.4,'读图提示',18,True,ink=BROWN)
            text(s,9.65,2.5,3.05,2.0,d.get('guide','先看原页中的条件、符号和结论，再读相邻的白话讲解。'),17)
            text(s,9.65,4.75,3.05,.4,'名词小抄',17,True,ink=BROWN)
            text(s,9.65,5.28,3.05,1.7,'\n'.join(d.get('terms',[])[:2]),15,ink=MUTED)
        else:
            text(s,.65,1.9,11.8,.95,d.get('one_liner',''),22,ink=BROWN)
            blocks=d.get('blocks',[])
            y=3.05
            for head,body in blocks:
                text(s,.65,y,7.85,.4,head,19,True)
                text(s,.65,y+.51,7.85,1.12,body,20)
                y+=1.8
            text(s,9.2,3.05,3.45,.4,'名词小抄',18,True,ink=BROWN)
            text(s,9.2,3.6,3.45,2.7,'\n'.join(d.get('terms',[])),17,ink=MUTED)
            if d.get('note'):text(s,.65,6.72,11.85,.43,d['note'],14,ink=BROWN)
        text(s,12.0,7.2,.7,.25,f'{i:03}',10,ink=MUTED)
        s.notes_slide.notes_text_frame.text='依据：'+label+'\n'+d.get('verification','')
    dest=work/data['output'];prs.save(dest)
    # 记录每张原图在成品内的位置，使大纲可脱离临时渲染图再次生成。
    ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    with zipfile.ZipFile(dest) as z:
        for i,d in enumerate(data['slides'],1):
            if d['type']!='source':continue
            node=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'))
            rid=node.find('.//a:blip',ns).get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            rel={e.get('Id'):e.get('Target') for e in E.fromstring(z.read(f'ppt/slides/_rels/slide{i}.xml.rels'))}
            d['embedded_source']={'deck':data['output'],'media':posixpath.normpath(posixpath.join('ppt/slides',rel[rid]))}
    (work/'outline.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),'utf-8')
    print(str(dest),len(prs.slides),'slides')
if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--outline',required=True);ap.add_argument('--work',required=True);a=ap.parse_args();make(Path(a.outline),Path(a.work))
