"""Branded 16:9 slide engine (python-pptx). No chevron motif. Adapt the palette below to the brand.
Put logo PNGs in an assets/ folder beside this file, or it falls back to styled text."""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ---- BRAND PALETTE (edit these) ----
RED=RGBColor(0xCF,0x33,0x38); RED_DARK=RGBColor(0xC0,0x00,0x00)
BG=RGBColor(0xEE,0xEC,0xE1); BODY=RGBColor(0x70,0x70,0x70)
DARK=RGBColor(0x1F,0x21,0x20); WHITE=RGBColor(0xFF,0xFF,0xFF)
CODEBG=RGBColor(0x2B,0x2B,0x2B); CODEFG=RGBColor(0xF5,0xF5,0xF5); GREEN=RGBColor(0x9E,0xC5,0x7A)
HEAD="Arial Black"; BODYF="Arial"; MONO="Consolas"

ASSETS=os.path.join(os.path.dirname(os.path.abspath(__file__)),"assets")
LOGO_DARK=os.path.join(ASSETS,"logo-dark.png"); LOGO_WHITE=os.path.join(ASSETS,"logo-white.png")
SW=Inches(13.333); SH=Inches(7.5)

def _bg(s,c): s.background.fill.solid(); s.background.fill.fore_color.rgb=c
def _box(s,l,t,w,h):
    tb=s.shapes.add_textbox(l,t,w,h); tf=tb.text_frame; tf.word_wrap=True
    tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0; return tb,tf
def _run(p,text,size,color,font=BODYF,bold=False,italic=False):
    r=p.add_run(); r.text=text; r.font.size=Pt(size); r.font.color.rgb=color; r.font.name=font; r.font.bold=bold; r.font.italic=italic; return r
def _logo(s,white=False,top_left=True):
    path=LOGO_WHITE if white else LOGO_DARK
    if not os.path.exists(path): return
    w=Inches(2.6); h=w*(193/1407)
    if top_left: s.shapes.add_picture(path, Inches(0.5), Inches(0.42), width=w, height=h)
    else: s.shapes.add_picture(path, SW-w-Inches(0.5), SH-h-Inches(0.45), width=w, height=h)
def cover(prs,kicker,title,subtitle,footer):
    s=prs.slides.add_slide(prs.slide_layouts[6]); _bg(s,BG); _logo(s)
    _,tf=_box(s,Inches(0.7),Inches(2.3),Inches(11.6),Inches(3.6))
    _run(tf.paragraphs[0],kicker.upper(),15,RED,HEAD,bold=True)
    p2=tf.add_paragraph(); p2.space_before=Pt(10); _run(p2,title,42,DARK,HEAD,bold=True)
    p3=tf.add_paragraph(); p3.space_before=Pt(14); _run(p3,subtitle,20,BODY,BODYF)
    _,ff=_box(s,Inches(0.7),SH-Inches(0.85),Inches(12),Inches(0.5)); _run(ff.paragraphs[0],footer,13,BODY,BODYF,italic=True); return s
def divider(prs,kicker,title,note=None):
    s=prs.slides.add_slide(prs.slide_layouts[6]); _bg(s,RED_DARK)
    _,tf=_box(s,Inches(0.9),Inches(2.5),Inches(11.3),Inches(3.0))
    if kicker: _run(tf.paragraphs[0],kicker.upper(),16,WHITE,HEAD,bold=True)
    tp=tf.add_paragraph() if kicker else tf.paragraphs[0]; tp.space_before=Pt(8); _run(tp,title,36,WHITE,HEAD,bold=True)
    if note:
        n=tf.add_paragraph(); n.space_before=Pt(14); _run(n,note,18,WHITE,BODYF)
    _logo(s,white=True,top_left=False); return s
def content(prs,title,bullets,subtitle=None):
    s=prs.slides.add_slide(prs.slide_layouts[6]); _bg(s,BG)
    _,tf=_box(s,Inches(0.7),Inches(0.5),Inches(11.9),Inches(1.3)); _run(tf.paragraphs[0],title,30,DARK,HEAD,bold=True)
    if subtitle:
        sp=tf.add_paragraph(); sp.space_before=Pt(4); _run(sp,subtitle,15,RED,BODYF,italic=True)
    top=Inches(1.95) if subtitle else Inches(1.7)
    _,bf=_box(s,Inches(0.7),top,Inches(11.9),Inches(5.3)); first=True
    for b in bullets:
        text,lvl=(b if isinstance(b,tuple) else (b,0))
        p=bf.paragraphs[0] if first else bf.add_paragraph(); first=False; p.space_after=Pt(7)
        if lvl==0: _run(p,"▶  ",12,RED,BODYF,bold=True); _run(p,text,16,DARK,BODYF)
        else: p.level=1; _run(p,"–  ",12,RED,BODYF); _run(p,text,14,BODY,BODYF)
    return s
def code_slide(prs,title,intro,code_lines,note=None):
    s=prs.slides.add_slide(prs.slide_layouts[6]); _bg(s,BG)
    panel=s.shapes.add_shape(MSO_SHAPE.RECTANGLE, SW-Inches(3.4),0,Inches(3.4),SH)
    panel.fill.solid(); panel.fill.fore_color.rgb=RED; panel.line.fill.background(); panel.shadow.inherit=False
    _,pf=_box(s,SW-Inches(3.15),Inches(0.7),Inches(2.7),Inches(2.0))
    _run(pf.paragraphs[0],"LIVE",26,WHITE,HEAD,bold=True); pp=pf.add_paragraph(); _run(pp,"DEMO",26,WHITE,HEAD,bold=True)
    _,tf=_box(s,Inches(0.6),Inches(0.45),Inches(9.0),Inches(1.2)); _run(tf.paragraphs[0],title,25,RED,HEAD,bold=True)
    if intro:
        ip=tf.add_paragraph(); ip.space_before=Pt(3); _run(ip,intro,14.5,BODY,BODYF,italic=True)
    ch=Inches(4.9) if not note else Inches(4.4)
    card=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6),Inches(1.75),Inches(9.0),ch)
    card.fill.solid(); card.fill.fore_color.rgb=CODEBG; card.line.fill.background(); card.shadow.inherit=False
    ctf=card.text_frame; ctf.word_wrap=True; ctf.vertical_anchor=MSO_ANCHOR.TOP
    ctf.margin_left=Inches(0.28); ctf.margin_right=Inches(0.2); ctf.margin_top=Inches(0.2); ctf.margin_bottom=Inches(0.12)
    for i,line in enumerate(code_lines):
        p=ctf.paragraphs[0] if i==0 else ctf.add_paragraph(); p.space_after=Pt(1); p.alignment=PP_ALIGN.LEFT
        color=GREEN if line.strip().startswith("#") else CODEFG
        _run(p,line if line else " ",13,color,MONO)
    if note:
        _,nf=_box(s,Inches(0.6),Inches(6.35),Inches(9.0),Inches(0.85)); _run(nf.paragraphs[0],"▶  "+note,13,RED,BODYF,bold=True)
    return s
def two_col(prs,title,lh,li,rh,ri,subtitle=None):
    s=prs.slides.add_slide(prs.slide_layouts[6]); _bg(s,BG)
    _,tf=_box(s,Inches(0.7),Inches(0.5),Inches(11.9),Inches(1.1)); _run(tf.paragraphs[0],title,30,DARK,HEAD,bold=True)
    if subtitle:
        sp=tf.add_paragraph(); sp.space_before=Pt(3); _run(sp,subtitle,15,RED,BODYF,italic=True)
    def col(l,head,items):
        card=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,l,Inches(2.0),Inches(5.7),Inches(4.8))
        card.fill.solid(); card.fill.fore_color.rgb=WHITE; card.line.color.rgb=RED; card.line.width=Pt(1.5); card.shadow.inherit=False
        _,cf=_box(s,l+Inches(0.35),Inches(2.25),Inches(5.05),Inches(4.3)); _run(cf.paragraphs[0],head,18,RED,HEAD,bold=True)
        for it in items:
            p=cf.add_paragraph(); p.space_before=Pt(7); _run(p,"▶  ",11,RED,BODYF,bold=True); _run(p,it,14,DARK,BODYF)
    col(Inches(0.7),lh,li); col(Inches(7.0),rh,ri); return s
def stat_slide(prs,title,stats,footer=None):
    s=prs.slides.add_slide(prs.slide_layouts[6]); _bg(s,BG)
    _,tf=_box(s,Inches(0.7),Inches(0.7),Inches(11.9),Inches(1.0)); _run(tf.paragraphs[0],title,29,DARK,HEAD,bold=True)
    n=len(stats); cw=Inches(11.9/n)
    for i,(big,label) in enumerate(stats):
        _,cf=_box(s,Inches(0.7)+cw*i,Inches(2.7),cw-Inches(0.3),Inches(2.8))
        bsize=58 if len(big)<=4 else (38 if len(big)<=6 else 30)
        _run(cf.paragraphs[0],big,bsize,RED,HEAD,bold=True); lp=cf.add_paragraph(); lp.space_before=Pt(6); _run(lp,label,15,BODY,BODYF)
    if footer:
        _,ff=_box(s,Inches(0.7),Inches(6.0),Inches(11.9),Inches(0.9)); _run(ff.paragraphs[0],footer,15,DARK,BODYF,italic=True)
    return s
def new_deck():
    prs=Presentation(); prs.slide_width=SW; prs.slide_height=SH; return prs
def add_notes(slide,text): slide.notes_slide.notes_text_frame.text=text
