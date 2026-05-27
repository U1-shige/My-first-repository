#!/usr/bin/env python3
"""
広建フォトコンテスト PPTX リデザイン
テーマ: ダーク + 変形(Morph)トランジション
変形の使い方: タイトルスライドのサムネイルが各賞スライドでフルスクリーンに拡大
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from lxml import etree

# ===== サイズ・色定数 =====
SW = Emu(12192000)   # 16:9 widescreen width (13.33")
SH = Emu(6858000)    # height (7.5")

BG    = RGBColor(0x0d, 0x0d, 0x14)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY = RGBColor(0xCC, 0xCC, 0xCC)

AWARDS = [
    ("最優秀賞", "次代を架けるー若き技術者の視点",             "/tmp/slides/photo1_q65.jpg"),
    ("優秀賞",   "潮騒が語る戦後80年",                         "/tmp/slides/photo2_q65.jpg"),
    ("優秀賞",   "「海の女王」広島港へ",                        "/tmp/slides/photo3_q65.jpg"),
    ("審査員賞", "秋の尾道水道に映る、双子斜張橋のライトアップ", "/tmp/slides/photo4_q65.jpg"),
    ("審査員賞", "船も車も行きかう音頭の瀬戸",                  "/tmp/slides/photo5_q65.jpg"),
]

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"

# ===== ヘルパー =====

def set_bg(slide, rgb):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb

def add_pic(slide, path, name, x, y, w, h):
    pic = slide.shapes.add_picture(path, x, y, w, h)
    pic.name = name
    return pic

def add_txt(slide, text, x, y, w, h, pt, color, bold=False,
            align=PP_ALIGN.LEFT, name=None):
    tx = slide.shapes.add_textbox(x, y, w, h)
    if name:
        tx.name = name
    tf = tx.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(pt)
    run.font.color.rgb = color
    run.font.bold = bold
    return tx

def add_dark_overlay(slide):
    """下部にフェードする半透明黒グラデーション矩形を追加"""
    h = Emu(2600000)
    y = SH - h

    # XML で直接 sp 要素を作成
    sp_xml = f'''<p:sp xmlns:p="{NS_P}" xmlns:a="{NS_A}">
  <p:nvSpPr>
    <p:cNvPr id="99" name="overlay"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="0" y="{y}"/>
      <a:ext cx="{SW}" cy="{h}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:gradFill>
      <a:gsLst>
        <a:gs pos="0">
          <a:srgbClr val="000000"><a:alpha val="0"/></a:srgbClr>
        </a:gs>
        <a:gs pos="100000">
          <a:srgbClr val="000000"><a:alpha val="88000"/></a:srgbClr>
        </a:gs>
      </a:gsLst>
      <a:lin ang="5400000" scaled="0"/>
    </a:gradFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/><a:lstStyle/>
    <a:p/>
  </p:txBody>
</p:sp>'''
    sp_el = etree.fromstring(sp_xml)
    slide.shapes._spTree.append(sp_el)

def add_morph(slide):
    """変形(Morph)トランジションをスライドに追加"""
    trans_xml = f'''<p:transition
        xmlns:p="{NS_P}"
        xmlns:p14="{NS_P14}"
        spd="slow" p14:dur="700">
      <p14:morph variation="byObject"/>
    </p:transition>'''
    slide._element.append(etree.fromstring(trans_xml))

# ===== プレゼンテーション構築 =====

prs = Presentation()
prs.slide_width  = SW
prs.slide_height = SH
blank = prs.slide_layouts[6]

# ---- スライド1: タイトル ----
s1 = prs.slides.add_slide(blank)
set_bg(s1, BG)

# サムネイルグリッド: 上段3枚 / 下段2枚（中央寄せ）
TW = Emu(3600000)
TH = Emu(2025000)   # 16:9比率
GAP = Emu(96000)

r1y = Emu(200000)
r1xs = [Emu(0), Emu(SW//2 - TW//2), SW - TW]

r2y = r1y + TH + GAP
r2xs = [Emu(SW//2 - TW - GAP//2), Emu(SW//2 + GAP//2)]

positions = [
    (r1xs[0], r1y), (r1xs[1], r1y), (r1xs[2], r1y),
    (r2xs[0], r2y), (r2xs[1], r2y),
]

for i, ((x, y), (award, title, path)) in enumerate(zip(positions, AWARDS)):
    add_pic(s1, path, f"!!photo{i+1}", x, y, TW, TH)

# タイトルテキスト（下部）
add_txt(s1, "広建フォトコンテスト",
        Emu(0), SH - Emu(1500000), SW, Emu(1000000),
        40, WHITE, bold=True, align=PP_ALIGN.CENTER)
add_txt(s1, "受賞作品紹介",
        Emu(0), SH - Emu(700000), SW, Emu(600000),
        22, LGRAY, align=PP_ALIGN.CENTER)

# ---- スライド2〜6: 各賞スライド ----
for i, (award, title, path) in enumerate(AWARDS):
    s = prs.slides.add_slide(blank)

    # 写真フルスクリーン（!!photo{n} で Morph が同オブジェクトと認識）
    add_pic(s, path, f"!!photo{i+1}", 0, 0, SW, SH)

    # 下部グラデーションオーバーレイ
    add_dark_overlay(s)

    # 賞名
    add_txt(s, award,
            Emu(400000), SH - Emu(1700000),
            SW - Emu(800000), Emu(900000),
            38, WHITE, bold=True, name="!!award_name")

    # 作品タイトル
    add_txt(s, title,
            Emu(400000), SH - Emu(900000),
            SW - Emu(800000), Emu(700000),
            22, LGRAY, name="!!photo_title")

    # 変形トランジション
    add_morph(s)

out = "/tmp/R8フォトコン資料_v3.pptx"
prs.save(out)

import os
print(f"✅ 保存: {out}")
print(f"   サイズ: {os.path.getsize(out)//1024} KB")
print(f"   スライド数: {len(prs.slides)}")
