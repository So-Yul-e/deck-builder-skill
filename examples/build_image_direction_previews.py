#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Inherit the same six-slide story into three image/composition directions."""
import argparse
from pathlib import Path
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches
from add_topic_imagery import build, place

PALETTES={
    'photo':dict(ink='222222',accent='B83222'),
    'collage':dict(ink='222222',accent='B83222'),
    'screenprint':dict(ink='15364C',accent='A47717'),
}


def rgb(value): return RGBColor.from_string(value)


def build_directions(baseline,assets,out):
    out.mkdir(parents=True,exist_ok=True)
    for direction,image in assets.items():
        target=out/f'{direction}-candidate.pptx'
        build(baseline,image,target)
        p=Presentation(target)
        color=rgb(PALETTES[direction]['ink'])
        accent=rgb(PALETTES[direction]['accent'])
        # Data is inherited, with only the illustrative brand accent changed.
        if direction=='screenprint':
            for slide in p.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if run.font.color.type is not None and str(run.font.color.rgb)=='B83222':
                                    run.font.color.rgb=accent
                    if shape.shape_type != 13 and shape.fill.type is not None:
                        try:
                            if str(shape.fill.fore_color.rgb)=='B83222': shape.fill.fore_color.rgb=accent
                        except (AttributeError,TypeError): pass
        if direction=='collage':
            cover=p.slides[0]
            for s in cover.shapes:
                if s.has_text_frame and s.text:
                    if '같은 틀에' in s.text: place(s,.75,1.45,5.0,3.4,36,color)
                    else:
                        for para in s.text_frame.paragraphs:
                            for r in para.runs: r.font.color.rgb=color
            body=p.slides[1]
            photo=next(s for s in body.shapes if s.shape_type==13)
            place(photo,.75,4.7,11.8,2.0)
            # Match the new native picture crop to the horizontal image band.
            from PIL import Image
            with Image.open(image) as im: iw,ih=im.size
            target_ratio=11.8/2.0; source_ratio=iw/ih
            photo.crop_left=photo.crop_right=0
            photo.crop_top=photo.crop_bottom=max(0,(1-source_ratio/target_ratio)/2)
            for s in body.shapes:
                if not s.has_text_frame or not s.text:
                    if s.shape_type!=13 and s.height>Inches(1): place(s,6.5,2.25,.035,1.2)
                    continue
                val=s.text
                if '색을 바꿔도' in val: place(s,.75,.55,11.8,1.5,32)
                elif val=='입력': place(s,.75,2.3,2,.4,13)
                elif val=='표현': place(s,7.1,2.3,3,.4,13)
                elif val=='절차 / 비교 / 실물': place(s,.75,2.8,5.4,.6,23)
                elif val=='흐름 / 대조 / 화면': place(s,7.1,2.8,5.4,.6,23)
                elif '절차는 흐름으로' in val: place(s,.75,3.65,11.8,.8,18)
                elif val=='2': place(s,12.05,7.02,.5,.25,10)
        elif direction=='screenprint':
            body=p.slides[1]
            photo=next(s for s in body.shapes if s.shape_type==13)
            photo.left=0;photo.width=Inches(5.1)
            from PIL import Image
            with Image.open(image) as im: iw,ih=im.size
            photo.crop_left=max(0,1-(5.1/7.5)/(iw/ih))
            photo.crop_right=photo.crop_top=photo.crop_bottom=0
            for s in body.shapes:
                if not s.has_text_frame or not s.text:
                    if s.shape_type!=13 and s.height>Inches(1): place(s,8.7,2.8,.035,2.65)
                    continue
                val=s.text
                if '색을 바꿔도' in val: place(s,5.7,.65,6.8,1.7,28)
                elif val=='입력': place(s,5.7,3,2,.4,13)
                elif val=='표현': place(s,9.1,3,3,.4,13)
                elif val=='절차 / 비교 / 실물': place(s,5.7,3.6,2.8,1.7,20)
                elif val=='흐름 / 대조 / 화면': place(s,9.1,3.6,3.4,1.7,20)
                elif '절차는 흐름으로' in val: place(s,5.7,5.8,6.8,1.0,17)
                elif val=='2': place(s,12.05,7.02,.5,.25,10)
                else: place(s,5.7,7.03,6.1,.23,10)
        p.save(target)
        print(target)


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--baseline',type=Path,required=True)
    parser.add_argument('--photo',type=Path,required=True)
    parser.add_argument('--collage',type=Path,required=True)
    parser.add_argument('--screenprint',type=Path,required=True)
    parser.add_argument('--output-dir',type=Path,required=True)
    a=parser.parse_args()
    build_directions(a.baseline,{'photo':a.photo,'collage':a.collage,'screenprint':a.screenprint},a.output_dir)
