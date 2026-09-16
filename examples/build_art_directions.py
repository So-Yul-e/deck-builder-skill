#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Subject-led, editable slide studies using this repository's real catalog."""
from pathlib import Path
import argparse
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'deck' / 'scripts'))
from deck import Deck
from build_catalog import PAGE_TYPES

# One named palette per brief: correction ink, type specimen, construction drawing.
PALETTES = {
    'working-notes': dict(deep='222222', key='B83222', sub='555555', tint='F1F1EF', pale='D6D6D2', rule='DADAD6'),
    'type-specimen': dict(deep='171717', key='F0C83E', sub='777777', tint='F5F5F2', pale='D6D6D2', rule='DADAD6'),
    'construction': dict(deep='123758', key='215E91', sub='51768E', tint='EAF1F5', pale='C6D9E7', rule='C6D9E7'),
}
GROUPS = {
    '도입·주장': {'cover','section','statement','quote'},
    '구조·설명': {'bullets','cards','table','deflist','tree'},
    '흐름·데이터': {'flow','timeline','bar','column','donut','progress','matrix'},
    '증빙·판단': {'shots','compare','stat','gate'},
}
assert set.union(*GROUPS.values()) == set(PAGE_TYPES)
COUNTS = [(name, sum(p in types for p in PAGE_TYPES)) for name, types in GROUPS.items()]
SOURCE = '출처: examples/build_catalog.py · PAGE_TYPES / 편집상 분류'
CLAIM = '같은 틀에 넣으면\n다른 내용도 같아 보인다.'
BODY = '색을 바꿔도,\n내용의 관계는 바뀌지 않는다.'
EXPLANATION = '절차는 흐름으로, 비교는 대조로, 실물은 화면으로 보여준다.'


def text(value, x, y, w, h, size=19, color='ink', bold=False):
    return dict(kind='text', text=value, x=x, y=y, w=w, h=h, size=size, color=color, bold=bold)


def rect(x, y, w, h, color='rule'):
    return dict(kind='rect', x=x, y=y, w=w, h=h, color=color)


def source(page):
    return [text(SOURCE if page == 3 else '구성 제안: 덱 스킬 내용 기반 디자인 연구', .75, 7.03, 11.6, .23, 10), text(str(page),12.05,7.02,.5,.25,10)]


def cover(d, direction):
    if direction == 'working-notes':
        elements = [text('덱 스킬 / 디자인 작업 노트',.75,.55,8,.35,13),
                    text(CLAIM,.75,1.5,11.8,2.5,55,bold=True),
                    rect(.8,4.5,7.4,.045,'key'),
                    text('정렬은 유지한다.\n표현은 내용에서 찾는다.',8.2,5.0,4.1,1.2,23,color='key'),
                    text('고정 컨셉에서 내용 기반 구성으로',.75,6.9,9,.32,12)]
        d.compose(elements)
    elif direction == 'type-specimen':
        d.compose([text('덱 스킬 / 활자 표본',.75,.55,8,.35,13,'white'),
                   text(CLAIM,.75,1.15,11.8,3.0,65,'white',True),
                   text('정렬은 유지한다.\n표현은 내용에서 찾는다.',.75,5.55,9,1.0,24,'key'),
                   text('고정 컨셉에서 내용 기반 구성으로',.75,6.95,9,.32,12,'white')],background='deep')
    else:
        d.compose([text('덱 스킬 / 레이아웃 설계도',.75,.55,8,.35,13,'key'),
                   rect(.75,1.4,.035,4.6,'key'), rect(12.5,1.4,.035,4.6,'key'),
                   text(CLAIM,1.05,1.75,11.0,2.2,48,'deep',True),
                   text('정렬은 유지한다.\n표현은 내용에서 찾는다.',1.05,4.8,9,1.0,24,'deep'),
                   text('고정 컨셉에서 내용 기반 구성으로',.75,6.95,9,.32,12)],background='tint')


def body(d, direction):
    elements=[]
    if direction == 'working-notes':
        elements=[text(BODY,.75,.65,11.8,1.65,38,bold=True),
                  text('입력',.75,3.0,2,.4,13,'mute'),text('절차 / 비교 / 실물',.75,3.6,5.5,.6,27,bold=True),
                  rect(6.65,2.8,.035,2.65,'key'),
                  text('표현',7.1,3.0,3,.4,13,'mute'),text('흐름 / 대조 / 화면',7.1,3.6,5.4,.6,27,'key',True),
                  text(EXPLANATION,.75,6.0,11.7,.8,21)]
    elif direction == 'type-specimen':
        elements=[text('색을 바꿔도,',.75,.8,11.8,1.1,48,bold=True),
                  text('내용의 관계는 바뀌지 않는다.',.75,2.0,11.8,.9,32,bold=True),
                  rect(.75,3.65,11.8,.06,'ink'),
                  text('절차\n흐름',.75,4.1,3.5,1.4,27,bold=True),
                  text('비교\n대조',5,4.1,3.5,1.4,27,bold=True),
                  text('실물\n화면',9.4,4.1,3.2,1.4,27,bold=True),
                  text(EXPLANATION,.75,6.2,11.7,.6,19)]
    else:
        elements=[text(BODY,.75,.65,11.8,1.5,34,'deep',True)]
        for i,(a,b) in enumerate([('절차','흐름'),('비교','대조'),('실물','화면')]):
            y=2.6+i*.94
            elements += [text(a,1.1,y,3,.6,25,'deep'),rect(4.7,y+.24,3.3,.025,'key'),
                         text(b,8.5,y,3.5,.6,25,'key',True)]
        elements += [text(EXPLANATION,.75,6.2,11.7,.6,19,'deep')]
    d.compose(elements+source(2))


def data(d,direction):
    elements=[text('20장의 대표 슬라이드, 네 가지 역할',.75,.6,11.8,.9,32,bold=True),
              text('각 형태를 편집상 역할로 묶은 목록이다. 성과 수치가 아니다.',.75,1.65,11.8,.6,17,'mute')]
    if direction == 'type-specimen':
        for i,(name,count) in enumerate(COUNTS):
            x=.75+i*3.02
            elements += [rect(x,5.6-count*.38,2.5,count*.38,'key'),
                         text(str(count),x,5.6-count*.38-.9,2.5,.85,42,bold=True),
                         text(name,x,5.9,2.65,.5,17)]
    else:
        for i,(name,count) in enumerate(COUNTS):
            y=2.65+i*.78
            elements += [text(name,.75,y,3,.45,18),
                         rect(4.0,y+.02,count*.8,.32,'key'),
                         text(str(count),4.2+count*.8,y,1,.48,21,bold=True)]
    elements += [text('단위: 슬라이드 수 / 합계 20장',.75,6.55,11.7,.28,11)]
    d.compose(elements+source(3))


def build(out):
    out.mkdir(parents=True,exist_ok=True)
    for direction,palette in PALETTES.items():
        d=Deck(palette=palette)
        cover(d,direction);body(d,direction);data(d,direction)
        d.save(out/f'{direction}.pptx')
    d=Deck(palette=PALETTES['working-notes'])
    cover(d,'working-notes');body(d,'working-notes');data(d,'working-notes')
    d.compose([
        text('나란히 놓기 전에,\n관계를 먼저 본다.',.75,.7,11.8,1.7,40,bold=True),
        text('flow()',.75,3.0,4,1.0,43,'key',True),
        text('발견',.75,4.5,2,.5,25),rect(3.0,4.75,1.6,.025,'ink'),
        text('선택',5.0,4.5,2,.5,25),rect(7.3,4.75,1.6,.025,'ink'),
        text('생성',9.3,4.5,2,.5,25),
        text('이름을 나열한 카드가 아니라, 단계 사이의 이동을 보여준다.',.75,6.15,11.8,.7,19),
    ]+source(4))
    d.compose([text('모든 장에 같은 헤더가\n필요한 것은 아니다.',.75,.8,11.8,1.7,40,bold=True),
               text('표지',.75,3.4,3,.5,24,'key',True),text('한 문장을 크게 남긴다.',4.8,3.4,7.4,.5,22),
               text('근거',.75,4.4,3,.5,24,'key',True),text('라벨·단위·출처를 빠짐없이 둔다.',4.8,4.4,7.4,.6,22),
               text('설명',.75,5.4,3,.5,24,'key',True),text('관계에 맞는 구성을 만든다.',4.8,5.4,7.4,.5,22)]+source(5))
    d.compose([text('표현은 달라도,\n내용은 남아야 한다.',.75,1.1,11.8,2.1,48,bold=True),
               rect(.75,4.05,11.8,.035,'key'),
               text('줄이지 않고 배치한다.\n넘치면 나누고, 렌더해서 확인한다.',.75,4.7,11.8,1.6,28),
               text('덱 스킬 / 내용 기반 구성 연구',.75,6.95,11.8,.3,12)])
    d.save(out/'content-led-deck.pptx')
    (out/'brief.md').write_text('# 덱 스킬: 내용의 모양을 고르다\n\n청중: 덱 스킬을 보강하는 사용자. 목적: 반복 틀에서 내용 관계 기반 구성으로 바뀌는 방식 설명.\n\n선택: 작업 노트. 교정색과 열린 배치는 설계 도구를 설명하는 주제에 맞는다. 활자 표본은 발표에 강하지만 상세 설명에 좁고, 설계도는 흐름에 유리하지만 보조선이 늘어날 수 있다. 세 방향은 같은 주장과 실제 20종 목록을 사용한다.\n\n슬라이드 지도: 주장(큰 문장) → 변환(입력/표현) → 근거(역할별 목록) → 절차(연결) → 기준(비대칭 설명) → 결론(문장). 숫자 그룹은 편집상 분류이며 성과 자료가 아니다.\n\n검증 기준: 텍스트 보존, 라벨·단위·출처, 잘림 없음, 제목/본문/메타 위계. AI 여부를 증명하는 자료가 아니라 디자인 방향 실험이다.\n',encoding='utf-8')
    print(out)


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir',type=Path,default=ROOT/'examples'/'output'/'art-directions')
    build(parser.parse_args().output_dir)
