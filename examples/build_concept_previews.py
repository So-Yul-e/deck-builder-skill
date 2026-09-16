#!/usr/bin/env python3
"""Generate equal-content cover/body/data candidates for visual selection."""
from pathlib import Path
import argparse
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "deck" / "scripts"))
from deck import Deck, CONCEPTS


def build(output, image=None):
    output.mkdir(parents=True, exist_ok=True)
    for concept in CONCEPTS:
        d = Deck(concept=concept, footer="컨셉 비교 · 예시 데이터")
        d.cover("선택을 더 쉽게", "같은 내용, 다른 표현", "디자인 컨셉 미리보기", image=image)
        d.cards("선명하게 전달합니다", [("메시지", "선명", "한 가지 주장"),
                                         ("근거", "명확", "수치로 설명"),
                                         ("화면", "직관", "실물로 전달")])
        d.chart("응답 시간이 줄었습니다", [("이전", 80), ("이후", 45)], kind="bar",
                lead="컨셉 비교를 위한 예시 데이터입니다.")
        target = output / f"concept-{concept}.pptx"
        d.save(target)
        print(target)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=ROOT / "examples" / "output" / "concepts")
    parser.add_argument("--image", type=Path, help="Optional real product image shared by all candidates")
    args = parser.parse_args()
    build(args.output_dir, args.image)
