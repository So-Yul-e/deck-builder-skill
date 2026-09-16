<h1 align="center">Deck Builder Skill</h1>

<p align="center">
  <a href="README.md">English</a> · <b>한국어</b>
</p>

<p align="center">
  구조화된 콘텐츠를 세련되고 편집 가능한 PowerPoint 덱으로 만듭니다.<br/>
  적합한 슬라이드 형식 선택부터 PPTX 생성, PDF 내보내기, 폰트 임베딩 검증까지 한 흐름으로 처리합니다.
</p>

<p align="center">
  <a href="https://github.com/So-Yul-e/deck-builder-skill/actions/workflows/ci.yml"><img src="https://github.com/So-Yul-e/deck-builder-skill/actions/workflows/ci.yml/badge.svg" alt="CI 상태" /></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/code_license-MIT-blue.svg" alt="코드 라이선스: MIT" /></a>
</p>

<p align="center">
  <b>macOS 검증 완료</b> · <b>Windows 어댑터 beta</b> · Claude와 Codex 호환
</p>

<p align="center">
  <img src="examples/output/deck-builder-catalog-contact-sheet.png" alt="대표 슬라이드 20종 카탈로그" width="920" />
</p>

<p align="center">
  <a href="examples/output/deck-builder-catalog.pptx">편집 가능한 20장 PPTX</a> ·
  <a href="examples/output/deck-builder-catalog.pdf">검증된 20페이지 PDF</a> ·
  <a href="examples/output/deck-builder-catalog-contact-sheet.png">전체 카탈로그 이미지</a>
</p>

---

## 무엇을 만들 수 있나

새로운 덱을 만들거나 기존 `.pptx`를 이미지로 평면화하지 않고 정리할 수 있습니다.

| 덱 단위 용도 | 예시 |
|---|---|
| 전략·제안 | 사업계획, 제품 전략, 선택지 비교, Go/No-Go 제안 |
| 보고·로드맵 | 임원 보고, 프로젝트 현황, 마일스톤, 실행 로드맵 |
| 데이터 스토리텔링 | KPI 리뷰, 설문 결과, 구성비, 진행률, 포지셔닝 매트릭스 |
| 제품·사례 소개 | 기능 내러티브, 전후 비교, 스크린샷, 사용자 목소리, 성과 |
| 설명 자료 | 프로세스 가이드, 계층 구조, 용어 정의, 교육·워크숍 자료 |
| 기존 덱 정리 | 제목 위계, 내부 여백, 불필요한 테두리, 시각 일관성 |

저장소의 카탈로그는 **대표 슬라이드 20종**을 포함합니다. 18개 builder
메서드를 사용하며, `chart()`가 세 가지 차트 형식을 만듭니다.

| 콘텐츠 관계 | 슬라이드 형식 |
|---|---|
| 내러티브·강조 | 표지, 섹션, 핵심 문장, 인용, 대표 수치 |
| 구조·의사결정 | 불릿, 카드, 표, 정의 목록, 트리, 플로우, 타임라인, 전후 비교, 게이트 |
| 데이터·현황 | 가로 막대, 세로 막대, 도넛, 진행률, 매트릭스 |
| 시각 증거 | 스크린샷과 결과 산출물 |

`bullets()`는 기본값이 아니라 최후의 형식입니다. 프로세스는 flow,
날짜는 timeline, 구성비는 donut, 통과 기준은 gate로 표현합니다.

## 왜 만들었나

LLM이 만드는 덱은 모든 아이디어를 불릿으로 바꾸기 쉽습니다. 그 과정에서
타이포그래피와 시각 위계가 흐트러지고, PDF 변환 시 폰트가 조용히 대체됩니다.
편집은 가능하지만 발표에 바로 쓰기 어려운 결과가 남습니다.

이 스킬은 덱 제작을 반복 가능한 계약으로 바꿉니다. 콘텐츠의 관계를 파악하고,
슬라이드마다 하나의 주 builder를 선택해 편집 가능한 PPTX를 만든 뒤 렌더링합니다.
필수 폰트가 임베딩되지 않았다면 결과를 전달하지 않고 실패로 처리합니다.

## 빠른 시작

### 1. 스킬 설치

Claude용 macOS 설치:

```bash
git clone https://github.com/So-Yul-e/deck-builder-skill.git
cd deck-builder-skill
mkdir -p ~/.claude/skills
ln -s "$(pwd)/deck" ~/.claude/skills/deck
```

Claude용 Windows PowerShell 설치:

```powershell
git clone https://github.com/So-Yul-e/deck-builder-skill.git
$repo = (Resolve-Path ".\deck-builder-skill").Path
$target = "$HOME\.claude\skills\deck"
New-Item -ItemType Directory -Force $target | Out-Null
Copy-Item -Recurse -Force "$repo\deck\*" $target
```

Codex도 같은 `deck/` 폴더를 사용합니다. macOS에서는
`~/.codex/skills/deck`, Windows에서는 `$HOME\.codex\skills\deck`에
링크하거나 복사하세요. 공통 작업 계약은 [`deck/SKILL.md`](deck/SKILL.md),
Codex 표시 메타데이터는 [`deck/agents/openai.yaml`](deck/agents/openai.yaml)에 있습니다.

### 2. 런타임 점검

점검 명령은 설치 없이 누락된 항목만 알려줍니다. Python·폰트·렌더러 변경
내용을 확인한 뒤에만 `--install`을 실행하세요.

macOS:

```bash
bash ~/.claude/skills/deck/scripts/deps-macos.sh --check
bash ~/.claude/skills/deck/scripts/deps-macos.sh --install
```

Windows beta:

```powershell
powershell -ExecutionPolicy Bypass -File "$HOME\.claude\skills\deck\scripts\deps-windows.ps1" -Check
powershell -ExecutionPolicy Bypass -File "$HOME\.claude\skills\deck\scripts\deps-windows.ps1" -Install
```

### 3. 에이전트에게 덱 요청

요청 예시:

```text
$deck을 사용해서 이 사업계획을 10장짜리 투자 제안 덱으로 만들어줘.

$deck으로 로드맵, KPI, 릴리즈 게이트가 포함된 임원 프로젝트 보고서를 만들어줘.

$deck으로 이 PPTX의 내용과 박스 위치는 바꾸지 말고 시각적으로 정리해줘.
```

## 이미지 방향 선택과 업데이트

**이미지 생성은 선택 사항입니다.** 사용자가 가진 사진·제품 이미지·스크린샷과
사용 조건을 확인한 기존 이미지를 먼저 활용합니다. 사용자가 생성을 원하거나,
적절한 자료가 없어 생성 방식을 선택한 경우에만 새 이미지를 만듭니다.
이미지가 도움이 되지 않는 주제는 텍스트·도표·다이어그램으로 구성합니다.

주제에 맞는 실제 화면·사진 또는 생성 이미지를 포함한 시안을 2~3개 비교하고,
고른 방향으로 전체 덱을 만들 수 있습니다. 색만 바꾸는 대신 이미지 소재와
배치를 함께 달리하고, 같은 문구와 데이터를 유지합니다. 이미지 생성·검색은
설치한 에이전트의 해당 도구가 필요하며 스킬 자체가 이미지 API를 번들하지는 않습니다.
제작 계약은 [이미지 가이드](deck/references/imagery.md), 시안 재현 코드는
[이미지 방향 생성기](examples/build_image_direction_previews.py)에 있습니다.

자동 업데이트는 **공식 저장소의 깨끗한 main clone에 연결한 설치본**에서
한 번 켤 수 있습니다. 활성화하면 에이전트가 스킬을 사용할 때 확인하며, 원격
확인은 1시간에 한 번입니다. 로컬 수정·갈라진 이력·다른 브랜치·오프라인이면
업데이트를 건너뛰고 덱 제작을 계속합니다.

```bash
python3 deck/scripts/update_skill.py --enable
python3 deck/scripts/update_skill.py --check
python3 deck/scripts/update_skill.py --disable
```

위 명령은 해당 clone에서 실행합니다. Windows는 `py -3` 또는 `python`을
사용하세요. 복사 설치본은 이 기능으로 갱신할 수 없으므로 재설치하거나 clone에
연결해야 합니다. 기존 설치본은 새 실행기를 받는 최초 한 번의 수동 갱신이
필요합니다. 실행 중인 모든 세션에 전역 훅을 설치하지 않습니다.
[업데이트 가이드](deck/references/updates.md)에 동작과 한계를 설명했습니다.

현재 Codex 공식 문서의 사용자 스킬 위치는 `~/.agents/skills`이며 링크된
폴더를 지원합니다. 기존 Codex 환경에서 `~/.codex/skills`를 사용하고 있다면
해당 환경의 호환 경로를 유지할 수 있습니다. 로컬 파일 변경 감지는 GitHub에서
새 코드를 내려받는 동작과 다릅니다.
[공식 스킬 문서](https://learn.chatgpt.com/docs/build-skills)를 확인하세요.

## 동작 방식

```text
원본 콘텐츠
    ↓ 콘텐츠 관계 판정
슬라이드별 주 builder 하나
    ↓
편집 가능한 PPTX
    ↓ macOS 또는 Windows 렌더 어댑터
PDF
    ↓ 폰트 임베딩·fallback 검증
전달 가능한 결과 또는 fail-closed 거부
```

모델은 메시지와 콘텐츠 관계를 판단합니다. Builder는 여백, 타이포그래피,
시각 위계, 도형 구성을 책임집니다. 이 역할 분리는 PPT 편집 가능성을 유지하면서
임의적인 레이아웃 결정을 줄입니다.

### 최소 Python 예제

```python
import sys
sys.path.insert(0, "deck/scripts")

from deck import Deck

d = Deck(palette="indigo", footer="Project · Team")
d.cover("Quarterly Review", "What changed and what we do next", "2026 Q3")
d.cards("Top metrics", [
    ("Activation", "42%", "+8pp"),
    ("Latency", "0.8s", "P95"),
])
d.flow("Delivery flow", [
    ("Scope", "decide"),
    ("Build", "generate"),
    ("Verify", "render"),
])
d.save("review.pptx")
```

기본 팔레트는 `indigo`, `navy`, `mono`입니다. 기존 디자인 시스템이 있다면
새 색상을 임의로 만들지 않고 palette dictionary로 연결할 수 있습니다.

## 핵심 판단

차별점은 무엇을 만들었는지뿐 아니라 어떤 실패를 의도적으로 제외했는지에 있습니다.

| 판단 | 버린 선택지 | 영향과 증거 |
|---|---|---|
| 슬라이드 문구보다 형식을 먼저 선택 | 모든 슬라이드를 불릿으로 시작 | Builder 선택표와 제한은 [`deck/SKILL.md`](deck/SKILL.md), 지원 형식은 [20종 카탈로그](examples/build_catalog.py)에 고정했습니다. |
| 하나의 공통 덱 엔진 유지 | macOS와 Windows builder 저장소 분리 | [`deck.py`](deck/scripts/deck.py)를 단일 진실원천으로 두고 의존성·렌더 어댑터만 OS별로 나눠 시각 동작의 드리프트를 막았습니다. |
| Pretendard Regular·Bold 번들 | 대상 PC에 폰트가 있다고 가정 | OS 설치 스크립트가 번들한 OFL 폰트를 사용하고, [`verify_pdf_fonts.py`](deck/scripts/verify_pdf_fonts.py)가 임베딩 누락과 폭이 좁은 fallback 폰트를 거부합니다. 저장소에 약 3MB의 폰트가 추가되는 비용이 있습니다. |
| 렌더 검증을 전달 조건으로 설정 | PPTX 생성 성공을 완료로 간주 | macOS 렌더, PowerPoint/LibreOffice 어댑터, PDF 검사, 커밋된 결과물로 전달 전 clipping과 폰트 대체를 드러냅니다. |

## 현재 검증 결과

| 검사 | 결과 | 증거 |
|---|---|---|
| 자동 계약 테스트 | macOS 로컬 29개 PASS | [`tests/`](tests)의 패키지·라이선스·PDF·구성·업데이트 검사. 기존 CI 결과는 [GitHub Actions](https://github.com/So-Yul-e/deck-builder-skill/actions/workflows/ci.yml) 참조 |
| Builder 범위 | 고유 카탈로그 20페이지 | [`examples/build_catalog.py`](examples/build_catalog.py)가 정확한 페이지 수를 assertion으로 고정 |
| PDF 규격 | 20페이지, 16:9 | [커밋된 PDF](examples/output/deck-builder-catalog.pdf)는 `959.981 × 540 pt` |
| PDF 타이포그래피 | Pretendard Regular/Bold 임베딩, Arial Narrow·STHeiti 거부 | [`verify_pdf_fonts.py`](deck/scripts/verify_pdf_fonts.py)와 [폰트 회귀 테스트](tests/test_pdf_verifier.py) |
| macOS 런타임 | 검증 완료 | 격리된 Python 환경, LibreOffice 경로, PowerPoint fallback, `pdffonts` 사전 점검 |
| Windows 패키지 경로 | CI PASS, 어댑터 beta | `windows-latest`에서 20페이지 PPTX를 재생성하고 [`test_windows_scripts.ps1`](tests/test_windows_scripts.ps1)을 실행 |

저장소에 포함된 검사를 실행하려면:

```bash
python3 -m unittest discover -s tests -p 'test_*.py'
pwsh -NoProfile -File tests/test_windows_scripts.ps1
bash deck/scripts/deps-macos.sh --check
```

macOS에서 카탈로그를 렌더링하고 검증하려면:

```bash
DECK_PY="$(bash deck/scripts/deps-macos.sh --python)"
"$DECK_PY" examples/build_catalog.py examples/output/deck-builder-catalog.pptx
bash deck/scripts/render-macos.sh \
  examples/output/deck-builder-catalog.pptx \
  examples/output/deck-builder-catalog.pdf
```

## 제작자·역할·증거

| 항목 | 내용 |
|---|---|
| 제작자 | 윤소율 (So-Yul) |
| 역할 | 워크플로 설계, 시각 시스템 규칙, 구현, 패키징, 렌더 QA |
| 팀 | 1인 프로젝트, AI를 구현 보조와 독립 검토에 활용 |
| 기간 | 2026-08-04 ~ 2026-08-07 |

이 저장소는 윤소율의 기존 개인용 덱 워크플로를 공개·이식 가능한 스킬로
패키징한 결과입니다. 문제 정의, builder 선택 규칙, 이식성 판단, 합격 기준,
릴리즈 판단은 사람이 소유하고 AI는 구현과 검토를 위한 지렛대로 사용했습니다.

| 수행 축 | 구체적인 증거 |
|---|---|
| 워크플로 기획 | [`deck/SKILL.md`](deck/SKILL.md)에 라우팅, builder 선택, 제한, 렌더 후 전달 계약을 정의했습니다. |
| 시각 시스템 | [`deck.py`](deck/scripts/deck.py)에 타이포그래피, 여백, 팔레트, 위계, 18개 builder를 고정했습니다. |
| 크로스플랫폼 개발 | [`deck/scripts/`](deck/scripts)에 macOS·Windows 의존성 및 렌더 어댑터를 분리했습니다. |
| 품질·릴리즈 | 단위 테스트, PowerShell 계약 검사, CI, 20장 PPTX, 20페이지 PDF, contact sheet를 증거로 커밋했습니다. |

## 알려진 한계와 다음 검증

- **Windows PDF 동등성은 beta입니다.** Windows CI에서 설치·렌더 스크립트
  계약과 PPTX 생성은 확인했지만 실제 Windows PowerPoint와 LibreOffice의
  PDF 결과를 비교하는 검증이 남아 있습니다.
- **Linux는 지원하지 않습니다.** Linux용 의존성 설치나 렌더 어댑터를 제공한다고
  주장하지 않습니다.
- **브랜드를 추측하지 않습니다.** 기업 아이덴티티를 임의로 생성하지 않고 기존
  디자인 토큰을 custom palette로 연결합니다.
- **Polish는 보수적으로 작동합니다.** 위계, 내부 여백, 불필요한 테두리는 조정하지만
  overflow를 만들 수 있는 박스 이동과 본문 크기 변경은 자동화하지 않습니다.

## 프로젝트 구조

```text
deck/                         # 설치 가능한 스킬 본체
  SKILL.md                    # 에이전트 워크플로와 builder 선택 계약
  agents/openai.yaml          # Codex 표시 메타데이터
  scripts/
    deck.py                   # 공통 18-builder 프레젠테이션 엔진
    polish.py                 # 기존 덱의 보수적 정리
    deps-macos.sh             # macOS 런타임 점검·설치
    deps-windows.ps1          # Windows beta 점검·설치
    render-macos.sh           # macOS PDF 변환·검증
    render-windows.ps1        # Windows beta PDF 변환·검증
    verify_pdf_fonts.py       # 폰트 임베딩·fallback 가드
  assets/fonts/               # 번들 Pretendard와 OFL 라이선스
examples/                     # 20종 builder 카탈로그와 결과물
tests/                        # 패키지·PDF·Windows 스크립트 계약
.github/workflows/ci.yml      # Ubuntu·Windows 품질 게이트
THIRD_PARTY_NOTICES.md        # 번들 제3자 자산의 라이선스 경계
```

## 기술 구성

| 영역 | 선택 | 이유 |
|---|---|---|
| PPTX 생성 | Python 3 + `python-pptx` | 결정적인 geometry를 가진 편집 가능한 Office 문서 생성 |
| 이미지 처리 | Pillow | 덱 전체를 평면화하지 않고 raster 증거를 생성·배치 |
| PDF 검사 | `pypdf` + 사용 가능 시 Poppler `pdffonts` | 눈으로만 판단하지 않고 폰트 resource를 검사 |
| macOS 렌더 | LibreOffice 격리 profile, PowerPoint fallback | 반복 가능한 CLI 경로와 네이티브 Office 대안 확보 |
| Windows 렌더 | PowerPoint COM, LibreOffice fallback | 플랫폼의 강한 네이티브 변환 경로를 우선 사용 |
| 자동화 | Ubuntu·Windows GitHub Actions | 공통 코어와 OS 어댑터의 회귀 차단 |

## 라이선스

이 프로젝트에서 작성한 원본 코드와 문서는 [MIT License](LICENSE)로
공개합니다. 다른 사람은 저작권·허가문을 유지하는 조건으로 사용, 수정,
재배포, sublicense, 판매할 수 있습니다.

번들한 Pretendard 폰트 파일에는 MIT가 적용되지 않습니다. 폰트는 원저작자와
Reserved Font Name을 유지한 SIL Open Font License 1.1 적용 대상입니다.
[제3자 고지](THIRD_PARTY_NOTICES.md)와 전체 [`OFL.txt`](deck/assets/fonts/OFL.txt)를
확인하세요.
