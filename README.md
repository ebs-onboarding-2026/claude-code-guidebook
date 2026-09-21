# 클로드 코드 입문자 가이드북

클로드 코드를 처음 쓰는 사람을 위한 한국어 가이드북 원고와, 그 내용을 HTML 슬라이드로
만들어 PPTX로 변환하는 빌더입니다. 2026년 EBS 신규사원 연수과정 실습 결과물입니다.

## 구조

```
클로드코드-입문자-가이드북.md   가이드북 본문 (12장)
클로드코드-입문가이드.md         확장판 원고
클로드코드-가이드북-샘플.html    본문 HTML 렌더링 샘플
deck-sample.html                3장 슬라이드 (1280x720)
deck-sample.pptx                위 HTML을 변환한 결과물
build_pptx.py                   HTML -> PPTX 변환기
.pptx-build/                    변환 중간 산출물 (슬라이드 미리보기 PNG)
```

## PPTX 빌드

```
python build_pptx.py
```

표준 라이브러리만 씁니다. `python-pptx`나 `lxml` 없이 OOXML을 직접 조립해 zip으로
묶는 방식이라 설치가 필요 없습니다. HTML의 1280x720 px 좌표를 그대로 EMU로 환산하므로
슬라이드 배치가 HTML과 일치합니다. 폰트는 윈도우 기본인 맑은 고딕과 Consolas를 씁니다.

> `build_pptx.py`의 `OUT` 경로는 작성 당시 절대 경로로 하드코딩돼 있습니다. 다른 환경에서
> 돌릴 때는 이 값을 바꿔 주세요.
