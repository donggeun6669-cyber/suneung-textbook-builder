# -*- coding: utf-8 -*-
"""
원문 무결성 검사 — 이 프로젝트의 제1원칙("원문을 절대 빼먹지 않는다")을 강제한다.

원리
----
- 조교가 붙여넣은 원문 그대로의 복사본을 각 record 의 `RAW` 필드에 보관한다.
  (RAW 는 분석/요약/수정의 대상이 아니다. 말 그대로 '복붙 원본'.)
- 빌드가 docx 를 만든 뒤, RAW 를 문장·선지 단위(unit)로 쪼개
  **각 unit 이 최종 docx 텍스트 안에 실제로 존재하는지** 대조한다.
- 하나라도 없으면 = 분석 과정에서 원문이 빠지거나 변형된 것 → 치명적 실패로 처리.

왜 이게 동작하나
----------------
docx 에는 PASSAGE/CHOICES 가 글자 그대로 박힌다. 만약 분석 단계에서 PASSAGE 가
한 문장 누락되면, '손대지 않은' RAW 에는 그 문장이 남아 있으므로 대조에서 걸린다.
즉 RAW 가 진실의 기준이고, 분석 결과가 RAW 에서 벗어나는 순간 빌드가 멈춘다.

대소문자·따옴표(' ' " ")·대시(– —)·빈칸(____)·공백 차이는 정규화로 흡수하되,
'문장이 통째로 사라진' 경우는 반드시 잡는다(이쪽으로는 안전하게 과검출).
"""

from __future__ import annotations

import html
import re
import unicodedata
import zipfile
from pathlib import Path


# 따옴표/대시/특수공백을 표준 형태로 통일
_TRANSLATE = {
    "‘": "'", "’": "'", "‛": "'",        # ' ' ‛
    "“": '"', "”": '"', "„": '"',        # " " „
    "–": "-", "—": "-", "‒": "-", "‐": "-", "‑": "-",
    " ": " ", " ": " ", " ": " ", " ": " ",  # 각종 공백
    "…": "...",                                     # … 줄임표
}


# RAW 검사가 멈췄을 때, 조교(또는 모델)가 검사기/RAW 를 고쳐 통과시키려는 사고를 막는 안내.
# 이 검사는 마커(원문자 등)는 이미 무시하므로, 멈췄다면 '진짜 문장'이 빠진 것이다.
_FIX_GUIDE = (
    "※ 해결법(중요): integrity.py(검사기)도 RAW 도 고치지 말 것. "
    "이 검사는 원문자·라벨 같은 표시기호는 이미 무시한다 — 그래도 멈췄다면 "
    "분석 중 '문장/선지'가 docx 에서 통째로 빠진 것이다. "
    "빠진 그 문장을 PASSAGE/CHOICES(또는 SENTENCE_TRIPLES)에 되살린 뒤 다시 빌드하라."
)


def normalize(s: str) -> str:
    """대조용 정규화: NFKC + 따옴표/대시 통일 + 빈칸/공백 축약 + 소문자."""
    s = unicodedata.normalize("NFKC", str(s))
    for a, b in _TRANSLATE.items():
        s = s.replace(a, b)
    s = s.lower()
    s = re.sub(r"_+", " ", s)        # 빈칸(밑줄)은 공백으로
    s = re.sub(r"\s+", " ", s)       # 연속 공백 → 1칸
    return s.strip()


def split_units(text: str) -> list[str]:
    """RAW 를 검사 단위로 분할: 줄바꿈으로 먼저 나누고, 줄 안에서 . ? ! 로 또 나눈다.

    선지(① …)는 보통 줄 단위라 각 선지가 한 unit 이 된다. 문장 분할이
    과해도(약어 등) '포함 여부' 검사라 거짓경보가 나지 않는다 — 더 작게 쪼개질 뿐.
    """
    units: list[str] = []
    for line in str(text).split("\n"):
        line = line.strip()
        if not line:
            continue
        for piece in re.split(r"(?<=[.?!])\s+", line):
            piece = piece.strip()
            if len(piece) >= 2:          # 한두 글자 토막은 무시(노이즈)
                units.append(piece)
    return units


def docx_fulltext(docx_path: str | Path) -> str:
    """docx 의 모든 w:t 텍스트를 뽑아 하나의 문자열로(텍스트박스 내부 포함)."""
    xml = zipfile.ZipFile(str(docx_path)).read("word/document.xml").decode("utf-8")
    parts = re.findall(r"<w:t[^>]*>(.*?)</w:t>", xml, re.S)
    return " ".join(html.unescape(p) for p in parts)


def find_missing(raw: str, haystack_cleaned: str) -> list[str]:
    """RAW unit 중 docx 에 없는 것들을 돌려준다. 비어 있으면 누락 없음.

    ★ 대조는 `_coverage_clean` 기준(표시기호를 양쪽에서 똑같이 제거)으로 한다.
      어법/어휘 지문의 **인라인 원문자(② …)는 docx 에선 떼고 렌더**되는데(build 자동 밑줄),
      옛 버전은 normalize 가 NFKC 로 '②'→'2' 로 바꿔 비교해 *가짜 누락경보*를 냈다.
      이제 양쪽에서 원문자·(A)~(E) 라벨·빈 괄호를 제거하고 '문장 내용'만 대조한다.
    ★ 마커/문장부호만 남는 토막(문삽 '①②③④⑤' 선지줄, 요약 '(A)…(B)' 헤더)은
      '문장'이 아니므로 건너뛴다. **문장이 통째로 빠지면 글자가 남아 여전히 걸린다(안전).**
      → haystack_cleaned 는 `_coverage_clean(docx_fulltext(...))` 를 넘겨야 한다.
    """
    missing = []
    for unit in split_units(raw):
        norm = _coverage_clean(unit)
        if not norm:
            continue
        if not re.search(r"[a-z]", norm):    # 글자 없는 토막=마커/라벨 → 검증대상 아님
            continue
        if norm not in haystack_cleaned:
            missing.append(unit)
    return missing


def check_records(docx_path: str | Path, records: list, raw_key: str = "RAW",
                  id_key: str = "NUMBER") -> list[str]:
    """각 record 의 RAW 가 docx 안에 온전히 들어갔는지 검사.

    반환: 사람이 읽을 수 있는 누락 메시지 리스트(비면 통과).
    """
    blob = _coverage_clean(docx_fulltext(docx_path))   # RAW unit 과 같은 기준으로 정리
    issues: list[str] = []
    for i, rec in enumerate(records):
        raw = rec.get(raw_key, "")
        if not str(raw).strip():
            issues.append(f"record[{i}] ({id_key}={rec.get(id_key, '?')}): "
                          f"'{raw_key}' 가 비어 있어 원문 대조 불가")
            continue
        missing = find_missing(raw, blob)
        if missing:
            where = f"{id_key}={rec.get(id_key, '?')}"
            issues.append(f"[원문 누락] {where}: {len(missing)}개 문장/선지가 docx 에 없음")
            for m in missing:
                snippet = m if len(m) <= 70 else m[:70] + "…"
                issues.append(f"        └ {snippet}")
    if issues:
        issues.append(_FIX_GUIDE)
    return issues


_CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮"


def _coverage_clean(s: str) -> str:
    """문장별 해설 커버리지 대조용 정규화.

    유형별 '표시 기호'를 양쪽(지문 문장 / 해설 영어)에서 똑같이 제거해, 표시 유무와
    무관하게 '문장 내용'만으로 대조한다 → 문삽·순서·장문(어휘) 모두에서 오탐 방지.
      - 원문자(①②…)·밑줄(_) 제거
      - 문삽 삽입표시 '( ① )' 가 남긴 '빈 괄호' 제거
      - 순서 단락 라벨 '(A)~(E)' / 장문 어휘 표시 '(a)~(e)' 제거
        (여러 글자 괄호 '(e.g., …)', '(training …)' 등은 건드리지 않음)
    """
    s = re.sub(rf"[{_CIRCLED}_]", "", s)
    s = re.sub(r"\(\s*[A-Ea-e]\s*\)", " ", s)
    s = re.sub(r"\(\s*\)", " ", s)
    return normalize(s)


def check_given_in_triples(records: list, given_key: str = "GIVEN",
                           triples_key: str = "SENTENCE_TRIPLES",
                           id_key: str = "NUMBER",
                           tag: str = "삽입문장 해설 누락",
                           what: str = "주어진 문장(GIVEN)") -> list[str]:
    """문삽/순서의 '주어진 문장(GIVEN)'이 문장별 해설(SENTENCE_TRIPLES)에도 들어갔는지 검사.

    ★ 사각지대였던 곳: GIVEN 은 PASSAGE 가 아니라 별도 필드라 check_sentence_coverage 가
      못 본다. 게다가 문삽은 선지가 원문자뿐이라 choices_are_in_passage 가 True → coverage
      대상에서 제외된다. 그래서 '분석 페이지(2쪽)에서 삽입문장이 통째로 빠져도' RAW 검사
      (1쪽에 있으면 통과)·coverage 둘 다 못 잡았다(실측: 문삽 2쪽 삽입문장 누락).
      → 이 검사로, GIVEN 이 있으면 그 문장이 분석에도 반드시 들어가게 강제한다.
    """
    issues: list[str] = []
    for rec in records:
        given = str(rec.get(given_key, "")).strip()
        if not given:
            continue
        triples = rec.get(triples_key, []) or []
        blob = " ".join(_coverage_clean(str(t[0]))
                        for t in triples if isinstance(t, (list, tuple)) and t)
        miss = [u for u in split_units(given)
                if _coverage_clean(u) and re.search(r"[a-z]", _coverage_clean(u))
                and _coverage_clean(u) not in blob]
        if miss:
            where = f"{id_key}={rec.get(id_key, '?')}"
            issues.append(f"[{tag}] {where}: {what}이 "
                          f"SENTENCE_TRIPLES(분석 페이지)에 없음 — 분석에 포함해야 한다")
            for m in miss:
                snippet = m if len(m) <= 70 else m[:70] + "…"
                issues.append(f"        └ {snippet}")
    return issues


def check_sentence_coverage(records: list, passage_key: str = "PASSAGE",
                            triples_key: str = "SENTENCE_TRIPLES",
                            id_key: str = "NUMBER") -> list[str]:
    """문장별 해설(SENTENCE_TRIPLES)이 지문(PASSAGE)의 '모든' 문장을 덮는지 검사한다.
    한 문장이라도 해설에서 빠지면 = 해설 페이지에 원문이 누락된 것 → 빌드를 막는다.
    (RAW 검사가 지문/선지를 본다면, 이건 '문장별 해설'의 완전성을 본다.)"""
    issues: list[str] = []
    for rec in records:
        passage = str(rec.get(passage_key, ""))
        triples = rec.get(triples_key, []) or []
        eng_blob = " ".join(_coverage_clean(str(t[0]))
                            for t in triples if isinstance(t, (list, tuple)) and t)
        missing = []
        for sent in split_units(passage):
            ns = _coverage_clean(sent)
            if ns and ns not in eng_blob:
                missing.append(sent)
        if missing:
            where = f"{id_key}={rec.get(id_key, '?')}"
            issues.append(f"[문장 해설 누락] {where}: 지문 {len(missing)}개 문장이 "
                          f"SENTENCE_TRIPLES 에 없음 (해설에 전 문장이 들어가야 함)")
            for m in missing:
                snippet = m if len(m) <= 70 else m[:70] + "…"
                issues.append(f"        └ {snippet}")
    return issues
