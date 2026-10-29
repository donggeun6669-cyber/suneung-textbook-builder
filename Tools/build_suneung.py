"""
수능 TOP 문항 DOCX 생성기 (안정화 버전, JSON 기반).

설계 원칙
---------
   구성을 그대로 보존하고, body 내용을 문항 수만큼 복제한 뒤 토큰만 치환한다.
   JSON 키 = 템플릿 토큰. 이 스크립트가 JSON을 읽어 record로 변환하고,
   토큰이 비어 있으면 즉시 에러로 멈춘다(조용히 깨지지 않게).
   Word에서 열리지 않는 문제를 막는다.

실행
----
    python Tools/build_suneung.py
    -> Documents/TOP_suneung_<날짜>.docx 생성 + 구조 검증 로그 출력
"""

from __future__ import annotations

import argparse
import csv as _csv
import json
import re
import sys
import zipfile
from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
import integrity


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "Example" / "suneung_auto_template.docx"
INPUT_JSON = ROOT / "Tools" / "suneung_input.json"
OUTPUT = ROOT / "Documents" / "TOP_suneung_서식본.docx"
OUTPUT_ONESHOT = ROOT / "Documents" / "TOP_suneung_원샷모음.docx"
# CSV(정답·정답률·선지비율의 진실의 원천). 빌드가 JSON 을 이 CSV 와 대조해
# '엉뚱한 연도'·'지어낸 수치'를 물리적으로 차단한다(아래 crosscheck_csv).
CSV_PATH = ROOT / "Documents" / "정답률별_정리_평가원.csv"

YEAR_RE = re.compile(r"^\d{4}\.\d{2}\(\d{2}시행\)$")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPD = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
EMU_PER_TWIP = 635
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# ── 서식 상수 (공개본: 전부 예시 값) ─────────────────────────────
# ⚠️ 아래 여백 · 표 폭 · 글꼴 크기 · 색은 **예시 값**입니다.
#    실제 학원 교재 서식과 다르며, 학원 서식을 보호하기 위해 일부러 바꿔 공개합니다.
TOP_MARGIN_TWIPS = "1080"
SIDE_MARGIN_TWIPS = "1080"
ANALYSIS_TABLE_WIDTH = "6800"
ANALYSIS_BOX_WIDTH = "3000"
BOX_IND_LEFT = "240"
BOX_SPACE_TOP = "8"
BOX_IND_RIGHT = "360"
TITLE_BOX_X, TITLE_BOX_Y = "3000", "1000"
TITLE_BOX_W, TITLE_BOX_H = "4000", "600"
CHOICE_LINE = "320"
CHOICE_IND_LEFT = "120"
CHOICE_IND_RIGHT = "0"
CHOICE_FONT = "Cambria Math"
PASSAGE_IND_RIGHT = "240"
PCT_CORRECT_HEAD_FILL = "444444"
PCT_CORRECT_BODY_FILL = "CCCCCC"
PCT_OTHER_HEAD_FILL = "F2F2F2"


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


BODY = w("body")
P = w("p")
R = w("r")
T = w("t")
RPR = w("rPr")
PPR = w("pPr")
BR = w("br")
TBL = w("tbl")
TR = w("tr")
TC = w("tc")
TCPR = w("tcPr")
SECTPR = w("sectPr")
TXBX = w("txbxContent")
DRAWING = w("drawing")

CIRCLES = ["①", "②", "③", "④", "⑤"]

# JSON에 반드시 있어야 하는 키(=템플릿 토큰 + 구조화 데이터)
# RAW = 조교가 붙여넣은 원문(지문+각주+선지) 그대로의 복사본. 무결성 검사 기준.
REQUIRED_KEYS = [
    "RAW",
    "RANK", "NUMBER", "CORRECT_RATE", "YEAR", "QUESTION_TYPE", "QUESTION_TEXT",
    "PASSAGE", "CHOICES", "ONE_SHOT", "VOCABULARY", "KOREAN_TITLE",
    "GIST", "TOPIC", "TITLE_EN", "SUMMARY", "SENTENCE_TRIPLES", "STRUCTURE_ANALYSIS",
    "CORRECT_ANSWER", "PERCENTAGES", "KOREAN_CHOICES", "ANSWER_ANALYSIS",
]
# 비어 있어도 되는 키 (각주 없는 문항 존재)
OPTIONAL_EMPTY = {"FOOTNOTES"}  # 각주만 비울 수 있음. 전개방식은 모든 유형 필수(순서 포함).

# 자료형 기대치 (조용한 오삽입 방지 — 리스트 자리에 str, str 자리에 list 차단)
#   STR_KEYS       : 반드시 str
#   LIST_STR_KEYS  : 반드시 '문자열들의 리스트'
#   나머지(SENTENCE_TRIPLES/CORRECT_ANSWER)는 아래에서 개별 검사.
STR_KEYS = {
    "RAW", "RANK", "NUMBER", "CORRECT_RATE", "YEAR", "QUESTION_TYPE",
    "QUESTION_TEXT", "PASSAGE", "ONE_SHOT", "KOREAN_TITLE",
    "GIST", "TOPIC", "TITLE_EN", "SUMMARY", "STRUCTURE_ANALYSIS", "ANSWER_ANALYSIS",
}
LIST_STR_KEYS = {"CHOICES", "VOCABULARY", "PERCENTAGES", "KOREAN_CHOICES"}

_PCT_RE = re.compile(r"^\d+(?:\.\d+)?\s*%$")

# 이 빌드가 치환하는 '템플릿 토큰' 전체. verify_output 은 사용자 원문에 우연히 들어 있는
TEMPLATE_TOKENS = [
    "RANK", "YEAR", "NUMBER", "CORRECT_RATE", "QUESTION_TEXT", "QUESTION_TYPE",
    "TYPE_COMMENT", "GIST", "TOPIC", "TITLE_EN", "SUMMARY", "ONE_SHOT",
    "PASSAGE", "KOREAN_TITLE", "VOCABULARY", "STRUCTURE_ANALYSIS",
    "KOREAN_CHOICES", "ANSWER_ANALYSIS", "PERCENT_TABLE", "FOOTNOTES",
    "CHOICES", "ANNOTATED_SENTENCES", "KOREAN_TRANSLATIONS",
]
_LEFTOVER_TOKEN_RE = re.compile(r"\{\{(" + "|".join(TEMPLATE_TOKENS) + r")\}\}")


def has_rate_data(rec: dict) -> bool:
    """이 record 가 실제 정답률/선지비율 수치를 갖고 있는지.

    CSV 에 없는 사설 모의고사라도 자체 정답률 자료가 있으면 정답률 박스를 그대로 렌더한다.
    수치가 없는 시험은 '-' 같은 자리표시자를 쓰므로, 그 경우에만 박스를 생략한다.
    """
    pcts = rec.get("PERCENTAGES") or []
    return bool(pcts) and all(_PCT_RE.match(str(p).strip()) for p in pcts)


def _typename(val) -> str:
    return type(val).__name__


def _check_str_list(where: str, key: str, val) -> None:
    """val 이 '문자열들의 리스트'인지 확인. 문자열/딕트 등이 오면 명확히 멈춘다."""
    if not isinstance(val, list):
        raise ValueError(
            f"{where}: '{key}' 는 문자열 리스트여야 하는데 {_typename(val)} 가 옴 "
            f"(리스트 자리에 단일 값을 넣지 말 것)"
        )
    for j, item in enumerate(val):
        if not isinstance(item, str):
            raise ValueError(
                f"{where}: '{key}'[{j}] 는 문자열이어야 하는데 {_typename(item)} 가 옴"
            )


# ───────────────────────────────────────────────────────────────────
#  검증 (조용히 깨지지 않게)
# ───────────────────────────────────────────────────────────────────
def validate_record(idx: int, rec: dict) -> None:
    where = f"record[{idx}] (NUMBER={rec.get('NUMBER', '?')})"
    for key in REQUIRED_KEYS:
        if key not in rec:
            raise ValueError(f"{where}: 필수 키 '{key}' 누락")
        val = rec[key]
        # 자료형 검사 (조용한 ['...'] repr 출력 방지)
        if key in STR_KEYS and not isinstance(val, str):
            raise ValueError(
                f"{where}: '{key}' 는 문자열이어야 하는데 {_typename(val)} 가 옴 "
                f"(str 자리에 리스트/딕트를 넣지 말 것)"
            )
        if key in LIST_STR_KEYS:
            _check_str_list(where, key, val)
        if isinstance(val, str) and not val.strip() and key not in OPTIONAL_EMPTY:
            raise ValueError(f"{where}: '{key}' 값이 비어 있음")
        if isinstance(val, list) and len(val) == 0:
            raise ValueError(f"{where}: '{key}' 리스트가 비어 있음")
    # FOOTNOTES 가 있으면 문자열이어야 한다(선택적이지만 자료형은 강제)
    if "FOOTNOTES" in rec and not isinstance(rec["FOOTNOTES"], str):
        raise ValueError(
            f"{where}: 'FOOTNOTES' 는 문자열이어야 하는데 {_typename(rec['FOOTNOTES'])} 가 옴"
        )
    # SENTENCE_TRIPLES 는 '리스트들의 리스트'
    if not isinstance(rec["SENTENCE_TRIPLES"], list):
        raise ValueError(
            f"{where}: 'SENTENCE_TRIPLES' 는 리스트여야 하는데 "
            f"{_typename(rec['SENTENCE_TRIPLES'])} 가 옴"
        )
    if len(rec["CHOICES"]) != 5:
        raise ValueError(f"{where}: CHOICES 는 5개여야 함 (현재 {len(rec['CHOICES'])})")
    if len(rec["KOREAN_CHOICES"]) != 5:
        raise ValueError(f"{where}: KOREAN_CHOICES 는 5개여야 함 (현재 {len(rec['KOREAN_CHOICES'])})")
    if len(rec["PERCENTAGES"]) != 5:
        raise ValueError(f"{where}: PERCENTAGES 는 5개여야 함")
    try:
        correct = int(rec["CORRECT_ANSWER"])
    except (TypeError, ValueError):
        raise ValueError(
            f"{where}: CORRECT_ANSWER 는 1~5 정수여야 하는데 "
            f"{_typename(rec['CORRECT_ANSWER'])}({rec['CORRECT_ANSWER']!r}) 가 옴"
        )
    if not (1 <= correct <= 5):
        raise ValueError(f"{where}: CORRECT_ANSWER 는 1~5")
    if len(rec["VOCABULARY"]) < 20:
        raise ValueError(f"{where}: VOCABULARY 는 20개 권장 (현재 {len(rec['VOCABULARY'])})")
    for ti, t in enumerate(rec["SENTENCE_TRIPLES"]):
        if not isinstance(t, list):
            raise ValueError(
                f"{where}: SENTENCE_TRIPLES[{ti}] 는 [영어, 구문, 해석] 리스트여야 하는데 "
                f"{_typename(t)} 가 옴"
            )
        if len(t) != 3:
            raise ValueError(f"{where}: SENTENCE_TRIPLES 각 항목은 [영어, 구문, 해석] 3개여야 함")
        for k, part in enumerate(t):
            if not isinstance(part, str):
                raise ValueError(
                    f"{where}: SENTENCE_TRIPLES[{ti}][{k}] 는 문자열이어야 하는데 "
                    f"{_typename(part)} 가 옴"
                )

    # 전개방식(STRUCTURE_ANALYSIS): ★ 모든 유형 필수(순서 포함). 비면 분석 페이지의 전개방식
    #   박스가 빈 채로 나가므로 — 데이터 단계에서 강제로 막는다. 첫 줄은 반드시
    #   '전개방식: 순접형(잇지)/역접형(꺾지)/복합형' 헤더 + 트리 본문이 있어야 한다.
    qtype = str(rec.get("QUESTION_TYPE", ""))
    sa = str(rec.get("STRUCTURE_ANALYSIS", "")).strip()
    if not sa:
        raise ValueError(
            f"{where}: '{qtype}' 유형도 전개방식(STRUCTURE_ANALYSIS)이 필수인데 비어 있음 "
            f"(모든 유형 필수 — 빈 트리로 내보내면 안 됨)"
        )
    sa_lines = [ln for ln in sa.split("\n") if ln.strip()]
    head = sa_lines[0].strip() if sa_lines else ""
    if not head.startswith("전개방식:"):
        raise ValueError(
            f"{where}: 전개방식 첫 줄은 '전개방식: …' 로 시작해야 함 (현재: {head[:40]!r})"
        )
    if not any(k in head for k in ("순접형", "역접형", "복합형")):
        raise ValueError(
            f"{where}: 전개방식 첫 줄에 '순접형(잇지)/역접형(꺾지)/복합형' 중 하나가 있어야 함 "
            f"(현재: {head[:40]!r})"
        )
    if len(sa_lines) < 2:
        raise ValueError(
            f"{where}: 전개방식에 트리 본문이 없음(헤더 한 줄뿐) — 노드를 채울 것"
        )


# ───────────────────────────────────────────────────────────────────
#  run 텍스트 도우미
# ───────────────────────────────────────────────────────────────────
def block_text(el: etree._Element) -> str:
    return "".join(t.text or "" for t in el.iter(T))


def _set_run_text(run: etree._Element, text: str) -> None:
    """run 안의 기존 w:t / w:br 를 지우고(rPr·drawing 보존) 줄바꿈을 w:br 로 기록."""
    for child in list(run):
        local = etree.QName(child).localname
        if local in ("t", "br"):
            run.remove(child)
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        if i:
            run.append(etree.Element(BR))
        node = etree.Element(T)
        if line[:1] == " " or line[-1:] == " ":
            node.set(XML_SPACE, "preserve")
        node.text = line
        run.append(node)


def apply_tokens_once(text: str, mapping: list[tuple[str, str]]) -> str:
    """(공개본 — 수치가 든 설명은 뺐습니다)

    사용자 문자열(PASSAGE 등)에 우연히 들어 있던 '{{...}}' 가 뒤이은 토큰 치환에서
    다시 바뀌는 사고를 막는다. 같은 위치에서 여러 key 가 매칭되면 mapping 순서(앞=우선).
    """
    pairs = [(o, str(n)) for (o, n) in mapping if o]
    out: list[str] = []
    i, length = 0, len(text)
    while i < length:
        for old, new in pairs:
            if text.startswith(old, i):
                out.append(new)
                i += len(old)
                break
        else:
            out.append(text[i])
            i += 1
    return "".join(out)


def replace_in_runs(element: etree._Element, mapping: list[tuple[str, str]]) -> None:
    """run 단위 치환.

    run 의 '직속' w:t 만 본다. run.iter(T) 를 쓰면 drawing 안 텍스트박스의 텍스트까지
    끌어와 평면화시켜 도형을 망가뜨리므로 절대 금지. 텍스트박스 내부 run 들은
    element.iter(R) 가 별도로 방문하므로 따로 처리된다.

    """
    for run in element.iter(R):
        # drawing/그림을 담은 run 은 텍스트 치환 대상이 아니다.
        if any(etree.QName(c).localname in ("drawing", "AlternateContent", "pict")
               for c in run):
            continue
        text_nodes = run.findall(T)
        if not text_nodes:
            continue
        original = "".join(n.text or "" for n in text_nodes)
        if not any(old and old in original for old, _ in mapping):
            continue
        changed = apply_tokens_once(original, mapping)
        if changed != original:
            _set_run_text(run, changed)


def make_rpr(font: str, half_pt: str, color: str | None = None, bold: bool = False,
             italic: bool = False, east_asia: str | None = None) -> etree._Element:
    # 한글 글자는 eastAsia 폰트로 렌더된다. 미지정 시 기본 '맑은 고딕'(기존 동작 유지).
    #   구문·해석처럼 한글을 특정 폰트(예: 맑은 고딕)로 내려면 east_asia 를 명시한다.
    rpr = etree.Element(RPR)
    fonts = etree.SubElement(rpr, w("rFonts"))
    fonts.set(w("ascii"), font)
    fonts.set(w("hAnsi"), font)
    fonts.set(w("eastAsia"), east_asia if east_asia is not None else "맑은 고딕")
    if bold:
        etree.SubElement(rpr, w("b"))
    if italic:
        etree.SubElement(rpr, w("i"))
    if color:
        etree.SubElement(rpr, w("color")).set(w("val"), color)
    etree.SubElement(rpr, w("sz")).set(w("val"), half_pt)
    etree.SubElement(rpr, w("szCs")).set(w("val"), half_pt)
    return rpr


def append_text_run(paragraph: etree._Element, text: str, rpr: etree._Element,
                    line_break_after: bool = False) -> None:
    run = etree.SubElement(paragraph, R)
    run.append(deepcopy(rpr))
    _set_run_text(run, text)
    if line_break_after:
        run.append(etree.Element(BR))


def _rpr_underlined(rpr: etree._Element) -> etree._Element:
    """같은 서식에 밑줄만 더한 rPr 사본."""
    new = deepcopy(rpr)
    if new.find(w("u")) is None:
        etree.SubElement(new, w("u")).set(w("val"), "single")
    return new


def append_text_with_underline(paragraph: etree._Element, text: str,
                               rpr: etree._Element, targets: list,
                               line_break_after: bool = False) -> None:
    """(공개본 — 수치가 든 설명은 뺐습니다)

    문제의 밑줄이 분석 페이지에서 사라지면 풀이 과정을 설명할 대상이 없어진다.
    구간 뒤에 붙은 문장부호는 함께 묶는다 — run 이 부호 앞에서 갈리면 docx 전체
    텍스트에 공백이 끼어 원문 무결성 대조가 깨질 수 있다(underline_choice_spans 와 동일 규칙).
    """
    spans = []
    for item in targets or []:
        # (마커, 낱말) 쌍 — 어법/어휘는 튜플, 장문 어휘·지칭의 UNDERLINE 은 JSON 이라 리스트로 온다.
        #   마커((a)·① 등)는 밑줄 밖에 두고 그 뒤의 낱말에만 밑줄을 긋는다(원문 서식과 동일).
        if isinstance(item, (tuple, list)) and len(item) == 2:
            marker, t = item
        else:
            marker, t = None, item
        t = str(t).strip()
        if not t:
            continue
        # 어법/어휘는 인라인 원문자(①~⑤) 바로 뒤가 그 선지다 — 마커부터 찾아 오탐을 막는다.
        #   (그냥 find 하면 'is' 가 'This' 안에서 잡힌다.)
        start = 0
        if marker:
            mpos = text.find(marker)
            if mpos < 0:
                continue          # 이 문장에는 해당 선지가 없다
            start = mpos + len(marker)
        m = re.compile(r"(?<![A-Za-z])" + re.escape(t) + r"(?![A-Za-z])").search(text, start)
        i = m.start() if m else text.find(t, start)
        if i >= 0:
            e = i + len(t)
            while e < len(text) and text[e] in ".,;:!?)]”’'\"":
                e += 1
            spans.append((i, e))
    spans.sort()
    merged: list = []
    for s, e in spans:
        if merged and s <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((s, e))
    if not merged:
        append_text_run(paragraph, text, rpr, line_break_after=line_break_after)
        return
    ul_rpr = _rpr_underlined(rpr)
    pos = 0
    for s, e in merged:
        if s > pos:
            append_text_run(paragraph, text[pos:s], rpr)
        append_text_run(paragraph, text[s:e], ul_rpr)
        pos = e
    if pos < len(text):
        append_text_run(paragraph, text[pos:], rpr, line_break_after=line_break_after)
    elif line_break_after:
        list(paragraph)[-1].append(etree.Element(BR))


def clear_paragraph_runs(paragraph: etree._Element) -> None:
    for child in list(paragraph):
        if etree.QName(child).localname != "pPr":
            paragraph.remove(child)


def title_with_number(rec: dict) -> str:
    """(설명 생략 — 공개본)"""
    num = str(rec.get("NUMBER", "")).strip()
    title = str(rec.get("KOREAN_TITLE", "")).strip()
    if not num or title.startswith(f"{num}."):
        return title
    return f"{num}. {title}"


def emphasize_title(element: etree._Element, title: str) -> bool:
    """(공개본 — 수치가 든 설명은 뺐습니다)

    ★ 제목 문단은 템플릿에서 불릿 목록(numPr)이라 왼쪽에 '빨간 점'이 찍힌다. 제목 앞에는
      문제 번호가 붙으므로 점은 군더더기 → numPr 과 불릿용 들여쓰기를 함께 걷어낸다."""
    for run in element.iter(R):
        tnodes = run.findall(T)
        if not tnodes or (tnodes[0].text or "") != title:
            continue
        para = run.getparent()
        if para is not None and para.tag == P:
            para_ppr = para.find(PPR)
            if para_ppr is not None:
                for npr in para_ppr.findall(w("numPr")):
                    para_ppr.remove(npr)        # 빨간 점(불릿) 제거
                for ind in para_ppr.findall(w("ind")):
                    # 불릿이 쓰던 들여쓰기 자리를 없애 제목을 왼쪽 끝으로 당긴다.
                    for attr in ("left", "leftChars", "hanging", "hangingChars", "firstLine"):
                        if ind.get(w(attr)) is not None:
                            del ind.attrib[w(attr)]
        rpr = run.find(RPR)
        if rpr is None:
            rpr = etree.Element(RPR)
            run.insert(0, rpr)
        for old in rpr.findall(w("rFonts")):
            rpr.remove(old)
        fonts = etree.SubElement(rpr, w("rFonts"))
        fonts.set(w("ascii"), "맑은 고딕")
        fonts.set(w("hAnsi"), "맑은 고딕")
        fonts.set(w("eastAsia"), "맑은 고딕")
        for old in rpr.findall(w("b")):
            rpr.remove(old)  # 별도 Bold 처리 안 함(폰트가 이미 Bold)
        for old in rpr.findall(w("shd")):
            rpr.remove(old)
        shd = etree.SubElement(rpr, w("shd"))
        shd.set(w("val"), "clear")
        shd.set(w("color"), "auto")
        shd.set(w("fill"), "FFF2CC")  # 노란 형광
        for st in ("sz", "szCs"):
            for old in rpr.findall(w(st)):
                rpr.remove(old)
            etree.SubElement(rpr, w(st)).set(w("val"), "24")
        return True
    return False


# ───────────────────────────────────────────────────────────────────
#  ID 유일화 / Fallback 제거 (Word 호환)
# ───────────────────────────────────────────────────────────────────
def uniquify_ids(element: etree._Element, state: dict) -> None:
    for node in element.iter():
        for name, value in list(node.attrib.items()):
            qn = etree.QName(name)
            if qn.namespace == REL:
                continue  # r:embed 등 미디어 참조는 보존
            if qn.localname == "id" and str(value).isdigit():
                state["num"] += 1
                node.set(name, str(state["num"]))
            elif qn.localname in ("paraId", "textId", "anchorId", "editId"):
                state["hex"] += 1
                node.set(name, f"{state['hex'] & 0xFFFFFFFF:08X}")


def remove_mc_fallbacks(element: etree._Element) -> None:
    tag = f"{{{MC}}}Fallback"
    for parent in list(element.iter()):
        for child in list(parent):
            if child.tag == tag:
                parent.remove(child)


# ───────────────────────────────────────────────────────────────────
#  특수 영역 처리
# ───────────────────────────────────────────────────────────────────
#   one-shot 라벨 박스에 남은 '깨진 중복 라벨' 판별용(템플릿 인코딩 잔재).
_ONESHOT_JUNK_RE = re.compile(r"shoshot|\?\?\?")

def fix_oneshot_label(element: etree._Element) -> None:
    """one-shot 라벨 박스를 정리한다.

       (예: '노sshoshot]: ??? Trigger'). 학생 교재에 그대로 인쇄되던 잔재다.
    """
    for para in list(element.iter(P)):
        joined = block_text(para).strip()
        if "Trigger" in joined and _ONESHOT_JUNK_RE.search(joined):
            # 도형 구조를 건드리지 않도록 지우는 대신 올바른 라벨로 덮어쓴다
            #   (폴백 도형에 들어 있어 문단을 없애면 Word 가 싫어한다).
            first = True
            for run in para.iter(R):
                _set_run_text(run, "[One Shot] : 핵심 풀이" if first else "")
                first = False
    for para in element.iter(P):
        joined = block_text(para).strip()
        if joined == ": 핵심 풀이":
            for run in para.iter(R):
                rt = "".join(n.text or "" for n in run.iter(T))
                if rt.strip() == ":":
                    _set_run_text(run, "[One Shot] :")
                    return


def set_textbox_autofit(block: etree._Element) -> None:
    """텍스트박스(wps) bodyPr 의 noAutofit/normAutofit 을 spAutoFit 으로 교체.

    내용이 길어지면 도형이 자동으로 늘어나 글자가 잘리지 않게 한다(서식 보존).
    """
    for bp in block.iter():
        if etree.QName(bp).localname != "bodyPr":
            continue
        for child in list(bp):
            if etree.QName(child).localname in ("noAutofit", "spAutoFit", "normAutofit"):
                bp.remove(child)
        etree.SubElement(bp, f"{{{A}}}spAutoFit")
    # 선언된 박스 높이(extent/ext cy)를 최소값으로 줄인다 → spAutoFit 이 내용만큼만 키움.
    MIN_CY = 274320
    for el in block.iter():
        if etree.QName(el).localname in ("extent", "ext") and el.get("cy"):
            el.set("cy", str(MIN_CY))


def write_into_textbox(block: etree._Element, text: str, rpr: etree._Element) -> bool:
    """빈 텍스트박스(도형) 안 첫 문단에 text 를 주입(줄바꿈은 w:br). 나머지 빈 문단 제거.

    """
    tb = next((t for t in block.iter(TXBX)), None)
    if tb is None:
        return False
    paras = tb.findall(P)
    if paras:
        target = paras[0]
        for extra in paras[1:]:
            tb.remove(extra)
    else:
        target = etree.SubElement(tb, P)
    clear_paragraph_runs(target)
    ppr = target.find(PPR)
    if ppr is None:
        ppr = etree.Element(PPR)
        target.insert(0, ppr)
    for old in ppr.findall(w("jc")):
        ppr.remove(old)
    etree.SubElement(ppr, w("jc")).set(w("val"), "left")
    append_text_run(target, text, rpr)
    return True


def write_structure_into_textbox(block: etree._Element, text: str) -> bool:
    """전개방식 트리를 토큰 경계별 다색 run 으로 빈 텍스트박스에 주입한다.

    단어 중간은 나누지 않고 헤더, 기능 라벨, 신호어처럼 독립된 토큰만
    별도 run 으로 만든다. 줄바꿈은 w:br 로 기록한다.
    """
    tb = next((t for t in block.iter(TXBX)), None)
    if tb is None:
        return False
    paras = tb.findall(P)
    if paras:
        target = paras[0]
        for extra in paras[1:]:
            tb.remove(extra)
    else:
        target = etree.SubElement(tb, P)
    clear_paragraph_runs(target)

    ppr = target.find(PPR)
    if ppr is None:
        ppr = etree.Element(PPR)
        target.insert(0, ppr)
    for old in ppr.findall(w("jc")):
        ppr.remove(old)
    etree.SubElement(ppr, w("jc")).set(w("val"), "left")

    # 예시 문서 에서 추출한 팔레트.
    default_rpr = make_rpr("맑은 고딕", "20", "555555", bold=False)
    accent_rpr = make_rpr("맑은 고딕", "20", "1F4E79", bold=True)
    signal_rpr = make_rpr("맑은 고딕", "20", "2E75B6", bold=True)
    star_rpr = make_rpr("맑은 고딕", "20", "E06666", bold=True)
    header_rpr = make_rpr("맑은 고딕", "20", "1F4E79", bold=True)
    header_shd = etree.SubElement(header_rpr, w("shd"))
    header_shd.set(w("val"), "clear")
    header_shd.set(w("color"), "auto")
    header_shd.set(w("fill"), "FFF2CC")

    signal_re = re.compile(r"\[신호어:[^\]]+\]")
    function_re = re.compile(r"\((?:~?G(?:\s+결론)?|G'|S\d*)\)")
    token_re = re.compile(rf"({signal_re.pattern}|{function_re.pattern})")

    lines = str(text).split("\n")
    for line_idx, line in enumerate(lines):
        stripped = line.lstrip()
        leading = line[:len(line) - len(stripped)]
        pieces: list[tuple[str, etree._Element]] = []

        if line_idx == 0 and stripped.startswith("전개방식:"):
            pieces.append((line, header_rpr))
        elif stripped.startswith("★"):
            pieces.append((line, star_rpr))
        else:
            pos = 0
            header = re.match(r"^\[[^\]]+\]", stripped)
            if header and not stripped.startswith("[신호어:"):
                header_end = len(leading) + header.end()
                pieces.append((line[:header_end], header_rpr))
                pos = header_end

            for match in token_re.finditer(line, pos):
                if match.start() > pos:
                    pieces.append((line[pos:match.start()], default_rpr))
                token = match.group(0)
                pieces.append((
                    token,
                    signal_rpr if signal_re.fullmatch(token) else accent_rpr,
                ))
                pos = match.end()
            if pos < len(line):
                pieces.append((line[pos:], default_rpr))
            if not pieces:
                pieces.append(("", default_rpr))

        for piece_idx, (piece, rpr) in enumerate(pieces):
            append_text_run(
                target,
                piece,
                rpr,
                line_break_after=(
                    piece_idx == len(pieces) - 1
                    and line_idx != len(lines) - 1
                ),
            )
    return True


def is_bare_markers(items: list) -> bool:
    """선지가 '내용 없이 원문자(①~⑤)만'인지 — 문삽 유형 판별용.
    지문에 ①~⑤ 삽입 위치가 이미 있으므로, 이런 선지는 따로 렌더하지 않는다."""
    return bool(items) and all(re.fullmatch(r"\s*[①②③④⑤]\s*", str(c)) for c in items)


def remove_first_line_indent(element: etree._Element) -> None:
    """(공개본 — 수치가 든 설명은 뺐습니다)
    (set_left_align 과 달리 keepLines 를 넣지 않아 긴 지문의 컬럼 흐름을 막지 않는다)."""
    for para in element.iter(P):
        ppr = para.find(PPR)
        if ppr is None:
            ppr = etree.Element(PPR)
            para.insert(0, ppr)
        ind = ppr.find(w("ind"))
        if ind is None:
            ind = etree.SubElement(ppr, w("ind"))
        for a in (w("firstLine"), w("firstLineChars"), w("hanging"), w("hangingChars")):
            if ind.get(a) is not None:
                del ind.attrib[a]
        ind.set(w("firstLine"), "0")


def set_runs_size(element: etree._Element, sz: str) -> None:
    """element 안 모든 run 의 글자 크기를 sz(half-point)로 — 한 페이지 맞춤용."""
    for cr in element.iter(R):
        crpr = cr.find(RPR)
        if crpr is None:
            crpr = etree.Element(RPR)
            cr.insert(0, crpr)
        for st in ("sz", "szCs"):
            for old in crpr.findall(w(st)):
                crpr.remove(old)
            etree.SubElement(crpr, w(st)).set(w("val"), sz)


def _choice_body(ch: str) -> str:
    """선지에서 앞 원문자(①~⑤)를 떼어 낸 본문."""
    return re.sub(r"^\s*[①②③④⑤]\s*", "", str(ch)).strip()


def _strip_label(s: str) -> str:
    """앞에 붙은 '[...]' 라벨(예: '[요지]')을 떼어 낸다 — 빌드가 라벨을 또 붙이는 중복 방지."""
    return re.sub(r"^\s*\[[^\]]*\]\s*", "", str(s))


def _first_sentence(s: str) -> str:
    """첫 문장만 추린다(정답 해설을 오답 해설과 비슷한 길이로 — 너무 길지 않게)."""
    s = str(s).strip()
    i = s.find("다.")
    if i != -1:
        return s[:i + 2]
    i = s.find(". ")
    return s[:i + 1] if i != -1 else s


def choices_are_in_passage(choices: list, passage: str) -> bool:
    """선지(원문자 제외)가 '지문 안에' 들어 있는 유형인지 — 어법/어휘(밑줄 오류) 판별.
    선지가 지문에 없는 유형)는 False 라서 선지가 그대로 유지된다(유형 이름에 의존 X)."""
    p = str(passage)
    hit = sum(1 for ch in choices if _choice_body(ch) and _choice_body(ch) in p)
    return hit >= max(3, len(choices) - 1)


def underline_choice_spans(paragraph: etree._Element, choices: list,
                           phrases: list | None = None) -> None:
    """지문 본문 run 에 밑줄을 입힌다.
    phrases 가 주어지면 그 구절들(밑줄 의미 유형)에, 아니면 각 선지 본문(어법/어휘)에."""
    target, full = None, ""
    for r in paragraph.iter(R):
        txt = "".join(t.text or "" for t in r.findall(w("t")))
        if len(txt) > 40:                      # 지문 본문이 담긴 run
            target, full = r, txt
            break
    if target is None:
        return
    spans = []
    if phrases is not None:
        for ph in phrases:
            # (마커, 낱말) 쌍이면 마커 뒤부터 찾는다 — 장문 어휘·지칭의 (a)~(e) 처럼
            #   같은 낱말이 앞쪽에도 있는 경우의 오탐을 막고, 라벨은 밑줄 밖에 남긴다.
            if isinstance(ph, (tuple, list)) and len(ph) == 2:
                marker, body_txt = str(ph[0]).strip(), str(ph[1]).strip()
            else:
                marker, body_txt = None, str(ph).strip()
            if not body_txt:
                continue
            start = 0
            if marker:
                mpos = full.find(marker)
                if mpos < 0:
                    continue        # 이 문단에는 그 라벨이 없다(제시문/본문이 갈린 경우)
                start = mpos + len(marker)
            idx = full.find(body_txt, start)
            if idx >= 0:
                spans.append((idx, idx + len(body_txt)))
    else:
        for i, ch in enumerate(choices):
            body = _choice_body(ch)
            if not body:
                continue
            marker_pos = full.find("①②③④⑤"[i]) if i < 5 else -1
            bidx = full.find(body, marker_pos if marker_pos >= 0 else 0)
            if bidx >= 0:
                spans.append((bidx, bidx + len(body)))
    # 밑줄 구간 바로 뒤에 '붙은' 문장부호(. , ; : ! ?)는 구간에 포함한다 — 안 그러면 run 이
    #   분리되며 docx 전체텍스트에서 'word .'처럼 공백이 끼어 원문 무결성 대조가 깨진다.
    def _ext(e: int) -> int:
        while e < len(full) and full[e] in ".,;:!?)]”’'\"":
            e += 1
        return e
    spans = [(s, _ext(e)) for s, e in spans]
    spans.sort()
    merged = []
    for s, e in spans:
        if merged and s < merged[-1][1]:
            continue
        merged.append((s, e))
    if not merged:
        return
    base_rpr = target.find(RPR)
    segs, pos = [], 0
    for s, e in merged:
        if pos < s:
            segs.append((full[pos:s], False))
        segs.append((full[s:e], True))
        pos = e
    if pos < len(full):
        segs.append((full[pos:], False))
    parent = target.getparent()
    idx = list(parent).index(target)
    parent.remove(target)
    for k, (txt, ul) in enumerate(segs):
        nr = etree.Element(R)
        nrpr = deepcopy(base_rpr) if base_rpr is not None else etree.Element(RPR)
        for old in nrpr.findall(w("u")):
            nrpr.remove(old)
        if ul:
            etree.SubElement(nrpr, w("u")).set(w("val"), "single")
        nr.append(nrpr)
        te = etree.SubElement(nr, w("t"))
        te.set(XML_SPACE, "preserve")
        te.text = txt
        parent.insert(idx + k, nr)


def _strip_restatement(note: str, *bodies: str) -> str:
    """해설 앞에 다시 적힌 '선지 원문/해석'을 떼어 낸다.

    Confirming 박스는 선지와 해석을 CHOICES·KOREAN_CHOICES 에서 이미 렌더하므로,
    해설에 같은 말이 또 있으면 한 선지가 세 번 찍힌다(조교가 문항마다 다섯 줄씩 지워야 했다).
    앞머리에 붙은 재진술은 순서와 상관없이 반복해서 제거한다.
    """
    s = str(note).strip()
    changed = True
    while changed:
        changed = False
        for body in bodies:
            body = re.sub(r"^\s*[①②③④⑤]\s*", "", str(body or "")).strip()
            if body and s.startswith(body):
                s = s[len(body):].strip(" :—–-·,")
                changed = True
    return s


def _parse_choice_notes(aa: str, correct: int, choices: list,
                        korean_choices: list | None = None) -> tuple:
    """ANSWER_ANALYSIS 를 선지별 해설로 분해.
    - 정답: '[정답]' 블록 본문(선지 재진술은 제거).
    - 오답: '① [무관] …' 줄(앞 원문자 제거).
    선지별 줄이 없으면 per 가 비어 leftover/블록 폴백으로 처리된다."""
    circled = "①②③④⑤"
    per: dict = {}
    cur = None
    leftover: list = []
    for raw in str(aa).split("\n"):
        s = raw.strip()
        if not s:
            continue
        m = re.match(r"^([①②③④⑤])\s*(.*)$", s)
        if m:
            cur = circled.index(m.group(1)) + 1
            per[cur] = m.group(2).strip()
        elif s.startswith("[정답]"):
            reason = re.sub(r"^\[정답\]\s*[①②③④⑤]?\s*", "", s)
            cb = _choice_body(choices[correct - 1]) if 1 <= correct <= len(choices) else ""
            if cb and reason.startswith(cb):
                reason = reason[len(cb):].strip(" :—-")
            cur = correct
            per[cur] = reason
        elif cur is not None:
            per[cur] = (per.get(cur, "") + " " + s).strip()
        else:
            leftover.append(s)
    kc = korean_choices or []
    for i in list(per):
        ch = choices[i - 1] if 1 <= i <= len(choices) else ""
        kco = kc[i - 1] if 1 <= i <= len(kc) else ""
        per[i] = _strip_restatement(per[i], ch, kco)
    return per, "\n".join(leftover)


def write_choices_into_textbox(block: etree._Element, choices: list,
                               korean_choices: list, answer_analysis: str,
                               correct: int, qtype: str = "") -> bool:
    """(공개본 — 수치가 든 설명은 뺐습니다)
    선지별 해설을 ANSWER_ANALYSIS 에서 못 뽑으면(장문 합본 등) 전체 분석을 블록으로 폴백.
    """
    tb = next((t for t in block.iter(TXBX)), None)
    if tb is None:
        return False
    paras = tb.findall(P)
    if paras:
        target = paras[0]
        for extra in paras[1:]:
            tb.remove(extra)
    else:
        target = etree.SubElement(tb, P)
    clear_paragraph_runs(target)
    ppr = target.find(PPR)
    if ppr is None:
        ppr = etree.Element(PPR)
        target.insert(0, ppr)
    for old in ppr.findall(w("jc")):
        ppr.remove(old)
    etree.SubElement(ppr, w("jc")).set(w("val"), "left")

    qt = str(qtype)
    sel_en = make_rpr("Times New Roman", "22", "222222", bold=False)                       # 선지(영어)
    sel_ko = make_rpr("맑은 고딕", "22", "222222", east_asia="맑은 고딕")  # 선지(한글)
    trans_rpr = make_rpr("맑은 고딕", "20", "2E75B6", east_asia="맑은 고딕")  # 해석=구문 색(밝은 파랑)
    note_rpr = make_rpr("맑은 고딕", "20", "1F4E79", east_asia="맑은 고딕")   # 해설=짙은 파랑(정답·오답 공통)
    ans_mark = make_rpr("Times New Roman", "22", "AA0000", bold=True)                       # [정답] 마커
    ans_note = make_rpr("맑은 고딕", "20", "1F4E79", east_asia="맑은 고딕")  # 정답 해설/단일(짙은 파랑)

    if is_bare_markers(choices):
        for ln in str(answer_analysis).split("\n"):
            append_text_run(target, ln, ans_note, line_break_after=True)
        return True

    per, leftover = _parse_choice_notes(answer_analysis, correct, choices, korean_choices)
    has_per = sum(1 for v in per.values() if str(v).strip()) >= 2
    is_order = "순서" in qt
    type2 = is_order or ("요지" in qt) or ("주장" in qt)  # 해석 불필요

    if is_order:
        ans = str(choices[correct - 1]) if 1 <= correct <= len(choices) else ""
        append_text_run(target, ans, sel_ko if re.search(r"[가-힣]", ans) else sel_en,
                        line_break_after=False)
        append_text_run(target, "  [정답]", ans_mark, line_break_after=True)
        reason = str(per.get(correct, "")).strip()
        if not reason:
            reason = " ".join(l.strip() for l in str(answer_analysis).split("\n")
                              if l.strip() and not l.strip().startswith("[정답]"))
        append_text_run(target, reason, ans_note, line_break_after=True)
        return True

    for i, ch in enumerate(choices):
        idx = i + 1
        is_corr = (idx == correct)
        rpr_ch = sel_ko if re.search(r"[가-힣]", str(ch)) else sel_en
        append_text_run(target, str(ch), rpr_ch, line_break_after=not is_corr)
        if is_corr:
            append_text_run(target, "  [정답]", ans_mark, line_break_after=True)
        if not type2 and i < len(korean_choices):
            kc = re.sub(r"^\s*[①②③④⑤]\s*", "", str(korean_choices[i])).strip()
            if kc:
                append_text_run(target, kc, trans_rpr, line_break_after=True)
        # 해설 줄 — 선지별 해설이 충분할 때만 인라인 (정답 해설은 첫 문장만, 너무 길지 않게)
        if has_per:
            note = str(per.get(idx, "")).strip()
            if is_corr:
                # 정답 선지 줄에 이미 '[정답]' 마커가 붙으므로(위 append), 해설 줄에는
                #   '[정답]' 을 다시 붙이지 않는다 — 같은 정답에 라벨이 두 번 찍히던 버그.
                first = _first_sentence(note).strip()
                if first:
                    append_text_run(target, first, ans_note, line_break_after=True)
            elif note:
                # 어휘 등 오답 해설 끝에 붙은 ' — 단어'(예: '… 어려움 — differentiate')는 군더더기라 제거.
                #   (정답 선지의 교정은 위에서 유지) 끝이 '대시 + 영어 한 단어'일 때만 떼어 낸다.
                note = re.sub(r"\s*[—–\-]\s*[A-Za-z][A-Za-z\-]*\s*$", "", note)
                append_text_run(target, note, note_rpr, line_break_after=True)

    # 선지별 해설이 부족하면(장문 합본·순서 등) 전체 분석을 블록으로 덧붙인다.
    if not has_per:
        append_text_run(target, "", note_rpr, line_break_after=True)
        for ln in str(answer_analysis).split("\n"):
            append_text_run(target, ln, ans_note, line_break_after=True)
    elif leftover.strip():
        for ln in leftover.split("\n"):
            append_text_run(target, ln, note_rpr, line_break_after=True)
    return True


def is_empty_textbox(block: etree._Element) -> bool:
    """도형(텍스트박스)을 담고 있고 보이는 텍스트가 전혀 없는 빈 박스인지."""
    has_box = next((t for t in block.iter(TXBX)), None) is not None
    return has_box and block_text(block).strip() == ""


def circled(n: int) -> str:
    """(설명 생략 — 공개본)"""
    return chr(0x2460 + n - 1) if 1 <= n <= 20 else f"({n})"


def set_para_spacing(para: etree._Element, line: str | None = None,
                     before: str | None = None, after: str | None = None) -> None:
    """(공개본 — 수치가 든 설명은 뺐습니다)

    """
    ppr = para.find(PPR)
    if ppr is None:
        ppr = etree.Element(PPR)
        para.insert(0, ppr)
    sp = ppr.find(w("spacing"))
    if sp is None:
        sp = etree.SubElement(ppr, w("spacing"))
    if line is not None:
        sp.set(w("line"), line)
        sp.set(w("lineRule"), "auto")
    if before is not None:
        sp.set(w("before"), before)
    if after is not None:
        sp.set(w("after"), after)


def is_empty_filler(p: etree._Element) -> bool:
    """순수 빈 여백 문단인지(텍스트·그림·섹션·번호 없음). 분석 페이지 하단에 깔린
    그림(박스 앵커)·텍스트·sectPr·numPr(어휘 불릿 슬롯)이 있으면 보존."""
    if p.tag != P:
        return False
    if p.find(f".//{SECTPR}") is not None:
        return False
    if p.find(f".//{w('numPr')}") is not None:
        return False
    if any(etree.QName(e).localname in ("drawing", "pict", "AlternateContent")
           for e in p.iter()):
        return False
    return not block_text(p).strip()


_OWN_MARKER_RE = re.compile(
    r"^\s*"
    r"(?:\([A-E]\)\s*)?"   # 순서 유형의 (A) (B) (C) 라벨은 건너뛰고 그 뒤를 본다
    r"(?:"
    r"\[제시문\]"          # 제시문 블록 — 번호를 먹지 않는다(본문 첫 문장이 ①)
    r"|\[요약문\]"         # 요약 유형의 요약문(빈칸 (A)·(B) 그대로) — 본문 문장이 아니다
    r"|\(\s*[①-⑳]\s*\)"  # 문삽 원문 표기 ( ① ) ( ② ) …
    r"|[①-⑳]"    # 이미 원문자로 시작
    r"|[a-z]\.\s"          # a. b. c.
    r")"
)


def has_own_marker(eng: str) -> bool:
    """문장이 이미 자기 표시기호를 달고 있는가(자동번호를 붙이면 두 겹이 되는 경우).

    문삽의 ( ① ) 는 원문 표기라 반드시 살려야 하는데, 여기에 자동번호까지 붙으면
    「③ ( ① ) One possibility…」처럼 번호가 두 겹으로 박혀 조교가 손으로 지워야 했다.
    제시문도 본문 문장이 아니므로 번호를 먹지 않는다(본문 첫 문장이 ①).
    """
    return bool(_OWN_MARKER_RE.match(str(eng)))


def write_sentence_triples(element: etree._Element, triples: list,
                           strip_markers: bool = False,
                           underline_targets: list | None = None) -> None:
    """(공개본 — 수치가 든 설명은 뺐습니다)

    각 문장: [① 영어 원문] / [구문 분석] / [한글 해석] — 서로 다른 run 스타일.
    영어 줄 맨 앞에 원문자 번호(①②③…)를 붙여 목표 예시처럼 문장 순서를 드러낸다.
    """
    target = None
    for para in element.iter(P):
        if "{{ANNOTATED_SENTENCES}}" in block_text(para):
            target = para
            break
    if target is None:
        return
    clear_paragraph_runs(target)
    set_para_spacing(target, line="276", before="0")
    # 목표 예시(예시 문서) 폰트에 맞춤(#D): 원문=Times New Roman, 구문·해석=맑은 고딕.
    #   한글은 eastAsia 폰트로 렌더되므로 east_asia 를 반드시 명시한다.
    eng_rpr = make_rpr("Times New Roman", "24", "222222", bold=False)  # 원문(영어)
    mark_rpr = make_rpr("맑은 고딕", "18", "222222", bold=False)
    syn_rpr = make_rpr("맑은 고딕", "20", "2E75B6",
                       east_asia="맑은 고딕")                 # 구문 (목표 폰트/색)
    kor_rpr = make_rpr("맑은 고딕", "20", "1F4E79",
                       east_asia="맑은 고딕")
    last = len(triples) - 1
    num = 0  # 실제로 번호를 찍은 문장 수(제시문·( ① ) 문장은 건너뛰므로 i 와 다르다)
    for i, (eng, syn, kor) in enumerate(triples):
        is_last = (i == last)
        # 어법/어휘: 지문 본문 안 '선지 숫자(인라인 ①~⑤)'는 제거한다(문장 순번은 아래에서 따로 붙임).
        if strip_markers:
            eng = re.sub(r"\s*[①②③④⑤]\s*", " ", str(eng)).strip()
            syn = re.sub(r"\s*[①②③④⑤]\s*", " ", str(syn)).strip()
            kor = re.sub(r"\s*[①②③④⑤]\s*", " ", str(kor)).strip()
        # 영어 줄 맨 앞에 원문자 번호 — 작은 글씨(mark_rpr)로 별도 run, 이어서 영어 원문.
        #   단, 문장이 이미 제 표시기호를 달고 있으면(문삽의 ( ① ), 제시문) 붙이지 않는다.
        if not has_own_marker(eng):
            num += 1
            append_text_run(target, f"{circled(num)} ", mark_rpr, line_break_after=False)
        if underline_targets:
            append_text_with_underline(target, eng, eng_rpr, underline_targets,
                                       line_break_after=True)
        else:
            append_text_run(target, eng, eng_rpr, line_break_after=True)
        if str(syn).strip():
            append_text_run(target, syn, syn_rpr, line_break_after=True)
        append_text_run(target, kor, kor_rpr, line_break_after=not is_last)

    #   템플릿이 고정 레이아웃용으로 셀에 빈 문단 수십 개를 깔아둬 표가 페이지를 넘쳤다.
    for tc in element.iter(TC):
        cell_paras = tc.findall(P)
        empties = [p for p in cell_paras if is_empty_filler(p)]
        removable = empties if len(empties) < len(cell_paras) else empties[1:]
        for p in removable:
            tc.remove(p)


def set_cell_fill(cell: etree._Element, hex6: str) -> None:
    tcpr = cell.find(TCPR)
    if tcpr is None:
        tcpr = etree.Element(TCPR)
        cell.insert(0, tcpr)
    for old in tcpr.findall(w("shd")):
        tcpr.remove(old)
    shd = etree.SubElement(tcpr, w("shd"))
    shd.set(w("val"), "clear")
    shd.set(w("color"), "auto")
    shd.set(w("fill"), hex6)


def set_cell_text(cell: etree._Element, text: str, rpr: etree._Element) -> None:
    for para in cell.findall(P):
        cell.remove(para)
    para = etree.SubElement(cell, P)
    ppr = etree.SubElement(para, PPR)
    etree.SubElement(ppr, w("jc")).set(w("val"), "center")
    append_text_run(para, text, rpr)


def fill_percent_table(table: etree._Element, correct: int, rates: list) -> None:
    rows = table.findall(TR)
    if len(rows) < 2:
        return
    head = rows[0].findall(TC)
    body = rows[1].findall(TC)
    label_rpr = make_rpr("맑은 고딕", "22", "FFFFFF", bold=True)
    label_rpr_n = make_rpr("맑은 고딕", "22", "222222", bold=True)
    pct_rpr = make_rpr("맑은 고딕", "20", "222222", bold=False)
    pct_rpr_w = make_rpr("맑은 고딕", "20", "FFFFFF", bold=True)
    for i in range(min(5, len(head), len(body))):
        is_correct = (i + 1 == correct)
        if is_correct:
            set_cell_fill(head[i], PCT_CORRECT_HEAD_FILL)
            set_cell_fill(body[i], PCT_CORRECT_BODY_FILL)
            set_cell_text(head[i], CIRCLES[i], label_rpr)
            set_cell_text(body[i], rates[i], pct_rpr)
        else:
            set_cell_fill(head[i], PCT_OTHER_HEAD_FILL)
            set_cell_text(head[i], CIRCLES[i], label_rpr_n)
            set_cell_text(body[i], rates[i], pct_rpr)


# ───────────────────────────────────────────────────────────────────
#  어휘 불릿 슬롯 / 선지 정렬 도우미
# ───────────────────────────────────────────────────────────────────
def is_empty_numbered_para(p: etree._Element) -> bool:
    """numId(불릿)가 붙은 빈 문단인지 — 어휘 슬롯(빨간 점) 판별용."""
    if p.tag != P:
        return False
    if p.find(f".//{w('numPr')}") is None:
        return False
    return block_text(p).strip() == ""


def set_para_text(para: etree._Element, text: str, rpr: etree._Element | None) -> None:
    """문단의 기존 run 을 지우고(pPr·불릿 보존) 텍스트 run 하나를 넣는다."""
    for child in list(para):
        if etree.QName(child).localname != "pPr":
            para.remove(child)
    run = etree.SubElement(para, R)
    if rpr is not None:
        run.append(deepcopy(rpr))
    _set_run_text(run, text)


def set_left_align(element: etree._Element) -> None:
    """(공개본 — 수치가 든 설명은 뺐습니다)
    ※ keepLines 는 넣지 않는다 — Word 편집 화면에서 해당 문단 왼쪽 여백에 '검은 사각형(■)'
       표시를 띄워 사용자에게 거슬린다(인쇄/PDF엔 안 나오지만). 상속된 keepLines 도 제거한다."""
    for para in element.iter(P):
        ppr = para.find(PPR)
        if ppr is None:
            ppr = etree.Element(PPR)
            para.insert(0, ppr)
        for tag in ("jc", "keepLines"):
            for old in ppr.findall(w(tag)):
                ppr.remove(old)
        etree.SubElement(ppr, w("jc")).set(w("val"), "left")
        ind = ppr.find(w("ind"))
        if ind is None:
            ind = etree.SubElement(ppr, w("ind"))
        for a in (w("firstLine"), w("firstLineChars"), w("hanging"), w("hangingChars")):
            if ind.get(a) is not None:
                del ind.attrib[a]
        ind.set(w("firstLine"), "0")


def _ppr(para: etree._Element) -> etree._Element:
    ppr = para.find(PPR)
    if ppr is None:
        ppr = etree.Element(PPR)
        para.insert(0, ppr)
    return ppr


def apply_choice_format(element: etree._Element) -> None:
    """선지 문단을 최종서식본 '선택지' 스타일에 맞춘다(조교 요청).
      · 원문자/영문 글꼴을 Cambria Math 로 — 최종서식본의 ①②③ 모양이 이 글꼴이다.
    """
    for para in element.iter(P):
        ppr = _ppr(para)
        for old in ppr.findall(w("spacing")):
            ppr.remove(old)
        sp = etree.SubElement(ppr, w("spacing"))
        sp.set(w("line"), CHOICE_LINE)
        sp.set(w("lineRule"), "auto")
        ind = ppr.find(w("ind"))
        if ind is None:
            ind = etree.SubElement(ppr, w("ind"))
        ind.set(w("leftChars"), "40")
        ind.set(w("left"), CHOICE_IND_LEFT)
        ind.set(w("right"), CHOICE_IND_RIGHT)
    for run in element.iter(R):
        rpr = run.find(RPR)
        if rpr is None:
            rpr = etree.Element(RPR)
            run.insert(0, rpr)
        for old in rpr.findall(w("rFonts")):
            rpr.remove(old)
        fonts = etree.Element(w("rFonts"))
        for attr in ("ascii", "hAnsi", "cs"):
            fonts.set(w(attr), CHOICE_FONT)
        rpr.insert(0, fonts)


def apply_passage_right_indent(element: etree._Element) -> None:
    """(설명 생략 — 공개본)"""
    for para in element.iter(P):
        ppr = _ppr(para)
        ind = ppr.find(w("ind"))
        if ind is None:
            ind = etree.SubElement(ppr, w("ind"))
        ind.set(w("right"), PASSAGE_IND_RIGHT)


def box_paragraph(element: etree._Element) -> None:
    """문단에 네모 테두리(pBdr)를 두른다 — 순서/문삽의 제시문(주어진 글·문장)용.
    대조(문단 텍스트)도 그대로 통과해 더 안전하다."""
    paras = [p for p in element.iter(P)]
    for i, para in enumerate(paras):
        ppr = _ppr(para)
        for old in ppr.findall(w("pBdr")):
            ppr.remove(old)
        bdr = etree.Element(w("pBdr"))
        for side in ("top", "left", "bottom", "right"):
            # 여러 문단이면 가운데 줄에 가로선이 생기지 않게 위/아래는 처음·끝에만.
            if side == "top" and i != 0:
                continue
            if side == "bottom" and i != len(paras) - 1:
                continue
            el = etree.SubElement(bdr, w(side))
            el.set(w("val"), "single")
            el.set(w("sz"), "4")
            # 위아래 균형(조교 지적): 줄간격이 넓어 글자는 줄 상자 위쪽에 붙는다. 그래서
            el.set(w("space"), BOX_SPACE_TOP if side == "top"
                   else "4" if side == "bottom" else "6")
            el.set(w("color"), "auto")
        ppr.append(bdr)
        #   네모의 왼쪽 변이 아래 지문의 (A)\첫 문장과 같은 세로선에서 시작한다(조교 지적).
        ind = ppr.find(w("ind"))
        if ind is None:
            ind = etree.SubElement(ppr, w("ind"))
        ind.set(w("left"), BOX_IND_LEFT)
        ind.set(w("right"), BOX_IND_RIGHT)
        for attr in ("leftChars", "hanging", "hangingChars", "rightChars"):
            if ind.get(w(attr)) is not None:
                del ind.attrib[w(attr)]


def set_table_width(element: etree._Element, twips: str) -> None:
    """부동 표(요지+문장해석)의 폭을 twips 로 넓힌다(tblW + tblGrid 동시)."""
    for tbl in element.iter(w("tbl")):
        pr = tbl.find(w("tblPr"))
        if pr is None:
            continue
        tw = pr.find(w("tblW"))
        if tw is None:
            tw = etree.SubElement(pr, w("tblW"))
        tw.set(w("w"), twips)
        tw.set(w("type"), "dxa")
        grid = tbl.find(w("tblGrid"))
        if grid is not None:
            cols = grid.findall(w("gridCol"))
            if len(cols) == 1:
                cols[0].set(w("w"), twips)
        for tc in tbl.iter(TC):
            tcpr = tc.find(w("tcPr"))
            if tcpr is None:
                continue
            cw = tcpr.find(w("tcW"))
            if cw is not None and cw.get(w("type")) == "dxa":
                cw.set(w("w"), twips)


# ───────────────────────────────────────────────────────────────────
#  치환 매핑 (composite 키 먼저: 깨진 한글 라벨 복구 + 토큰 채움)
# ───────────────────────────────────────────────────────────────────
_EXAM_LABEL = {"03": "3월", "04": "4월", "06": "6모", "07": "7월",
               "09": "9모", "10": "10월", "11": "수능"}


def exam_label(year: str) -> str:
    """(공개본 — 수치가 든 설명은 뺐습니다)
    'YYYY.MM(...)' 형식이 아니면(사설 모의고사 등) 빈 문자열."""
    m = re.match(r"^\s*\d{4}\.(\d{2})", str(year))
    if not m:
        return ""
    return _EXAM_LABEL.get(m.group(1), f"{int(m.group(1))}월")


def build_mapping(rec: dict, pyeongga: bool = True) -> list[tuple[str, str]]:
    # 출처 라벨 — 시험 이름을 YEAR 의 '월'에서 뽑아 정확히 쓴다(조교 요청).
    #   정답률은 있을 때만 붙인다(사설 등 정답률이 없으면 번호까지만).
    source_label = f"[{rec['YEAR']} {exam_label(rec['YEAR'])} {rec['NUMBER']}번".rstrip()
    source_label = re.sub(r"\s+", " ", source_label)
    if has_rate_data(rec):
        source_label += f"-{rec['CORRECT_RATE']}]"
    else:
        source_label += "]"
    # 오답 상위 순위 라벨: 정답률(오답률순)이 있어야 의미가 있다.
    if pyeongga or has_rate_data(rec):
        rank_label = f"오답 상위 {rec['RANK']}"
    else:
        rank_label = f"{rec['NUMBER']}번 문항"
    return [
        # ── 깨진 한글 라벨 + 토큰을 한 번에 (긴 키 우선) ──
        ("?? ?? {{RANK}}", rank_label),
        ("[{{YEAR}}-?? {{NUMBER}}?-{{CORRECT_RATE}}]", source_label),
        # 문제 줄: 발문만 남긴다. 유형 태그([순서]·[빈칸]…)와 TYPE_COMMENT 는 넣지 않는다 —
        #   발문 자체가 이미 유형을 말하고 있어 중복 표기였다(조교 요청).
        #   (One Shot 이 핵심 풀이팁을 담아 중복이므로 본문에선 빼고, ONE_SHOT 만 원샷모음으로.)
        ("{{QUESTION_TEXT}} [{{QUESTION_TYPE}} comment] {{TYPE_COMMENT}}",
         f"{rec['QUESTION_TEXT']}"),
        ("[??] {{GIST}}", f"[요지] {_strip_label(rec['GIST'])}"),
        ("[??] {{TOPIC}}", f"[주제] {_strip_label(rec['TOPIC'])}"),
        ("[??] {{TITLE_EN}}", f"[제목] {_strip_label(rec['TITLE_EN'])}"),
        ("[??] {{SUMMARY}}", f"[요약] {_strip_label(rec['SUMMARY'])}"),
        # one-shot 내용 박스: 서식본에서는 "비워둔다"(라벨 박스만 남고 내용 X).
        #   → one shot 내용은 별도 '원샷모음' docx(create_oneshot_doc)로 모은다.
        ("[??'s one shot]: ??? Trigger {{ONE_SHOT}}", ""),
        # 템플릿의 one-shot 라벨이 run 으로 쪼개져 위 복합키가 안 맞을 때를 대비한 폴백:
        #   서식본에서는 {{ONE_SHOT}} 토큰 자체를 항상 비운다(라벨 텍스트는 별도 run 이라 유지됨).
        #   create_oneshot_doc 는 build_mapping 을 안 쓰므로 원샷모음에는 영향 없음.
        ("{{ONE_SHOT}}", ""),
        # ── 단일 토큰 ──
        ("{{PASSAGE}}", rec["PASSAGE"]),
        # 우측 단의 원래 제목 자리는 비운다(제목은 위 요지 박스로 이동 #C).
        #   형광+굵게+크게를 입히고, 템플릿의 불릿(빨간 점)을 걷어낸다.
        #   ★ 제목 앞에 문제 번호를 붙인다(조교 요청) — 해설 페이지만 봐도 몇 번인지 알게.
        ("{{KOREAN_TITLE}}", title_with_number(rec)),
        #   은 본문 토큰을 비우고 바로 뒤 빈 텍스트박스(도형) 안에 주입한다.
        #   → build() 의 도형 주입 분기에서 처리.
    ]


def set_top_margin(root: etree._Element, twips: str) -> None:
    """여백 조정: ★ '분석(해설) 페이지만' 여백을 좁혀 넓게 쓰고,
    구분 기준 = 섹션의 단(column) 수. ★ 실측(렌더 프로브)으로 확인한 매핑:
    ⛔ 좌/우를 '전 페이지 공통'으로 좁히지 말 것 — 문제 페이지 지문 폭이 원본과 달라져
       """
    for sect in root.iter(w("sectPr")):
        pm = sect.find(w("pgMar"))
        if pm is None:
            continue
        cols = sect.find(w("cols"))
        num = cols.get(w("num")) if cols is not None else None
        is_two_col = num is not None and num.isdigit() and int(num) >= 2
        if is_two_col:      # 문제 원문 페이지 → 템플릿 원본 여백 유지
            continue
        pm.set(w("left"), SIDE_MARGIN_TWIPS)
        pm.set(w("right"), SIDE_MARGIN_TWIPS)
        pm.set(w("top"), twips)


def widen_analysis_boxes(root: etree._Element, twips: str) -> None:
    """해설 페이지 오른쪽 inline 텍스트박스(전개방식 트리 · Confirming Options)를 넓힌다.

    본문이 넓어져도 따라 늘지 않는다. 그 결과 해설 페이지 오른쪽에만 여백이 더 남아
    콘텐츠가 왼쪽으로 쏠려 보인다 → 폭을 직접 키워 오른쪽 끝을 본문 경계에 맞춘다.
    wp:extent(문서 흐름이 잡는 자리)와 a:ext(도형 자체 크기)를 함께 바꿔야 한다."""
    target = str(int(twips) * EMU_PER_TWIP)
    for inline in root.iter(f"{{{WPD}}}inline"):
        ext = inline.find(f"{{{WPD}}}extent")
        if ext is None:
            continue
        try:
            cx = int(ext.get("cx"))
        except (TypeError, ValueError):
            continue
        #   다른 도형까지 늘리지 않도록 원본 폭 근처만 골라 바꾼다.
        if not (1_700_000 <= cx <= 2_150_000):
            continue
        old = str(cx)
        ext.set("cx", target)
        for a_ext in inline.iter(f"{{{A}}}ext"):
            if a_ext.get("cx") == old:
                a_ext.set("cx", target)


def make_title_placeholder(state: dict) -> etree._Element:
    """문제 페이지 맨 위에 놓을 '빈 투명 텍스트 상자' 문단을 만든다(조교 요청).

    배경(noFill)·테두리(ln noFill)가 모두 없어서 글자를 넣지 않으면 인쇄물에 아무것도
    나오지 않는다. 필요 없으면 워드에서 상자를 골라 지우면 그만이다.
    페이지 기준 절대 좌표(wrapNone)라 본문 흐름을 밀지 않는다."""
    def emu(tw: str) -> str:
        return str(int(tw) * EMU_PER_TWIP)

    sid = state.get("title_box_id", 90000)
    state["title_box_id"] = sid + 1

    para = etree.Element(P)
    ppr = etree.SubElement(para, PPR)
    sp = etree.SubElement(ppr, w("spacing"))
    for k, v in (("before", "0"), ("after", "0"), ("line", "24"), ("lineRule", "exact")):
        sp.set(w(k), v)

    drawing = etree.SubElement(etree.SubElement(para, R), w("drawing"))
    anchor = etree.SubElement(drawing, f"{{{WPD}}}anchor")
    for k, v in (("distT", "0"), ("distB", "0"), ("distL", "0"), ("distR", "0"),
                 ("simplePos", "0"), ("relativeHeight", "251900000"), ("behindDoc", "0"),
                 ("locked", "0"), ("layoutInCell", "1"), ("allowOverlap", "1")):
        anchor.set(k, v)
    simple = etree.SubElement(anchor, f"{{{WPD}}}simplePos")
    simple.set("x", "0")
    simple.set("y", "0")
    for tag, off in ((f"{{{WPD}}}positionH", TITLE_BOX_X), (f"{{{WPD}}}positionV", TITLE_BOX_Y)):
        pos = etree.SubElement(anchor, tag)
        pos.set("relativeFrom", "page")
        etree.SubElement(pos, f"{{{WPD}}}posOffset").text = emu(off)
    ext = etree.SubElement(anchor, f"{{{WPD}}}extent")
    ext.set("cx", emu(TITLE_BOX_W))
    ext.set("cy", emu(TITLE_BOX_H))
    eff = etree.SubElement(anchor, f"{{{WPD}}}effectExtent")
    for side in ("l", "t", "r", "b"):
        eff.set(side, "0")
    etree.SubElement(anchor, f"{{{WPD}}}wrapNone")
    docpr = etree.SubElement(anchor, f"{{{WPD}}}docPr")
    docpr.set("id", str(sid))
    docpr.set("name", f"제목 입력란 {sid}")
    etree.SubElement(anchor, f"{{{WPD}}}cNvGraphicFramePr")

    gdata = etree.SubElement(etree.SubElement(anchor, f"{{{A}}}graphic"), f"{{{A}}}graphicData")
    gdata.set("uri", WPS)
    wsp = etree.SubElement(gdata, f"{{{WPS}}}wsp")
    etree.SubElement(wsp, f"{{{WPS}}}cNvSpPr").set("txBox", "1")
    sppr = etree.SubElement(wsp, f"{{{WPS}}}spPr")
    xfrm = etree.SubElement(sppr, f"{{{A}}}xfrm")
    off = etree.SubElement(xfrm, f"{{{A}}}off")
    off.set("x", "0")
    off.set("y", "0")
    aext = etree.SubElement(xfrm, f"{{{A}}}ext")
    aext.set("cx", emu(TITLE_BOX_W))
    aext.set("cy", emu(TITLE_BOX_H))
    geom = etree.SubElement(sppr, f"{{{A}}}prstGeom")
    geom.set("prst", "rect")
    etree.SubElement(geom, f"{{{A}}}avLst")
    etree.SubElement(sppr, f"{{{A}}}noFill")                       # 투명 배경
    etree.SubElement(etree.SubElement(sppr, f"{{{A}}}ln"), f"{{{A}}}noFill")   # 테두리 없음

    content = etree.SubElement(etree.SubElement(wsp, f"{{{WPS}}}txbx"), w("txbxContent"))
    inner = etree.SubElement(content, P)
    ippr = etree.SubElement(inner, PPR)
    etree.SubElement(ippr, w("jc")).set(w("val"), "left")
    # 미리 잡아 두는 글자 서식 — 상자를 클릭해 바로 타이핑하면 이 서식으로 들어간다.
    irpr = etree.SubElement(ippr, w("rPr"))
    fonts = etree.SubElement(irpr, w("rFonts"))
    for attr in ("ascii", "hAnsi", "eastAsia"):
        fonts.set(w(attr), "맑은 고딕")
    for tag in ("sz", "szCs"):
        etree.SubElement(irpr, w(tag)).set(w("val"), "26")

    bodypr = etree.SubElement(wsp, f"{{{WPS}}}bodyPr")
    for k, v in (("rot", "0"), ("spcFirstLastPara", "0"), ("vertOverflow", "overflow"),
                 ("horzOverflow", "overflow"), ("vert", "horz"), ("wrap", "square"),
                 ("lIns", "0"), ("tIns", "0"), ("rIns", "0"), ("bIns", "0"),
                 ("anchor", "t"), ("anchorCtr", "0")):
        bodypr.set(k, v)
    etree.SubElement(bodypr, f"{{{A}}}noAutofit")
    return para


def remove_caricature(root: etree._Element, doc_xml: Path) -> int:
    """문제 페이지 지문 위에 겹쳐 있는 작은 캐리커쳐 그림을 지운다(조교 요청).

    디자인 요소라 그대로 둔다.
    템플릿이 바뀌어 파일을 못 찾으면 조용히 지나간다 — 장식 하나로 빌드를 멈출 이유는 없다."""
    rels_path = doc_xml.parent / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        return 0
    try:
        rels = etree.parse(str(rels_path)).getroot()
    except etree.XMLSyntaxError:
        return 0
    rids = {r.get("Id") for r in rels if "image6.png" in (r.get("Target") or "")}
    if not rids:
        return 0
    removed = 0
    for pic in list(root.iter(f"{{{PIC}}}pic")):
        blip = pic.find(f".//{{{A}}}blip")
        if blip is None or blip.get(f"{{{REL}}}embed") not in rids:
            continue
        parent = pic.getparent()
        if parent is not None:
            parent.remove(pic)
            removed += 1
    return removed


def remove_dashed_hairlines(root: etree._Element) -> None:
    """오른쪽 컬럼 상단(전개방식 헤더 위)에 깔리던 '점선 가로줄'(의도치 않은 선)을 제거.
    (초록 박스 테두리·작은 점 도형은 높이가 커서 걸리지 않는다 — 오검출 방지.)"""
    ns = {"a": A}
    HAIRLINE_MAX_CY = 5000
    for wsp in list(root.iter(f"{{{WPS}}}wsp")):
        # 도형 자체의 크기(a:ext). spPr 은 wps 네임스페이스라 구조 경로 대신 후손검색으로 찾는다
        #   (wps:spPr/a:xfrm/a:ext — a:spPr 로 찾으면 못 잡던 버그가 있었다).
        ext = wsp.find(".//a:ext", ns)
        if ext is None:
            continue
        cy = ext.get("cy")
        if not (cy and cy.isdigit() and int(cy) <= HAIRLINE_MAX_CY):
            continue
        if wsp.find(".//a:prstDash", ns) is None:
            continue
        parent = wsp.getparent()
        if parent is not None:
            parent.remove(wsp)


# ───────────────────────────────────────────────────────────────────
#  메인 빌드
# ───────────────────────────────────────────────────────────────────
def build(template: Path, output: Path, records: list, pyeongga: bool = True) -> None:
    for i, rec in enumerate(records):
        validate_record(i, rec)

    with TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        with zipfile.ZipFile(template, "r") as zf:
            zf.extractall(tmp_dir)

        doc_xml = tmp_dir / "word" / "document.xml"
        parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)
        tree = etree.parse(str(doc_xml), parser)
        root = tree.getroot()
        body = root.find(BODY)
        if body is None:
            raise RuntimeError("w:body 를 찾을 수 없음")

        children = list(body)
        final_sectpr = None
        if children and children[-1].tag == SECTPR:
            final_sectpr = deepcopy(children[-1])
            children = children[:-1]

        # nextPage 로 만들면, 복제된 다음 문제가 항상 새 페이지에서 시작한다.
        for ch in children:
            sp = ch.find(f".//{SECTPR}") if ch.tag == P else None
            if sp is not None:
                for ty in sp.findall(w("type")):
                    sp.remove(ty)  # type 제거 = nextPage 기본
                break

        # passage 문단을 선지(choices) 문단의 스타일 원본으로 보관
        passage_tmpl = next(
            (deepcopy(c) for c in children if block_text(c).strip() == "{{PASSAGE}}"), None
        )

        for c in list(body):
            body.remove(c)

        state = {"num": 900000, "hex": 0x70000000}

        kchoice_rpr = make_rpr("맑은 고딕", "24", "222222", bold=False)

        for ri, rec in enumerate(records):
            mapping = build_mapping(rec, pyeongga)
            pending_percent = False
            pending_box = None
            #   반드시 박스에 주입돼야 한다. structure_written 으로 실제 주입을 추적한다.
            needs_tree = bool(str(rec["STRUCTURE_ANALYSIS"]).strip())
            structure_written = False
            vocab_queue = None
            vocab_rpr = None
            _bullet_seen = 0    # 지나온 빈 불릿 슬롯 수(마지막 슬롯 복제 판별용)
            kchoices = "\n".join(rec["KOREAN_CHOICES"])

            # 나머지를 '마지막 슬롯'에 합쳐(여러 줄) 한 단어도 소실되지 않게 한다.
            # C: 반의어(또는 유의어)가 '없음'이면 그 표기 자체를 지운다("(반) 없음" → 삭제).
            def _clean_vocab(s: str) -> str:
                s = re.sub(r"\s*\((?:반|유)\)\s*없음", "", str(s))
                s = re.sub(r"\s*\((?:반|유)\)\s*$", "", s)
                # D: "(유)"·"(반)"를 바로 뒤 단어와 비분리 공백( )으로 묶어, 줄바꿈 시
                #   "(반)"만 윗줄에 남지 않고 "(반) 단어"가 통째로 다음 줄로 내려가게 한다.
                s = re.sub(r"\((유|반)\)\s+", "(\\1) ", s)
                return s.rstrip()
            vocab_words = [_clean_vocab(w) for w in rec["VOCABULARY"]]
            #   빈 불릿 문단이 모자라면 렌더 단계에서 마지막 불릿 문단을 복제해 채운다
            #   (마지막 몇 단어에 빨간 점이 빠지던 서식 불통일 해결).
            vocab_slots = [[wd] for wd in vocab_words]
            n_empty_bullets = sum(1 for c in children if is_empty_numbered_para(c))

            # 문제 페이지 맨 위 '제목 입력란'(빈 투명 상자). 이 record 의 첫 요소로 넣어야
            #   문제 페이지에 얹힌다(페이지 기준 절대 좌표라 본문은 밀리지 않는다).
            body.append(make_title_placeholder(state))

            for child in children:
                raw = block_text(child)

                # [One Shot] 블록은 서식본에서 통째로 뺀다(조교 요청).
                #   append 하지 않으면 박스도 글자도 이미지도 남지 않는다.
                #   one shot '내용'은 원샷모음(create_oneshot_doc)으로만 나간다.
                if "{{ONE_SHOT}}" in raw or ("Trigger" in raw and _ONESHOT_JUNK_RE.search(raw)):
                    continue

                if "{{VOCABULARY}}" in raw:
                    clone = deepcopy(child)
                    uniquify_ids(clone, state)
                    remove_mc_fallbacks(clone)
                    for run in clone.iter(R):
                        if any("{{VOCABULARY}}" in (t.text or "") for t in run.findall(T)):
                            rp = run.find(RPR)
                            vocab_rpr = deepcopy(rp) if rp is not None else None
                            break
                    first = "\n".join(vocab_slots[0]) if vocab_slots else ""
                    replace_in_runs(clone, [("{{VOCABULARY}}", first)])
                    body.append(clone)
                    vocab_queue = list(vocab_slots[1:])
                    continue
                if vocab_queue is not None and is_empty_numbered_para(child):
                    _bullet_seen += 1
                    if vocab_queue:
                        clone = deepcopy(child)
                        uniquify_ids(clone, state)
                        remove_mc_fallbacks(clone)
                        set_para_text(clone, "\n".join(vocab_queue.pop(0)), vocab_rpr)
                        body.append(clone)
                        # 마지막 빈 불릿 슬롯인데 단어가 남았으면, 이 불릿 문단을 복제해
                        #   남은 단어도 각자 불릿을 갖게 한다(끝 단어 빨간 점 누락 방지).
                        if _bullet_seen == n_empty_bullets:
                            while vocab_queue:
                                extra = deepcopy(child)
                                uniquify_ids(extra, state)
                                remove_mc_fallbacks(extra)
                                set_para_text(extra, "\n".join(vocab_queue.pop(0)), vocab_rpr)
                                body.append(extra)
                    # 남은 내용 없으면 빈 불릿(빨간 점만)은 버린다
                    continue

                if "{{STRUCTURE_ANALYSIS}}" in raw:
                    # C(구조): 본문 설명(왼쪽 단)이 끝나고 전개방식부터는 '항상' 오른쪽 단에서
                    #   시작하도록 컬럼 브레이크를 넣는다. 지문 길이와 무관하게 전개방식/정답률/
                    #   Confirming 은 오른쪽, 본문 설명은 왼쪽으로 고정된다.
                    _cb = etree.SubElement(etree.SubElement(body, P), R)
                    etree.SubElement(_cb, BR).set(w("type"), "column")
                    clone = deepcopy(child)
                    uniquify_ids(clone, state)
                    remove_mc_fallbacks(clone)
                    # F: 첫 줄("전개방식: …(잇지/꺾지)")은 박스 위 헤더(토큰 자리)로, 나머지 트리만 박스에.
                    sa_lines = rec["STRUCTURE_ANALYSIS"].split("\n")
                    if sa_lines and sa_lines[0].lstrip().startswith("전개방식"):
                        head_line = sa_lines[0].strip()
                        tree_body = "\n".join(sa_lines[1:]).lstrip("\n")
                    else:
                        head_line, tree_body = "", rec["STRUCTURE_ANALYSIS"]
                    replace_in_runs(clone, [("{{STRUCTURE_ANALYSIS}}", head_line)])
                    body.append(clone)
                    pending_box = ("structure", tree_body, None)
                    continue
                if "{{KOREAN_CHOICES}}" in raw:
                    clone = deepcopy(child)
                    uniquify_ids(clone, state)
                    remove_mc_fallbacks(clone)
                    replace_in_runs(clone, [("{{KOREAN_CHOICES}}", "(Confirming Options)"),
                                            ("{{ANSWER_ANALYSIS}}", "")])
                    set_para_spacing(clone, before="200")  # G: 정답률 ↔ Confirming 사이 한 줄 간격
                    body.append(clone)
                    pending_box = (
                        "choices",
                        (rec["CHOICES"], rec["KOREAN_CHOICES"],
                         rec["ANSWER_ANALYSIS"], int(rec["CORRECT_ANSWER"]),
                         str(rec["QUESTION_TYPE"])),
                        None,
                    )
                    continue
                # 예약된 내용이 있으면 바로 뒤 빈 텍스트박스 안에 주입 + 자동확장
                if pending_box is not None and is_empty_textbox(child):
                    clone = deepcopy(child)
                    uniquify_ids(clone, state)
                    remove_mc_fallbacks(clone)
                    if pending_box[0] == "structure":
                        if write_structure_into_textbox(clone, pending_box[1]):
                            structure_written = True
                    elif pending_box[0] == "choices":
                        write_choices_into_textbox(clone, *pending_box[1])
                    else:
                        write_into_textbox(clone, pending_box[1], pending_box[2])
                    set_textbox_autofit(clone)
                    body.append(clone)
                    pending_box = None
                    continue

                # 정답률 데이터가 없는 시험(사설 등)은 라벨 문단과 표를 통째로 생략한다
                # (빈 표가 남지 않도록). 자체 정답률이 있으면 CSV 없이도 그대로 렌더한다.
                if pending_percent and child.tag == TBL:
                    if pending_percent == "skip":
                        pending_percent = False
                        continue
                    clone = deepcopy(child)
                    uniquify_ids(clone, state)
                    remove_mc_fallbacks(clone)
                    fill_percent_table(clone, int(rec["CORRECT_ANSWER"]), rec["PERCENTAGES"])
                    body.append(clone)
                    pending_percent = False
                    continue
                if "{{PERCENT_TABLE}}" in raw:
                    if not has_rate_data(rec):
                        pending_percent = "skip"
                        continue
                    clone = deepcopy(child)
                    uniquify_ids(clone, state)
                    remove_mc_fallbacks(clone)
                    replace_in_runs(clone, [("{{PERCENT_TABLE}}", "[정답률]")])
                    set_para_spacing(clone, before="200")  # G: 전개방식 ↔ 정답률 사이 한 줄 간격
                    body.append(clone)
                    pending_percent = True
                    continue

                if "{{FOOTNOTES}}" in raw and "{{CHOICES}}" in raw:
                    qtype = str(rec["QUESTION_TYPE"])
                    if rec.get("SUB_QUESTIONS") and passage_tmpl is not None:
                        for sq in rec["SUB_QUESTIONS"]:
                            pq = deepcopy(passage_tmpl)
                            uniquify_ids(pq, state)
                            remove_mc_fallbacks(pq)
                            replace_in_runs(pq, [("{{PASSAGE}}", str(sq.get("text", "")))])
                            set_left_align(pq)
                            set_para_spacing(pq, before="100", after="0")
                            body.append(pq)
                            cq = deepcopy(passage_tmpl)
                            uniquify_ids(cq, state)
                            remove_mc_fallbacks(cq)
                            replace_in_runs(cq, [("{{PASSAGE}}", "\n".join(sq.get("choices", [])))])
                            set_left_align(cq)
                            apply_choice_format(cq)
                            body.append(cq)
                        continue
                    # 문삽: 선지가 '원문자(①~⑤)만'이면 지문에 이미 삽입 위치가 있으므로
                    # 선지 목록을 지문 아래에 '렌더하지 않는' 유형:
                    #   문삽(원문자만) / 어법·어휘(선지가 지문에 있어 밑줄로 표시 → 목록 불필요)
                    if is_bare_markers(rec["CHOICES"]) or choices_are_in_passage(rec["CHOICES"], rec["PASSAGE"]):
                        continue
                    # 각주는 본문 바로 아래(일반 분기)에 붙인다. 여기 토큰 문단(fp)은 버린다.
                    if passage_tmpl is not None:
                        cp = deepcopy(passage_tmpl)
                        uniquify_ids(cp, state)
                        remove_mc_fallbacks(cp)
                        if "순서" in qtype or "요약" in qtype:
                            ch = rec["CHOICES"]
                            choices_text = "   ".join(ch[:3]) + "\n" + "   ".join(ch[3:])
                        else:
                            choices_text = "\n".join(rec["CHOICES"])
                        replace_in_runs(cp, [("{{PASSAGE}}", choices_text)])
                        set_left_align(cp)  # 선지 양쪽정렬 해제 (글씨 크기는 줄이지 않음)
                        apply_choice_format(cp)   # 최종서식본 '선택지' 서식(줄간격·폭·원문자 글꼴)
                        #   (사용자 요청: 선지 글씨를 줄이지 말고 '간격'을 줄인다).
                        #   페이지를 미는 요인 = 지문+제시문 길이 또는 선지 총길이.
                        _pg = len(str(rec.get("PASSAGE", ""))) + len(str(rec.get("GIVEN", "")))
                        _ch = sum(len(c) for c in rec.get("CHOICES", []))
                        if (_pg > 950 or _ch > 200) and len(body) \
                                and body[-1].tag == P and not block_text(body[-1]).strip() \
                                and not is_empty_numbered_para(body[-1]):
                            body.remove(body[-1])
                        body.append(cp)
                    continue

                #   위치를 잡는 역할이라 지우면 단어가 엉뚱한 칸에 들어간다. 넘침의 진짜

                # 일반 복제 + 치환
                clone = deepcopy(child)
                uniquify_ids(clone, state)
                remove_mc_fallbacks(clone)
                fix_oneshot_label(clone)
                replace_in_runs(clone, mapping)
                # 제목 문단(박스 밖, 맨 위)에 형광+굵게+크게(#A)
                if "{{KOREAN_TITLE}}" in raw:
                    emphasize_title(clone, title_with_number(rec))
                if "{{ANNOTATED_SENTENCES}}" in raw:
                    # 어법/어휘(선지가 지문에 있는 유형): 인라인 원문자 ①~⑤ 와 밑줄을
                    #   대상을 잃는다(조교 요청). 예전에는 여기서 원문자를 지웠다.
                    is_gv = choices_are_in_passage(rec["CHOICES"], rec["PASSAGE"])
                    ul_targets = ([("①②③④⑤"[i], _choice_body(c))
                                   for i, c in enumerate(rec["CHOICES"][:5])] if is_gv else [])
                    phr = rec.get("UNDERLINE")
                    if not ul_targets and phr:
                        ul_targets = phr if isinstance(phr, list) else [phr]
                    write_sentence_triples(
                        clone, rec["SENTENCE_TRIPLES"],
                        strip_markers=False, underline_targets=ul_targets)
                    # 부동 문장표는 텍스트 흐름 위치에 둔다(vertAnchor=text). 페이지 앵커로 상단에
                    #   고정하면 흐름의 제목을 표 아래로 밀어내므로, 컬럼 브레이크로 단을 고정한
                    #   지금은 텍스트 앵커가 맞다. (제목은 흐름 맨 위, 표는 요지 박스 아래로 흐름)
                    tblp = clone.find(f".//{w('tblpPr')}")
                    if tblp is not None:
                        tblp.set(w("vertAnchor"), "text")
                        # 문장해석(부동 표)이 오른쪽 단까지 침범하면 그 옆의 '정답률 표'(인라인
                        #   표는 부동 도형 옆에 못 흘러 단 하단으로 밀림)를 밀어낸다. 표를 왼쪽
                        tblp.set(w("tblpX"), "1080")
                        tblp.set(w("rightFromText"), "0")
                        set_table_width(clone, ANALYSIS_TABLE_WIDTH)
                # 문삽/순서 제시문(주어진 문장·글): 지문 '위'에 별도 문단으로 — 들여쓰기 없이,
                #   그 뒤 빈 문단으로 지문과 한 줄 띄운다(지문 본문은 원래 들여쓰기 유지).
                if ("{{PASSAGE}}" in raw and str(rec.get("GIVEN", "")).strip()
                        and passage_tmpl is not None):
                    gv = deepcopy(passage_tmpl)
                    uniquify_ids(gv, state)
                    remove_mc_fallbacks(gv)
                    replace_in_runs(gv, [("{{PASSAGE}}", str(rec["GIVEN"]).strip())])
                    set_left_align(gv)                       # 제시문은 들여쓰기 없음
                    set_para_spacing(gv, after="0")
                    box_paragraph(gv)                        # 순서·문삽 제시문에 네모박스
                    #   여기서 긋는다. 지문 본문만 긋던 예전 코드는 제시문에 있는 (a) 를 놓쳤다.
                    _gphr = rec.get("UNDERLINE")
                    if _gphr:
                        underline_choice_spans(
                            gv, [], phrases=_gphr if isinstance(_gphr, list) else [_gphr])
                    body.append(gv)
                    blank = deepcopy(passage_tmpl)
                    uniquify_ids(blank, state)
                    remove_mc_fallbacks(blank)
                    replace_in_runs(blank, [("{{PASSAGE}}", "")])
                    set_para_spacing(blank, before="0", after="0")
                    body.append(blank)
                if "{{PASSAGE}}" in raw:
                    apply_passage_right_indent(clone)
                    # 순서: 지문 (A)(B)(C) 는 들여쓰기 없이 시작.
                    if "순서" in str(rec["QUESTION_TYPE"]):
                        remove_first_line_indent(clone)
                    # 어법/어휘: 지문 안 선지 구간에 밑줄을 직접 긋는다(선지 목록은 위에서 생략됨).
                    if choices_are_in_passage(rec["CHOICES"], rec["PASSAGE"]):
                        underline_choice_spans(clone, rec["CHOICES"])
                    elif str(rec.get("UNDERLINE", "")).strip() or isinstance(rec.get("UNDERLINE"), list):
                        phr = rec["UNDERLINE"]
                        underline_choice_spans(clone, [], phrases=phr if isinstance(phr, list) else [phr])
                body.append(clone)
                # 각주: 본문(PASSAGE) 바로 아래에 붙인다. 본문과 '같은 서식'(passage 스타일)으로
                #   만들고 우측 정렬만 준다. 각주 없으면 안 만든다.
                if ("{{PASSAGE}}" in raw and str(rec.get("FOOTNOTES", "")).strip()
                        and passage_tmpl is not None):
                    fn = deepcopy(passage_tmpl)
                    uniquify_ids(fn, state)
                    remove_mc_fallbacks(fn)
                    # 각주는 무조건 한 줄로 이어 붙인다(조교 요청). 항목 사이는 스페이스 두 칸으로
                    #   살짝 띄우고, 줄바꿈은 넣지 않는다 — 길면 워드가 알아서 접는다.
                    #   각 항목이 통째로 남으므로 원문 무결성 대조(부분문자열)도 그대로 통과한다.
                    fn_text = "  ".join(
                        ln.strip() for ln in str(rec["FOOTNOTES"]).splitlines() if ln.strip()
                    )
                    replace_in_runs(fn, [("{{PASSAGE}}", fn_text)])
                    fn_ppr = fn.find(PPR)
                    if fn_ppr is None:
                        fn_ppr = etree.Element(PPR)
                        fn.insert(0, fn_ppr)
                    for old in fn_ppr.findall(w("jc")):
                        fn_ppr.remove(old)
                    etree.SubElement(fn_ppr, w("jc")).set(w("val"), "right")
                    #   다음 페이지로 밀리지 않게 본문 페이지 높이를 아낀다.
                    for fr in fn.iter(R):
                        frpr = fr.find(RPR)
                        if frpr is None:
                            frpr = etree.Element(RPR)
                            fr.insert(0, frpr)
                        for sztag in ("sz", "szCs"):
                            for old in frpr.findall(w(sztag)):
                                frpr.remove(old)
                            etree.SubElement(frpr, w(sztag)).set(w("val"), "20")
                    set_para_spacing(fn, before="0", after="0")
                    body.append(fn)

                # 요약: 요약문(빈칸 A·B 포함)을 지문(+각주) 아래, 선지 위에 같은 서식으로 둔다.
                if ("{{PASSAGE}}" in raw and str(rec.get("SUMMARY_SENTENCE", "")).strip()
                        and passage_tmpl is not None):
                    sm = deepcopy(passage_tmpl)
                    uniquify_ids(sm, state)
                    remove_mc_fallbacks(sm)
                    replace_in_runs(sm, [("{{PASSAGE}}", str(rec["SUMMARY_SENTENCE"]).strip())])
                    remove_first_line_indent(sm)        # 요약문은 들여쓰기 없음(글씨 크기는 지문과 동일)
                    set_para_spacing(sm, before="100")
                    body.append(sm)

            #   (템플릿에 빈 전개방식 텍스트박스가 없거나 소비 실패) 조용히 빈 박스로
            #   내보내지 말고 멈춘다. — '전개방식이 안 나오는' 사고를 물리적으로 차단.
            if needs_tree and not structure_written:
                raise RuntimeError(
                    f"[전개방식 누락] record[{ri}] (NUMBER={rec.get('NUMBER','?')}, "
                    f"유형={rec.get('QUESTION_TYPE','?')}): 데이터는 있는데 전개방식 트리가 "
                    f"박스에 주입되지 못했습니다. 템플릿의 전개방식 빈 텍스트박스(도형)를 "
                    f"확인하세요(분석 페이지에 전개방식이 비어 나갈 뻔함)."
                )

            # 문제(레코드) 끝: 마지막 분석 섹션(D)을 닫아 다음 문제와 분리. 마지막 문제 제외.
            # 이 sectPr 은 분석 섹션을 닫는 것이라 원본 type(continuous) 유지
            # (새 페이지 분리는 위의 첫 섹션 nextPage 가 담당 → 빈 페이지 안 생김).
            if final_sectpr is not None and ri != len(records) - 1:
                sep = etree.Element(P)
                ppr = etree.SubElement(sep, PPR)
                ppr.append(deepcopy(final_sectpr))
                body.append(sep)

        if final_sectpr is not None:
            body.append(final_sectpr)

        # 마무리 정리(전 페이지 대상):
        #   ① 위쪽 여백 축소 → 상단 공백 줄여 설명 자리 확보.
        #   ② 전개방식 헤더 위 '점선 가로줄'(의도치 않은 선) 제거.
        set_top_margin(root, TOP_MARGIN_TWIPS)
        widen_analysis_boxes(root, ANALYSIS_BOX_WIDTH)
        remove_caricature(root, doc_xml)
        remove_dashed_hairlines(root)

        tree.write(str(doc_xml), encoding="UTF-8", xml_declaration=True, standalone=True)

        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
            for fp in sorted(tmp_dir.rglob("*")):
                if fp.is_file():
                    zf.write(fp, fp.relative_to(tmp_dir).as_posix())


# ───────────────────────────────────────────────────────────────────
#  원샷모음 docx (선생님용) — 서식본에서 비운 one shot 만 따로 모은다
# ───────────────────────────────────────────────────────────────────
def _docx_run(para, text, *, size=12, bold=False, color=None):
    """python-docx run 에 한글(맑은 고딕)·영문 폰트를 함께 건다."""
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    run = para.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "맑은 고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    return run


def create_oneshot_doc(records: list, output: Path, pyeongga: bool = True) -> None:
    """(설명 생략 — 공개본)"""
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    year = str(records[0].get("YEAR", "")).strip() if records else ""
    _docx_run(title, f"[One Shot] 모음  {year}".rstrip(), size=19, bold=True, color="990000")

    for rec in records:
        head = doc.add_paragraph()
        # 정답률 데이터가 없는 문항은 순위·비율 표기를 뺀다.
        if pyeongga or has_rate_data(rec):
            head_txt = (
                f"{rec['RANK']}위 · {rec.get('YEAR','')} {rec['NUMBER']}번 "
                f"({rec['CORRECT_RATE']})  {rec['KOREAN_TITLE']}"
            )
        else:
            head_txt = f"{rec['NUMBER']}번  {rec['KOREAN_TITLE']}"
        _docx_run(head, head_txt, size=13, bold=True, color="333366")

        body = doc.add_paragraph()
        _docx_run(body, "[One Shot] ", size=12, bold=True, color="1F4E79")
        _docx_run(body, rec["ONE_SHOT"], size=12)

        doc.add_paragraph()  # 문항 사이 간격

    output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output))


# ───────────────────────────────────────────────────────────────────
#  생성 후 구조 검증
# ───────────────────────────────────────────────────────────────────
def verify_output(output: Path, n: int, n_ranked: int | None = None) -> list[str]:
    issues = []
    xml = zipfile.ZipFile(output).read("word/document.xml").decode("utf-8")
    leftover = _LEFTOVER_TOKEN_RE.findall(xml)
    if leftover:
        issues.append(f"미치환 템플릿 토큰 남음: {sorted(set(leftover))}")
    if "comment]" in xml:
        issues.append("'comment]' 문자열이 남음")
    # 도형/이미지 docPr id 중복 = Word "복구 필요" 에러의 원인 → 자동 차단
    import collections as _c, re as _re
    _dup = [k for k, v in _c.Counter(_re.findall(r'<wp:docPr[^>]*\bid="(\d+)"', xml)).items() if v > 1]
    if _dup:
        issues.append(f"도형 docPr id 중복: {_dup}")
    #   (내용은 원샷모음 docx 로만 나가고, 그쪽은 verify_oneshot_doc 가 따로 검사한다.)
    expect_labels = [("[One Shot]", 0), ("핵심 풀이", 0)]
    # '오답 상위' 라벨은 정답률(오답률순)이 있는 문항에만 쓰이고, 없으면 다른 문구를 쓴다.
    expect_labels.append(("오답 상위", n if n_ranked is None else n_ranked))
    for label, want in expect_labels:
        got = xml.count(label)
        if got != want:
            issues.append(f"'{label}' {got}개 (기대 {want})")
    for bad in ("[??]", "?? ??", "[??'s"):
        if bad in xml:
            issues.append(f"깨진 라벨 '{bad}' 남음")
    # 단어: 유의어(유)는 대부분 단어에 있으나 반의어(반)는 선택적이라 강제하지 않음.
    if xml.count("(유)") < 10 * n:
        issues.append(f"(유) 표기가 너무 적음: {xml.count('(유)')}개 (문항당 10+ 권장)")
    # python-docx 로 열리는지
    try:
        import docx  # noqa
        docx.Document(str(output))
    except Exception as exc:  # pragma: no cover
        issues.append(f"python-docx 열기 실패: {exc}")
    return issues


def check_percentage_format(records: list) -> list[str]:
    """(설명 생략 — 공개본)"""
    issues = []
    for i, rec in enumerate(records):
        for j, pct in enumerate(rec.get("PERCENTAGES", [])):
            if not _PCT_RE.match(str(pct).strip()):
                issues.append(
                    f"PERCENTAGES 형식 의심 (NUMBER={rec.get('NUMBER', '?')})"
                    f"[{j}]: '{pct}' (예상: '17%')"
                )
    return issues


def _snippet(text: str, n: int = 24) -> str:
    """ONE_SHOT 의 '가운데' 조각을 검증용으로 뽑는다(줄바꿈 제거).

    앞부분('[정답]'·'빈칸 뒤 …' 등)은 ANSWER_ANALYSIS 같은 다른 필드와 첫머리가
    겹쳐 오탐(서식본에 one shot 이 남았다는 거짓 경고)을 낸다. 그래서 다른 필드와
    겹칠 일이 거의 없는 '가운데'에서 떼어 낸다. 누출 검사와 존재 검사가 같은 _snippet
    을 쓰므로 둘 다 일관되게 동작한다.
    """
    t = " ".join(str(text).split())
    if len(t) <= n:
        return t
    start = max(0, len(t) // 2 - n // 2)
    return t[start:start + n]


def _leak_norm(text: str) -> str:
    """누출 검사용 정규화: 공백 축약 + 소문자."""
    return " ".join(str(text).split()).lower()


def _record_text_except(rec: dict, exclude_key: str) -> str:
    """record 의 'exclude_key 를 뺀' 모든 문자열을 이어 붙인다(서식본에 정상 포함되는 텍스트)."""
    parts: list[str] = []
    for k, v in rec.items():
        if k == exclude_key:
            continue
        if isinstance(v, str):
            parts.append(v)
        elif isinstance(v, list):
            for x in v:
                if isinstance(x, str):
                    parts.append(x)
                elif isinstance(x, (list, tuple)):
                    parts.extend(str(y) for y in x)
    return " ".join(parts)


def verify_form_oneshot_blank(output: Path, records: list) -> list[str]:
    """서식본에 one shot '내용'이 새어 들어가지 않았는지 확인(라벨은 남아 있어야 정상).

    핵심: ANSWER_ANALYSIS 등 '다른 필드'는 서식본에 정상적으로 들어가고, one shot 과
    같은 영어구·한글 설명을 공유해 단순 부분일치로는 오탐이 난다. 그래서 'one shot 에만
    있고 다른 필드엔 없는' 구간을 찾아 그것이 문서에 나타날 때만 누출로 판정한다.
    """
    issues = []
    xml = zipfile.ZipFile(output).read("word/document.xml").decode("utf-8")
    doc_txt = _leak_norm(" ".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", xml, re.S)))
    L = 30
    for rec in records:
        one = _leak_norm(rec.get("ONE_SHOT", ""))
        if len(one) < 12:
            continue
        other = _leak_norm(_record_text_except(rec, "ONE_SHOT"))
        leak = None
        for i in range(0, len(one) - 11, 4):
            window = one[i:i + L]
            if len(window) >= 12 and window not in other and window in doc_txt:
                leak = window
                break
        if leak:
            issues.append(f"서식본에 one shot 내용이 남음 (NUMBER={rec['NUMBER']}): '{leak}…'")
    return issues


def verify_oneshot_doc(output: Path, records: list) -> list[str]:
    """(설명 생략 — 공개본)"""
    issues = []
    try:
        import docx
        full = "\n".join(p.text for p in docx.Document(str(output)).paragraphs)
    except Exception as exc:  # pragma: no cover
        return [f"원샷모음 열기 실패: {exc}"]
    for rec in records:
        snip = _snippet(rec["ONE_SHOT"])
        if snip and snip not in full:
            issues.append(f"원샷모음에 one shot 누락 (NUMBER={rec['NUMBER']}): '{snip}…'")
    return issues


def _no_dup_pairs(pairs: list) -> dict:
    """JSON 객체에 같은 키가 두 번 나오면 즉시 에러.

    오늘 같은 사고('옛 샘플 위에 새 데이터를 부분 수정으로 끼워 넣어' 한 객체 안에
    RAW/NUMBER/YEAR 가 두 벌씩 남는 병합 깨짐)를 파싱 단계에서 잡는다.
    (json 기본 동작은 '마지막 키만 남기고 조용히 덮어쓰기'라 사고가 숨는다.)
    """
    seen: dict = {}
    for k, v in pairs:
        if k in seen:
            raise ValueError(
                f"한 객체 안에 키 '{k}' 가 중복됨 — 옛 데이터와 새 데이터가 병합됐을 "
                f"가능성이 큽니다. JSON 을 부분 수정하지 말고 '통째로 새로' 쓰세요."
            )
        seen[k] = v
    return seen


def load_records_strict(path: Path) -> list:
    """입력 JSON 을 중복키 금지로 읽는다(병합 깨짐 차단). 실패 시 ValueError."""
    records = json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=_no_dup_pairs)
    if not isinstance(records, list) or not records:
        raise ValueError("입력 JSON 이 비어 있거나 배열이 아님")
    return records


def check_year_consistency(records: list, require_format: bool = True) -> list:
    """YEAR 형식(사설 모의고사는 생략 가능) + 한 파일에 한 시험만 담겼는지 검사(치명적)."""
    issues: list = []
    years = {str(r.get("YEAR", "")) for r in records}
    if require_format:
        for y in years:
            if not YEAR_RE.match(y):
                issues.append(
                    f"YEAR 형식 오류: {y!r} — 'YYYY.MM(YY시행)' 형식이어야 함 "
                    f"(점 앞 4자리=학년도). 예: 2023.09(22시행)"
                )
    if len(years) > 1:
        issues.append(
            f"한 파일에 연도가 여러 개 섞임: {sorted(years)} — 한 시험만 담으세요 "
            f"(옛 데이터 잔류/병합 사고 의심)."
        )
    return issues


_CIRCLED_MARKS = "①②③④⑤⑥⑦⑧⑨⑩"
_SUMMARY_MAX = 60          # [요약] 한 줄 상한(공백 포함). 넘으면 빌드 중단.


def check_format_structure(records: list) -> list:
    """유형별 '서식 구조'가 올바른지 검사(치명적).

    원문 누락이 아니라 **모양이 틀어지는** 사고를 막는다 — 어떤 클로드가 json 을 만들든
    출력 서식이 일정하게 나오게 하는 안전장치. (조교는 붙여넣기만 하고, json 구조는 분석하는
    클로드가 만든다 → 그 클로드가 실수해도 빌드가 잡아야 한다.)

    가드:
    - 어법/어휘(밑줄 오류) 유형 = 유형에 '어법/어휘' + UNDERLINE·SUB_QUESTIONS 없음:
            잘못 렌더된다. 선지를 지문 원문 그대로(인라인 원문자 포함) 써야 함.
    (함축·밑줄의미는 UNDERLINE 을, 장문은 SUB_QUESTIONS 를 쓰므로 위 어법/어휘 가드에서 제외된다.)
    """
    issues: list = []
    for rec in records:
        qt = str(rec.get("QUESTION_TYPE", ""))
        nid = rec.get("NUMBER", "?")
        has_ul = bool(rec.get("UNDERLINE"))
        has_sq = bool(rec.get("SUB_QUESTIONS"))
        is_grammar_vocab = (("어법" in qt or "어휘" in qt) and not has_ul and not has_sq)
        if is_grammar_vocab:
            ch = rec.get("CHOICES", []) or []
            pa = str(rec.get("PASSAGE", ""))
            if not choices_are_in_passage(ch, pa):
                issues.append(
                    f"[서식오류] NUMBER={nid}: 어법/어휘인데 선지가 지문에서 안 찾아짐 → "
                    f"confirming 이 '빈칸식'으로 잘못 나간다. 선지를 지문 원문 그대로(인라인 원문자 포함) 쓰세요."
                )
            n_mark = sum(pa.count(c) for c in _CIRCLED_MARKS)
            if n_mark < len(ch):
                issues.append(
                    f"[서식오류] NUMBER={nid}: 어법/어휘 지문에 인라인 원문자(①~⑤)가 부족"
                    f"({n_mark}/{len(ch)}) → 번호가 안 박힌다. 원문의 원문자를 PASSAGE 에 그대로 두세요."
                )
        summ = _strip_label(str(rec.get("SUMMARY", ""))).strip()
        if len(summ) > _SUMMARY_MAX:
            issues.append(
                f"[서식오류] NUMBER={nid}: [요약]이 {len(summ)}자로 너무 길다"
                f"({_SUMMARY_MAX}자 이내 한 문장). 결론만 남기고 줄이세요 → \"{summ[:40]}…\""
            )
        if "요약" in qt and not str(rec.get("SUMMARY_SENTENCE", "")).strip():
            issues.append(
                f"[서식오류] NUMBER={nid}: 요약 유형인데 SUMMARY_SENTENCE 가 비었다 "
                f"(요약문이 지문 아래 렌더되지 못함)."
            )
        if ("순서" in qt or "문삽" in qt or "삽입" in qt) and not str(rec.get("GIVEN", "")).strip():
            issues.append(
                f"[서식오류] NUMBER={nid}: 순서/문삽 유형인데 GIVEN(주어진 글·문장)이 비었다."
            )
        issues += check_structure_nodes(rec, nid, qt)
        issues += check_ae_underline(rec, nid)
        issues += check_question_type(rec, nid)
        issues += check_insertion_given_node(rec, nid)
    return issues


# ── QUESTION_TYPE 허용 키워드 ──────────────────────────────────────
#   유형 문자열 하나에 렌더 분기(밑줄·박스 레이아웃·구조도 펼침/묶음)가 전부 걸려 있다.
#   내가 즉석에서 새 이름을 지어 쓰면 가드가 엉뚱하게 판정하므로, 아는 낱말이 하나도
#   없으면 멈춘다. (조합은 자유 — '장문(주제·낱말)' 처럼 붙여 써도 된다.)
_TYPE_KEYWORDS = (
    "어법", "어휘", "낱말", "빈칸", "순서", "문삽", "삽입", "요약", "제목", "주제",
    "요지", "목적", "주장", "내용", "밑줄", "함축", "지칭", "무관", "삭제", "장문", "도표", "심경",
)


def check_question_type(rec: dict, nid: str) -> list:
    """QUESTION_TYPE 이 아는 유형인지(치명적). 오타·즉석 작명으로 분기가 틀어지는 것을 막는다."""
    qt = str(rec.get("QUESTION_TYPE", ""))
    if any(k in qt for k in _TYPE_KEYWORDS):
        return []
    return [
        f"[서식오류] NUMBER={nid}: QUESTION_TYPE='{qt}' 를 모르겠다. "
        f"아는 낱말을 하나는 넣으세요 → {', '.join(_TYPE_KEYWORDS)}"
    ]


def check_insertion_given_node(rec: dict, nid: str) -> list:
    """문삽 유형의 전개방식에 '[제시문] … 이유:' 노드가 있는지(치명적).

    제시문 노드를 들어갈 자리에 끼워 넣어야 트리만 보고도 정답 자리가 보인다(조교 요청).
    """
    qt = str(rec.get("QUESTION_TYPE", ""))
    if "장문" in qt or not ("문삽" in qt or "삽입" in qt):
        return []
    for line in str(rec.get("STRUCTURE_ANALYSIS", "")).split("\n"):
        if "[제시문]" in line and "이유:" in line:
            return []
    return [
        f"[서식오류] NUMBER={nid}: 문삽인데 전개방식에 '[제시문] … 이유: …' 노드가 없다. "
        f"제시문을 '들어갈 자리'에 끼우고 `├── [제시문] 이유: 핵심어 = 핵심 내용` 형식으로 쓰세요."
    ]


_AE_LABELS = ("(a)", "(b)", "(c)", "(d)", "(e)")


def check_ae_underline(rec: dict, nid: str) -> list:
    """장문의 (a)~(e) 가 지문에 있는데 UNDERLINE 이 비었으면 중단(치명적).

    그때 빌드는 그냥 통과해 버려서 조교가 직접 발견해야 했다 → 이제 파일이 안 나온다.
    """
    text = str(rec.get("PASSAGE", "")) + str(rec.get("GIVEN", ""))
    if not all(lab in text for lab in _AE_LABELS):
        return []
    ul = rec.get("UNDERLINE") or []
    n_ul = len(ul) if isinstance(ul, list) else (1 if str(ul).strip() else 0)
    if n_ul >= len(_AE_LABELS):
        return []
    return [
        f"[서식오류] NUMBER={nid}: 지문에 (a)~(e) 가 있는데 UNDERLINE 이 {n_ul}개다(5개 필요). "
        f"어휘든 지칭이든 (a)~(e) 는 전부 밑줄 대상이다 — 1쪽은 원문과 똑같아야 한다. "
        f'UNDERLINE 을 [["(a)","he"], ["(b)","He"], …] 처럼 [라벨, 낱말] 쌍으로 주세요.'
    ]


# ── 전개방식 트리 노드 규칙 (조교 확정) ────────────────────────────────
#   '장문' 을 먼저 본다 — '장문(주제·낱말)' 처럼 두 키워드가 겹칠 때 장문이 이긴다.
_EXPAND_KEYS = ("순서", "문삽", "삽입", "어휘", "낱말", "삭제", "무관")
_NODE_MAX_LEN = 40          # 한 노드 = 한 줄. 트리문자·원문자·라벨·신호어 전부 포함.
_MERGED_NODE_RANGE = (3, 8)  # 묶음 유형의 노드 줄 수


def structure_mode(qtype: str) -> str:
    """이 유형의 전개방식 트리를 '펼침'으로 쓸지 '묶음'으로 쓸지."""
    qt = str(qtype)
    if "장문" in qt:
        return "묶음"
    return "펼침" if any(k in qt for k in _EXPAND_KEYS) else "묶음"


def check_structure_nodes(rec: dict, nid: str, qt: str) -> list:
    """전개방식 트리의 노드 수·길이 규칙(치명적).

    경고로 두면 그냥 지나쳐 조교가 문항마다 손으로 쪼개야 했다(실제 사고). 그래서 중단시킨다.
    ⛔ 이 검사를 무르게 고쳐서 통과시키지 말 것 — 고칠 곳은 STRUCTURE_ANALYSIS 쪽이다.
    """
    out: list = []
    lines = [l for l in str(rec.get("STRUCTURE_ANALYSIS", "")).split("\n") if l.strip()]
    nodes = [l for l in lines if l.lstrip().startswith(("├", "└"))]
    for l in lines:
        if len(l) > _NODE_MAX_LEN:
            out.append(
                f"[서식오류] NUMBER={nid}: 전개방식 줄이 {len(l)}자로 김"
                f"({_NODE_MAX_LEN}자 이내). 두 줄로 접혀 박스가 넘친다 → \"{l[:34]}…\""
            )
    mode = structure_mode(qt)
    n_sent = len(rec.get("SENTENCE_TRIPLES") or [])
    if mode == "펼침":
        if len(nodes) != n_sent:
            out.append(
                f"[서식오류] NUMBER={nid}: {qt} 는 '펼침' 유형인데 "
                f"노드 {len(nodes)}개 ≠ 문장 {n_sent}개. 원문자 하나당 한 줄로 쪼개세요 "
                f"(순서·문삽·어휘·삭제는 문장 연결이 곧 풀이 근거라 묶으면 위계가 뭉개진다)."
            )
    else:
        lo, hi = _MERGED_NODE_RANGE
        if nodes and not (lo <= len(nodes) <= hi):
            out.append(
                f"[서식오류] NUMBER={nid}: {qt} 는 '묶음' 유형인데 노드가 {len(nodes)}개다 "
                f"({lo}~{hi}줄). 인접 문장을 ①②③ 처럼 묶으세요."
            )
    return out


# ── RAW ↔ 원본 PDF 대조 ────────────────────────────────────────────
#   원문 무결성 검사는 'RAW 가 진실'이라고 가정하고 RAW↔docx 를 본다. 그런데 RAW 를
#   만드는 건 PDF 를 읽고 옮겨 적는 모델이라, 옮기다 틀리면 RAW 와 docx 가 사이좋게
#   같이 틀린 채 통과한다(순환). 그래서 진실의 원천을 PDF 로 끌어올린다.
_PDF_FURNITURE = re.compile(r"^\s*(공\s*통|영어\s*영역|공\s*란|\d{1,3}|={3}.*={3})\s*$")


def extract_pdf_text(paths: list) -> str:
    """문제지 PDF 들에서 텍스트를 뽑아 하나로 잇는다(머리말·쪽번호 같은 페이지 장식은 제거).

    페이지를 걸친 문장이 둘로 쪼개져 가짜 누락경보가 난다.
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise SystemExit(
            "⛔ --pdf 를 쓰려면 PyMuPDF 가 필요합니다: python -m pip install pymupdf"
        )
    chunks = []
    for p in paths:
        path = Path(p)
        if not path.exists():
            raise SystemExit(f"⛔ PDF 를 찾을 수 없음: {path}")
        with fitz.open(str(path)) as doc:
            for page in doc:
                for line in page.get_text().split("\n"):
                    if not _PDF_FURNITURE.match(line):
                        chunks.append(line)
    text = "\n".join(chunks)
    return re.sub(r"\[\s*\d+\s*점\s*\]", " ", text)


def crosscheck_raw_pdf(records: list, pdf_text: str) -> list:
    """각 record 의 RAW 가 '원본 PDF 안에 실제로 있는' 문장들인지 대조(치명적).

    ⛔ 막히면 검사기나 PDF 를 탓하지 말고 RAW 를 PDF 원문대로 고칠 것 —
       이 검사가 잡는 건 '옮겨 적다 바뀐 원문'이다.
    """
    # PDF 는 공백을 글자로 안 넣는 경우가 많다(한글 선지가 '①명작을이해하기…' 로 추출됨).
    #   그래서 양쪽에서 공백을 전부 지우고 대조한다 — 문장이 바뀌었는지만 보면 되므로
    #   띄어쓰기 차이로 가짜 경보가 나면 안 된다.
    def _clean(s: str) -> str:
        return re.sub(r"\s+", "", integrity._coverage_clean(s))

    hay = _clean(pdf_text)
    issues: list = []
    for rec in records:
        nid = rec.get("NUMBER", "?")
        bad = []
        for unit in integrity.split_units(rec.get("RAW", "")):
            norm = _clean(unit)
            # 글자(영문·한글)가 없는 토막(마커·구두점만)과 아주 짧은 토막은 대조 대상 아님
            if len(norm) < 8 or not re.search(r"[a-z가-힣]", norm):
                continue
            if norm not in hay:
                bad.append(unit)
        for unit in bad[:3]:
            issues.append(
                f"[원문불일치] NUMBER={nid}: RAW 의 이 문장이 PDF 원문에 없다 → \"{unit[:70]}…\""
            )
        if len(bad) > 3:
            issues.append(f"[원문불일치] NUMBER={nid}: … 외 {len(bad)-3}건")
    issues += _raw_omission_check(records, pdf_text, _clean)
    return issues


def _raw_omission_check(records: list, pdf_text: str, clean) -> list:
    """(공개본 — 수치가 든 설명은 뺐습니다)

    위 검사는 'RAW ⊂ PDF' 만 보므로 **문장을 통째로 지우면 통과**해 버린다. 그래서
    각 record 의 RAW 가 PDF 에서 차지하는 '구간'을 잡고, 그 구간 안의 PDF 문장이
    RAW 에 다 들어 있는지 되짚는다. (구간을 못 잡으면 조용히 건너뛴다 — 가짜 경보 방지.)
    """
    hay = clean(pdf_text)
    pdf_units = []                      # (정리된 문장, hay 안 시작위치)
    cursor = 0
    for unit in integrity.split_units(pdf_text):
        norm = clean(unit)
        if len(norm) < 8 or not re.search(r"[a-z가-힣]", norm):
            continue
        pos = hay.find(norm, cursor)
        if pos < 0:
            pos = hay.find(norm)
        if pos >= 0:
            pdf_units.append((unit, norm, pos))
            cursor = pos + len(norm)
    issues: list = []
    for rec in records:
        nid = rec.get("NUMBER", "?")
        raw_clean = clean(rec.get("RAW", ""))
        mine = [clean(u) for u in integrity.split_units(rec.get("RAW", ""))]
        mine = [m for m in mine if len(m) >= 8 and re.search(r"[a-z가-힣]", m)]
        if len(mine) < 2:
            continue
        spans = [hay.find(mine[0]), hay.find(mine[-1])]
        if spans[0] < 0 or spans[1] < 0:
            continue                    # 앞 검사에서 이미 걸렸다
        lo, hi = spans[0], spans[1] + len(mine[-1])
        missing = [u for u, norm, pos in pdf_units
                   if lo <= pos < hi and norm not in raw_clean]
        for u in missing[:3]:
            issues.append(
                f"[원문누락] NUMBER={nid}: PDF 에 있는 이 문장이 RAW 에서 빠졌다 → \"{u[:70]}…\""
            )
        if len(missing) > 3:
            issues.append(f"[원문누락] NUMBER={nid}: … 외 {len(missing)-3}건")
    return issues


def crosscheck_csv(records: list, csv_path: Path) -> tuple:
    """CSV(진실의 원천)와 JSON 의 정답·정답률·선지비율을 대조.

    반환: (fatal, info)
      fatal — CSV 와 어긋난 항목(엉뚱한 연도/지어낸 수치) → 빌드 중단 사유.
      info  — 대조를 생략한 항목(장문 합본 등) 안내.
    엉뚱한 연도로 만든 교재는 번호·수치가 CSV 와 안 맞아 여기서 전부 걸린다.
    """
    fatal: list = []
    info: list = []
    if not csv_path.exists():
        info.append(f"CSV 없음 — 수치 대조 생략: {csv_path}")
        return fatal, info
    rows = list(_csv.DictReader(csv_path.read_text(encoding="utf-8-sig").splitlines()))
    for rec in records:
        year = str(rec.get("YEAR", "")).strip()
        num = str(rec.get("NUMBER", "")).strip()
        if not num.isdigit():
            info.append(f"NUMBER={num!r} 은 단일 숫자가 아니어서 CSV 대조 생략(장문 합본?)")
            continue
        match = [r for r in rows
                 if (r.get("시험") or "").strip() == year
                 and (r.get("번호") or "").strip() == num]
        if not match:
            fatal.append(
                f"{num}번: CSV 에서 '시험={year}'·'번호={num}' 행을 못 찾음 "
                f"— 연도/번호 매칭 오류 의심(엉뚱한 연도?)."
            )
            continue
        r = match[0]
        csv_rank = (r.get("오답률순") or "").strip()
        rank_nums = re.findall(r"\d+", str(rec.get("RANK", "")))
        if csv_rank and csv_rank not in rank_nums:
            fatal.append(f"{num}번 RANK 불일치: JSON={rec.get('RANK')!r} CSV 오답률순={csv_rank}")
        try:
            if int(rec.get("CORRECT_ANSWER")) != int(r.get("정답", "")):
                fatal.append(f"{num}번 CORRECT_ANSWER 불일치: JSON={rec.get('CORRECT_ANSWER')} CSV={r.get('정답')}")
        except (TypeError, ValueError):
            fatal.append(f"{num}번 CORRECT_ANSWER 값 이상: {rec.get('CORRECT_ANSWER')!r}")
        if str(rec.get("CORRECT_RATE", "")).strip() != (r.get("정답률") or "").strip():
            fatal.append(f"{num}번 CORRECT_RATE 불일치: JSON={rec.get('CORRECT_RATE')!r} CSV={r.get('정답률')!r}")
        csv_pct = [(r.get(f"선지{i}") or "").strip() for i in range(1, 6)]
        json_pct = [str(p).strip() for p in (rec.get("PERCENTAGES") or [])]
        if json_pct != csv_pct:
            fatal.append(f"{num}번 PERCENTAGES 불일치: JSON={json_pct} CSV={csv_pct}")
    return fatal, info


def parse_args(argv) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="수능 TOP15 교재 빌드(서식본+원샷모음)")
    ap.add_argument("--in", dest="in_json", default=str(INPUT_JSON),
                    help="입력 JSON 경로(연도별 병렬 작업 시 학년도_월 파일명 권장)")
    ap.add_argument("--out", dest="out", default=str(OUTPUT), help="서식본 docx 출력 경로")
    ap.add_argument("--oneshot", dest="oneshot", default=str(OUTPUT_ONESHOT),
                    help="원샷모음 docx 출력 경로")
    ap.add_argument("--csv", dest="csv", default=str(CSV_PATH), help="정답률 CSV 경로")
    ap.add_argument("--no-csv-check", action="store_true",
                    help="CSV 수치 대조 생략(권장 안 함)")
    ap.add_argument("--pdf", dest="pdf", nargs="*", default=None,
                    help="원본 문제지 PDF 경로(여러 개 가능). RAW 가 PDF 원문과 같은지 대조한다.")
    return ap.parse_args(argv)


def main(argv=None) -> int:
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    args = parse_args(argv)
    in_json, out, oneshot = Path(args.in_json), Path(args.out), Path(args.oneshot)

    # ── 입력 정합성 가드(치명적): 병합 깨짐 / 연도 혼동 / 지어낸 수치를 빌드 전에 차단 ──
    try:
        records = load_records_strict(in_json)
    except ValueError as exc:
        print(f"⛔ 입력 JSON 오류 — 빌드 중단: {exc}", file=sys.stderr)
        return 4
    fatal = check_year_consistency(records, require_format=not args.no_csv_check)
    info: list = []
    # RAW 가 원본 PDF 와 같은지 — 진실의 원천을 '모델이 옮겨 적은 RAW' 에서 'PDF' 로 올린다.
    if args.pdf:
        fatal += crosscheck_raw_pdf(records, extract_pdf_text(args.pdf))
    if not args.no_csv_check:
        cf, ci = crosscheck_csv(records, Path(args.csv))
        fatal += cf
        info += ci
    if fatal:
        print("⛔ 입력 정합성 오류 — 빌드 중단(CSV/연도 불일치 또는 병합 깨짐):", file=sys.stderr)
        for it in fatal:
            print("  -", it, file=sys.stderr)
        return 4
    for it in info:
        print("  ·", it)

    def _rank_key(rec: dict) -> int:
        m = re.search(r"\d+", str(rec.get("RANK", "")))
        return int(m.group()) if m else 9999
    records.sort(key=_rank_key)

    try:
        build(TEMPLATE, out, records, pyeongga=not args.no_csv_check)
    except PermissionError:
        print(f"⛔ '{out}' 에 쓸 수 없음 — 이 docx 가 Word 에 열려 있지 않나요? "
              f"파일을 닫고 다시 실행하세요.", file=sys.stderr)
        return 5
    print(f"생성 완료(서식본): {out}")

    # RAW(붙여넣은 원문)의 모든 문장·선지가 docx 에 들어갔는지 대조.
    # 하나라도 빠지면 → 산출물 폐기 후 즉시 중단(누락된 교재는 절대 내보내지 않는다).
    missing = integrity.check_records(out, records, raw_key="RAW", id_key="NUMBER")
    # 추가: 문장별 해설이 지문 전 문장을 덮는지(해설 페이지에 원문 누락 금지).
    cov_records = [r for r in records
                   if not choices_are_in_passage(r.get("CHOICES", []), r.get("PASSAGE", ""))]
    missing += integrity.check_sentence_coverage(cov_records, id_key="NUMBER")
    # 추가: 주어진 문장(GIVEN, 문삽·순서)이 분석(SENTENCE_TRIPLES)에도 들어갔는지.
    #   GIVEN 은 PASSAGE 밖이라 위 두 검사의 사각지대 — 별도로 전 record 대상으로 본다.
    missing += integrity.check_given_in_triples(records, id_key="NUMBER")
    missing += integrity.check_given_in_triples(
        records, given_key="SUMMARY_SENTENCE", id_key="NUMBER",
        tag="요약문 해설 누락", what="요약문(SUMMARY_SENTENCE)")
    # 추가: 유형별 서식 구조 가드 — 원문 누락이 아니라 '모양이 틀어지는' 사고 차단
    #   (어법/어휘 빈칸식 오렌더·원문자 누락, 요약문 누락, 순서/문삽 GIVEN 누락).
    missing += check_format_structure(records)
    if missing:
        try:
            out.unlink()
        except OSError:
            pass
        kind = "서식오류" if all("[서식오류]" in str(m) for m in missing) else "원문 누락/서식오류"
        print(f"⛔ {kind} 감지 — 교재 생성 중단 (서식본 삭제됨):", file=sys.stderr)
        for it in missing:
            print("  -", it, file=sys.stderr)
        return 3

    try:
        create_oneshot_doc(records, oneshot, pyeongga=not args.no_csv_check)
    except PermissionError:
        print(f"⛔ '{oneshot}' 에 쓸 수 없음 — 이 docx 가 Word 에 열려 있지 않나요? "
              f"파일을 닫고 다시 실행하세요.", file=sys.stderr)
        return 5
    print(f"생성 완료(원샷모음): {oneshot}")

    n_ranked = sum(1 for r in records if not args.no_csv_check or has_rate_data(r))
    issues = verify_output(out, len(records), n_ranked=n_ranked)
    # 서식본에는 one shot 내용이 없어야 한다(라벨 박스만 남음).
    issues += verify_form_oneshot_blank(out, records)
    issues += verify_oneshot_doc(oneshot, records)
    # 따질 수치 자체가 없으므로 건너뛴다.
    issues += check_percentage_format(
        [r for r in records if not args.no_csv_check or has_rate_data(r)]
    )
    if issues:
        print("⚠️ 검증 경고:")
        for it in issues:
            print("  -", it)
        return 2
    print("✅ 구조 검증 통과 (서식본 one shot 비움 / 원샷모음 내용 정상 / placeholder 0)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
