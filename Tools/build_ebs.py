# -*- coding: utf-8 -*-
"""
EBS 지문 DOCX 생성기 (안정화 버전, JSON 기반).

설계 원칙 (build_suneung.py 와 동일 철학)
----------------------------------------
1. 템플릿(Example/EBS_auto_template.docx)의 이미지/도형/텍스트박스/배경/섹션/페이지
   구성을 그대로 보존하고, 토큰만 치환한다. 새로 그리지 않는다.
2. EBS 템플릿은 책 펼침면 구조: **좌페이지 유닛(블록 0~35) + 우페이지 유닛(블록 36~73)**
   이 한 세트로, 캐리커쳐/배경/로고 위치가 좌·우가 다르다(1-2-1-2 교대).
   → 지문을 2개씩 짝지어, 한 세트를 복제하고 좌=A·우=B 로 채운다.
   → 지문이 홀수면 마지막 세트의 우페이지는 토큰을 비운다(빈 우페이지).
3. 모든 토큰은 좌영역·우영역에 각각 1번씩만 등장하므로, 좌 블록엔 A 매핑, 우 블록엔
   B 매핑을 적용하면 이름 충돌 없이 채워진다.
4. 깨진 한글 라벨([??]=요지/주제/제목/요약/영독/어휘, [????]=전개방식)은 토큰과 함께
   같은 run 에 있으므로 composite 키로 복구한다.
5. lxml 로 파싱/직렬화(네임스페이스 prefix 보존) + 복제본 ID 유일화 + mc:Fallback 제거
   로 Word 호환성을 확보한다.

실행
----
    python Tools/build_ebs.py
    -> Documents/EBS_stable.docx 생성 + 구조 검증 로그 출력
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
import integrity  # 원문 무결성 검사 (제1원칙: 원문 절대 누락 금지)


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "Example" / "EBS_auto_template.docx"
INPUT_JSON = ROOT / "Tools" / "ebs_input.json"
OUTPUT = ROOT / "Documents" / "EBS_stable.docx"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# 좌영역(블록 0~LEFT_END-1) / 우영역(블록 LEFT_END~) 경계.
# block 36 = 우페이지 레슨 헤더 시작.
LEFT_END = 36


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


BODY = w("body")
P = w("p")
PPR = w("pPr")
R = w("r")
T = w("t")
BR = w("br")
SECTPR = w("sectPr")
TXBX = w("txbxContent")
DRAWING = w("drawing")

# JSON 에 반드시 있어야 하는 키 (= 템플릿 토큰)
# RAW = 조교가 붙여넣은 원문 지문 그대로의 복사본. 무결성 검사 기준.
REQUIRED_KEYS = [
    "RAW",
    "LESSON_NUMBER", "LESSON_TITLE", "GIST", "TOPIC", "TITLE_EN", "SUMMARY",
    "NUMBERED_SENTENCES", "KOREAN_TRANSLATIONS", "TEACHER_NOTE",
    "STRUCTURE_ANALYSIS", "STRUCTURE_TYPE", "VOCABULARY",
]
LIST_KEYS = {"NUMBERED_SENTENCES", "KOREAN_TRANSLATIONS", "VOCABULARY"}
# 나머지 필수 키는 모두 문자열이어야 한다.
STR_KEYS = set(REQUIRED_KEYS) - LIST_KEYS

# 이 빌드가 치환하는 '템플릿 토큰' 전체. verify_output 은 사용자 원문에 우연히 들어 있는
# {{ x }} 가 아니라 '미치환된 이 토큰들'만 잔여 검사한다(#4).
TEMPLATE_TOKENS = [
    "LESSON_NUMBER", "LESSON_TITLE", "GIST", "TOPIC", "TITLE_EN", "SUMMARY",
    "STRUCTURE_TYPE", "STRUCTURE_ANALYSIS", "TEACHER_NOTE",
    "KOREAN_TRANSLATIONS", "NUMBERED_SENTENCES", "VOCABULARY",
]
_LEFTOVER_TOKEN_RE = re.compile(r"\{\{(" + "|".join(TEMPLATE_TOKENS) + r")\}\}")


def _typename(val) -> str:
    return type(val).__name__


# ───────────────────────────────────────────────────────────────────
#  검증
# ───────────────────────────────────────────────────────────────────
def validate_passage(idx: int, p: dict) -> None:
    where = f"passage[{idx}] (LESSON={p.get('LESSON_NUMBER', '?')})"
    for key in REQUIRED_KEYS:
        if key not in p:
            raise ValueError(f"{where}: 필수 키 '{key}' 누락")
        val = p[key]
        # 자료형 검사 (조용한 ['...'] repr 출력 방지)
        if key in STR_KEYS and not isinstance(val, str):
            raise ValueError(
                f"{where}: '{key}' 는 문자열이어야 하는데 {_typename(val)} 가 옴 "
                f"(str 자리에 리스트/딕트를 넣지 말 것)"
            )
        if key in LIST_KEYS:
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
        if isinstance(val, str) and not val.strip():
            raise ValueError(f"{where}: '{key}' 값이 비어 있음")
        if isinstance(val, list) and len(val) == 0:
            raise ValueError(f"{where}: '{key}' 리스트가 비어 있음")
    if len(p["NUMBERED_SENTENCES"]) != len(p["KOREAN_TRANSLATIONS"]):
        raise ValueError(
            f"{where}: 영어문장({len(p['NUMBERED_SENTENCES'])})과 "
            f"해석({len(p['KOREAN_TRANSLATIONS'])}) 개수가 다름"
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
    """좌→우 1회성 치환(#3): 이미 채워 넣은 결과(new)는 다시 검사하지 않는다.

    사용자 문자열(원문 지문 등)에 우연히 들어 있던 '{{...}}' 가 뒤이은 토큰 치환에서
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
    """run 단위 치환. run 의 '직속' w:t 만 본다(텍스트박스 평면화 방지).

    drawing/그림을 담은 run 은 건드리지 않는다. 텍스트박스 내부 run 은
    element.iter(R) 가 별도로 방문하므로 함께 처리된다(EBS 토큰은 박스 안에 있음).
    치환은 apply_tokens_once 로 1회성 처리 — 채워 넣은 텍스트는 재치환하지 않는다(#3).
    """
    for run in element.iter(R):
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
#  치환 매핑 (composite 키 먼저: 깨진 한글 라벨 + 토큰 동시 복구)
# ───────────────────────────────────────────────────────────────────
def build_mapping(p: dict) -> list[tuple[str, str]]:
    vocab = "\n".join(p["VOCABULARY"])             # 단어 1줄/개
    sentences = " ".join(p["NUMBERED_SENTENCES"])  # 원문 연속
    trans = " ".join(p["KOREAN_TRANSLATIONS"])     # 해석 연속
    return [
        # ── 라벨+토큰 (단일 run, 긴 키 우선) ──
        ("[??] [{{LESSON_NUMBER}}] {{LESSON_TITLE}}",
         f"[영독] [{p['LESSON_NUMBER']}] {p['LESSON_TITLE']}"),
        ("[??] {{GIST}}", f"[요지] {p['GIST']}"),
        ("[??] {{TOPIC}}", f"[주제] {p['TOPIC']}"),
        ("[??] {{TITLE_EN}}", f"[제목] {p['TITLE_EN']}"),
        ("[??] {{SUMMARY}}", f"[요약] {p['SUMMARY']}"),
        ("[????] : {{STRUCTURE_TYPE}}", f"[전개방식] : {p['STRUCTURE_TYPE']}"),
        # 어휘 라벨+토큰 (run 안에서 줄바꿈이 w:br 또는 실제 \n 인 두 경우 모두 대비)
        ("[??]\n{{VOCABULARY}}", f"[어휘]\n{vocab}"),
        ("[??]{{VOCABULARY}}", f"[어휘]\n{vocab}"),
        # ── 단일 토큰 ──
        ("{{STRUCTURE_ANALYSIS}}", p["STRUCTURE_ANALYSIS"]),
        ("{{TEACHER_NOTE}}", p["TEACHER_NOTE"]),
        ("{{KOREAN_TRANSLATIONS}}", trans),
        ("{{NUMBERED_SENTENCES}}", sentences),
        ("{{VOCABULARY}}", vocab),  # 라벨과 분리된 경우 대비
    ]


# ───────────────────────────────────────────────────────────────────
#  메인 빌드
# ───────────────────────────────────────────────────────────────────
def set_textbox_autofit(element: etree._Element) -> None:
    """txbxContent 를 가진 drawing(텍스트박스)만 내용맞춤으로 바꾼다.

    수능 빌드(set_textbox_autofit)에서 검증된 방식 이식:
      1) 박스 bodyPr 의 noAutofit/normAutofit → a:spAutoFit 으로 교체
      2) 그 박스의 wp:extent / a:ext 높이(cy)를 최소값으로 줄임
    → Word 가 내용만큼만 박스를 키워, 고정 크기 박스가 예약하던 빈 공간이
      사라지고 빈 페이지가 없어진다(#2 원칙 위배 없음 — 글자는 안 잘림).

    캐리커쳐·로고·배경 그림 drawing(txbxContent 없음)은 건드리지 않는다.
    좌(A)·우(B) 양쪽 박스 모두 element.iter(DRAWING) 으로 함께 처리된다.
    """
    MIN_CY = 274320  # 0.3in
    for drawing in element.iter(DRAWING):
        if next((t for t in drawing.iter(TXBX)), None) is None:
            continue  # 텍스트박스가 아닌 그림 drawing 은 크기 보존
        for bp in drawing.iter():
            if etree.QName(bp).localname != "bodyPr":
                continue
            for child in list(bp):
                if etree.QName(child).localname in ("noAutofit", "spAutoFit", "normAutofit"):
                    bp.remove(child)
            etree.SubElement(bp, f"{{{A}}}spAutoFit")
        for el in drawing.iter():
            if etree.QName(el).localname in ("extent", "ext") and el.get("cy"):
                el.set("cy", str(MIN_CY))


def _is_empty_filler(p: etree._Element) -> bool:
    """치환 후 '순수 빈 여백 문단'인지 판정 (그림·텍스트·섹션 없음).

    EBS 템플릿은 고정 레이아웃을 위해 본문 흐름에 빈 문단 수십 개를 스페이서로
    깔아 둔다. 지문이 짧으면 이 스페이서가 흐름을 페이지 밖으로 밀어 빈 페이지를
    만든다(박스는 wrapNone 부유 개체라 페이지 수에 영향이 없음 → 진짜 원인은 이 문단).
    그림/텍스트/섹션이 있는 문단(박스·각주·푸터 '- 6 -'·sectPr)은 보존한다.
    """
    if etree.QName(p).localname != "p":
        return False
    if p.find(".//" + SECTPR) is not None:
        return False
    if any(etree.QName(e).localname in ("drawing", "pict", "AlternateContent")
           for e in p.iter()):
        return False
    text = "".join(t.text or "" for t in p.iter(T))
    return not text.strip()


def _fill_blocks(blocks, mapping, body, state):
    for child in blocks:
        clone = deepcopy(child)
        uniquify_ids(clone, state)
        remove_mc_fallbacks(clone)
        replace_in_runs(clone, mapping)
        set_textbox_autofit(clone)
        # 빈 여백 스페이서 문단은 버린다 → 예약된 빈 공간/빈 페이지 제거.
        # (박스·푸터·섹션은 위 판정에서 보존되므로 좌우 구성·색·푸터는 그대로다.)
        if _is_empty_filler(clone):
            continue
        body.append(clone)


def build(template: Path, output: Path, passages: list) -> None:
    for i, p in enumerate(passages):
        validate_passage(i, p)

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
        content = children
        if children and children[-1].tag == SECTPR:
            final_sectpr = deepcopy(children[-1])
            content = children[:-1]

        left_unit = content[:LEFT_END]    # 좌페이지 블록
        right_unit = content[LEFT_END:]   # 우페이지 블록

        for c in list(body):
            body.remove(c)

        state = {"num": 900000, "hex": 0x70000000}

        # 지문을 2개씩 짝지어 처리 (좌=A, 우=B)
        pairs = [passages[i:i + 2] for i in range(0, len(passages), 2)]
        for pi, pair in enumerate(pairs):
            is_last = (pi == len(pairs) - 1)
            a = pair[0]
            b = pair[1] if len(pair) > 1 else None

            _fill_blocks(left_unit, build_mapping(a), body, state)
            # 홀수 지문이면 마지막 쌍의 우페이지 유닛을 '비워서' 넣지 말고 아예 만들지 않는다.
            # (우 유닛 안에 있는 nextPage 섹션 나눔·우페이지 배경/캐리커쳐/푸터까지 통째로 빠지므로
            #  트레일링 빈 우페이지가 생기지 않는다. 좌페이지는 자체 배경·푸터로 정상 완결.)
            if b is not None:
                _fill_blocks(right_unit, build_mapping(b), body, state)

            # 세트 사이 페이지(섹션) 나눔: 마지막 세트 뒤엔 넣지 않음
            if not is_last and final_sectpr is not None:
                sep = etree.Element(P)
                ppr = etree.SubElement(sep, PPR)
                sect = deepcopy(final_sectpr)
                # 다음 세트가 새 페이지에서 시작하도록 type 제거(기본 nextPage)
                for ty in sect.findall(w("type")):
                    sect.remove(ty)
                ppr.append(sect)
                body.append(sep)

        if final_sectpr is not None:
            body.append(final_sectpr)

        tree.write(str(doc_xml), encoding="UTF-8", xml_declaration=True, standalone=True)

        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
            for fp in sorted(tmp_dir.rglob("*")):
                if fp.is_file():
                    zf.write(fp, fp.relative_to(tmp_dir).as_posix())


# ───────────────────────────────────────────────────────────────────
#  생성 후 구조 검증
# ───────────────────────────────────────────────────────────────────
def _check_docpr_unique(xml: str) -> list[str]:
    """wp:docPr/@id 중복 검사 (Word 호환: 도형 ID 는 문서 내 유일해야 함)."""
    import collections
    import re
    ids = re.findall(r'<wp:docPr[^>]*\bid="(\d+)"', xml)
    dup = [k for k, v in collections.Counter(ids).items() if v > 1]
    return [f"도형 docPr id 중복: {dup}"] if dup else []


def verify_output(output: Path, n_passages: int) -> list[str]:
    issues = []
    xml = zipfile.ZipFile(output).read("word/document.xml").decode("utf-8")
    leftover = _LEFTOVER_TOKEN_RE.findall(xml)
    if leftover:
        issues.append(f"미치환 템플릿 토큰 남음: {sorted(set(leftover))}")
    for bad in ("[??]", "[????]"):
        if bad in xml:
            issues.append(f"깨진 라벨 '{bad}' 남음 ({xml.count(bad)}개)")
    # 도형/이미지 docPr id 중복 = Word "복구 필요" 에러의 원인 → 자동 차단
    issues += _check_docpr_unique(xml)
    # 라벨이 채워진 횟수 = 지문 수와 일치해야 함 (좌/우 각 1회)
    for label in ("[영독]", "[요지]", "[전개방식]", "[어휘]"):
        got = xml.count(label)
        if got < n_passages:
            issues.append(f"'{label}' {got}개 (기대 {n_passages})")
    try:
        import docx  # noqa
        docx.Document(str(output))
    except Exception as exc:  # pragma: no cover
        issues.append(f"python-docx 열기 실패: {exc}")
    return issues


def _no_dup_pairs(pairs: list) -> dict:
    """JSON 객체에 같은 키가 두 번 나오면 즉시 에러(옛 데이터 잔류/병합 깨짐 차단)."""
    seen: dict = {}
    for k, v in pairs:
        if k in seen:
            raise ValueError(
                f"한 객체 안에 키 '{k}' 가 중복됨 — 옛 데이터와 새 데이터가 병합됐을 "
                f"가능성이 큽니다. JSON 을 부분 수정하지 말고 '통째로 새로' 쓰세요."
            )
        seen[k] = v
    return seen


def parse_args(argv) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="EBS 지문 교재 빌드")
    ap.add_argument("--in", dest="in_json", default=str(INPUT_JSON),
                    help="입력 JSON 경로(병렬 작업 시 분리된 파일명 권장)")
    ap.add_argument("--out", dest="out", default=str(OUTPUT), help="docx 출력 경로")
    return ap.parse_args(argv)


def main(argv=None) -> int:
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    args = parse_args(argv)
    in_json, out = Path(args.in_json), Path(args.out)
    try:
        passages = json.loads(in_json.read_text(encoding="utf-8"),
                              object_pairs_hook=_no_dup_pairs)
    except ValueError as exc:
        print(f"⛔ 입력 JSON 오류 — 빌드 중단: {exc}", file=sys.stderr)
        return 4
    if not isinstance(passages, list) or not passages:
        print("입력 JSON 이 비어 있음", file=sys.stderr)
        return 1
    build(TEMPLATE, out, passages)
    print(f"생성 완료: {out}  (지문 {len(passages)}개)")

    # ── 제1원칙: 원문 무결성 검사 (치명적) ──
    # RAW(붙여넣은 원문)의 모든 문장이 docx 에 들어갔는지 대조. 빠지면 폐기 후 중단.
    missing = integrity.check_records(out, passages, raw_key="RAW", id_key="LESSON_NUMBER")
    if missing:
        try:
            out.unlink()
        except OSError:
            pass
        print("⛔ 원문 누락 감지 — 교재 생성 중단 (파일 삭제됨):", file=sys.stderr)
        for it in missing:
            print("  -", it, file=sys.stderr)
        return 3

    issues = verify_output(out, len(passages))
    if issues:
        print("⚠️ 검증 경고:")
        for it in issues:
            print("  -", it)
        return 2
    print("✅ 구조 검증 통과 (placeholder 0 / 깨진 라벨 0 / 라벨·지문수 일치)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
