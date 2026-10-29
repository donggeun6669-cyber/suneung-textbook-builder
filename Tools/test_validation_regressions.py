# -*- coding: utf-8 -*-
"""
회귀 테스트 — Codex 감사에서 나온 버그 5건의 최소 재현 + 수정 확인.

실행:  python Tools/test_validation_regressions.py
모든 케이스 통과 시 종료코드 0, 하나라도 실패하면 1.
표준 라이브러리만 사용(외부 의존 없음). 빌드 산출물은 만들지 않고
순수 함수(validate_*, apply_tokens_once, verify 보조)만 단위로 친다.
"""

from __future__ import annotations

import copy
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

# 윈도우 한글 콘솔(cp949)에서 ✅/❌ 출력이 깨지지 않도록
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import build_suneung as bs
import build_ebs as be

ROOT = Path(__file__).resolve().parents[1]

_PASS = 0
_FAIL = 0


def check(name: str, cond: bool, detail: str = "") -> None:
    global _PASS, _FAIL
    if cond:
        _PASS += 1
        print(f"  ✅ {name}")
    else:
        _FAIL += 1
        print(f"  ❌ {name}  {detail}")


def expect_error(name: str, fn, needle: str = "") -> None:
    """fn() 이 ValueError 를 던지면 통과. needle 이 주어지면 메시지에 포함되어야."""
    try:
        fn()
    except ValueError as exc:
        msg = str(exc)
        check(name, (needle in msg) if needle else True,
              f"메시지에 '{needle}' 없음 → {msg!r}")
        return
    except Exception as exc:  # 잘못된 예외 종류
        check(name, False, f"ValueError 가 아닌 {type(exc).__name__}: {exc}")
        return
    check(name, False, "에러가 나야 하는데 통과해 버림")


# ── 베이스 record/passage (정상값) 를 'SAMPLE' 파일에서 빌려온다 ──
#   실제 작업 파일(suneung_input.json/ebs_input.json)은 빈 배열 []로 두어 옛 데이터
#   오염을 막으므로(병합 사고 방지), 테스트의 기준 레코드는 *.SAMPLE.json 에서 읽는다.
SUN = json.loads((ROOT / "Tools" / "suneung_input.SAMPLE.json").read_text(encoding="utf-8"))[0]
EBS = json.loads((ROOT / "Tools" / "ebs_input.SAMPLE.json").read_text(encoding="utf-8"))[0]


def sun(**over) -> dict:
    r = copy.deepcopy(SUN)
    r.update(over)
    return r


def ebs(**over) -> dict:
    r = copy.deepcopy(EBS)
    r.update(over)
    return r


# ═══════════════════════════════════════════════════════════════════
print("[#1-suneung] 자료형 검증 (str/list 자리 혼동 차단)")
# 정상값은 통과해야 한다
try:
    bs.validate_record(0, sun())
    check("정상 record 통과", True)
except Exception as exc:
    check("정상 record 통과", False, str(exc))

# PASSAGE 자리에 리스트 → 에러
expect_error("PASSAGE 가 list 면 에러", lambda: bs.validate_record(0, sun(PASSAGE=["a", "b"])),
             "문자열")
# CHOICES 자리에 문자열 → 에러 (조용히 ['...'] 출력 금지)
expect_error("CHOICES 가 str 이면 에러",
             lambda: bs.validate_record(0, sun(CHOICES="①②③④⑤")), "리스트")
# VOCABULARY 항목 중 하나가 숫자 → 에러
bad_vocab = list(SUN["VOCABULARY"])
bad_vocab[3] = 123
expect_error("VOCABULARY 원소가 비문자열이면 에러",
             lambda: bs.validate_record(0, sun(VOCABULARY=bad_vocab)), "문자열")
# CORRECT_ANSWER 가 문자열 "3" → int 변환되어 통과(자료형 너그럽게), "abc" → 에러
try:
    bs.validate_record(0, sun(CORRECT_ANSWER="3"))
    check("CORRECT_ANSWER='3' 통과(정수 변환)", True)
except Exception as exc:
    check("CORRECT_ANSWER='3' 통과(정수 변환)", False, str(exc))
expect_error("CORRECT_ANSWER='abc' 에러",
             lambda: bs.validate_record(0, sun(CORRECT_ANSWER="abc")), "1~5")
# SENTENCE_TRIPLES 항목이 문자열 → 에러
expect_error("SENTENCE_TRIPLES 항목이 list 가 아니면 에러",
             lambda: bs.validate_record(0, sun(SENTENCE_TRIPLES=["just a string"])), "리스트")

print("[#1-ebs] 자료형 검증")
try:
    be.validate_passage(0, ebs())
    check("정상 passage 통과", True)
except Exception as exc:
    check("정상 passage 통과", False, str(exc))
expect_error("LESSON_TITLE 가 list 면 에러",
             lambda: be.validate_passage(0, ebs(LESSON_TITLE=["x"])), "문자열")
expect_error("NUMBERED_SENTENCES 가 str 이면 에러",
             lambda: be.validate_passage(0, ebs(NUMBERED_SENTENCES="① a ② b")), "리스트")
bad_tr = list(EBS["KOREAN_TRANSLATIONS"])
bad_tr[0] = {"x": 1}
expect_error("KOREAN_TRANSLATIONS 원소가 dict 면 에러",
             lambda: be.validate_passage(0, ebs(KOREAN_TRANSLATIONS=bad_tr)), "문자열")


# ═══════════════════════════════════════════════════════════════════
print("[#2-suneung] 어휘 분배 — 슬롯보다 단어가 많아도 소실 금지")


def simulate_vocab_distribution(words: list, n_slots: int) -> list:
    """build() 내부 어휘 분배 로직을 그대로 떼어내 시뮬레이션.
    반환: 각 슬롯에 들어간 단어 리스트(빈 슬롯 제외하지 않음)."""
    if len(words) <= n_slots:
        vocab_slots = [[wd] for wd in words]
    else:
        vocab_slots = [[wd] for wd in words[:n_slots - 1]]
        vocab_slots.append(list(words[n_slots - 1:]))
    return vocab_slots


# 25개 단어, 슬롯 20개 → 마지막 슬롯에 나머지 6개가 합쳐져 총 25개 전부 보존
words25 = [f"w{i}" for i in range(25)]
slots = simulate_vocab_distribution(words25, 20)
flat = [w for s in slots for w in s]
check("25단어/20슬롯: 단어 소실 0", flat == words25,
      f"보존 {len(flat)}개 (기대 25)")
check("25단어/20슬롯: 슬롯 수 ≤ 20", len(slots) <= 20, f"슬롯 {len(slots)}개")
check("25단어/20슬롯: 마지막 슬롯에 6개 합침", len(slots[-1]) == 6, f"{len(slots[-1])}개")
# 정확히 맞는 경우
check("20단어/20슬롯: 각 1개", simulate_vocab_distribution([f"w{i}" for i in range(20)], 20)
      == [[f"w{i}"] for i in range(20)])
# 적은 경우
check("15단어/20슬롯: 소실 0",
      [w for s in simulate_vocab_distribution([f"w{i}" for i in range(15)], 20) for w in s]
      == [f"w{i}" for i in range(15)])

# 실제 빌드로 끝까지: 25개 어휘를 넣고 docx 안에 25개가 다 있는지
import zipfile


def build_count_vocab(words: list) -> int:
    import tempfile
    rec = sun(VOCABULARY=words)
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "t.docx"
        bs.build(bs.TEMPLATE, out, [rec])
        xml = zipfile.ZipFile(out).read("word/document.xml").decode("utf-8")
    return sum(xml.count(wmark) for wmark in [w.split()[0] for w in words])


# 각 단어의 첫 토큰을 고유하게 만들어 docx 에서 카운트
uniq_words = [f"UNIQVOCAB{i}  뜻{i}  (유) syn{i}  (반) ant{i}" for i in range(25)]
present = 0
import tempfile
rec25 = sun(VOCABULARY=uniq_words)
with tempfile.TemporaryDirectory() as td:
    out = Path(td) / "t.docx"
    bs.build(bs.TEMPLATE, out, [rec25])
    xml = zipfile.ZipFile(out).read("word/document.xml").decode("utf-8")
present = sum(1 for i in range(25) if f"UNIQVOCAB{i}" in xml)
check("실제 빌드: 25개 어휘 전부 docx 에 존재(소실 0)", present == 25,
      f"존재 {present}/25")


# ═══════════════════════════════════════════════════════════════════
print("[#3] 1회성 토큰 치환 — 사용자 텍스트 속 {{...}} 재치환 안 됨")
# PASSAGE 값 안에 우연히 '{{KOREAN_TITLE}}' 가 들어 있어도, KOREAN_TITLE 치환이
# 그 안으로 새어 들어가면 안 된다.
mapping = [
    ("{{PASSAGE}}", "지문에 {{KOREAN_TITLE}} 가 들어있다"),
    ("{{KOREAN_TITLE}}", "제목XYZ"),
]
res = bs.apply_tokens_once("본문: {{PASSAGE}} 끝", mapping)
check("PASSAGE 안의 {{KOREAN_TITLE}} 는 그대로 유지",
      res == "본문: 지문에 {{KOREAN_TITLE}} 가 들어있다 끝", f"→ {res!r}")
# 별도로 등장한 진짜 {{KOREAN_TITLE}} 는 치환된다
res2 = bs.apply_tokens_once("{{PASSAGE}} / {{KOREAN_TITLE}}", mapping)
check("진짜 토큰은 정상 치환",
      res2 == "지문에 {{KOREAN_TITLE}} 가 들어있다 / 제목XYZ", f"→ {res2!r}")
# 긴 키 우선(mapping 순서) 확인 — EBS 어휘 라벨
m_ebs = [("[??]\n{{VOCABULARY}}", "[어휘]\nW"), ("{{VOCABULARY}}", "W")]
check("긴 키 우선 매칭(EBS)",
      be.apply_tokens_once("[??]\n{{VOCABULARY}}", m_ebs) == "[어휘]\nW")


# ═══════════════════════════════════════════════════════════════════
print("[#4] verify_output 토큰 잔여 검사 — 사용자 {{ x }} 오판 안 함")
# 정상적인 사용자 텍스트의 {{ x }} 는 미치환 토큰으로 잡히면 안 된다
check("'{{ x }}' 는 템플릿 토큰으로 안 잡힘",
      bs._LEFTOVER_TOKEN_RE.findall("수식: {{ x }} 와 {{ y+1 }}") == [])
check("'{{소문자}}' 도 안 잡힘", bs._LEFTOVER_TOKEN_RE.findall("{{passage}}") == [])
# 진짜 미치환 템플릿 토큰은 잡는다
check("'{{PASSAGE}}' 미치환은 잡힘",
      bs._LEFTOVER_TOKEN_RE.findall("oops {{PASSAGE}} left") == ["PASSAGE"])
check("EBS '{{GIST}}' 미치환은 잡힘",
      be._LEFTOVER_TOKEN_RE.findall("{{GIST}}") == ["GIST"])
check("EBS '{{ 사용자 }}' 오판 안 함",
      be._LEFTOVER_TOKEN_RE.findall("공식 {{ a }}") == [])


# ═══════════════════════════════════════════════════════════════════
print("[#5] PERCENTAGES 형식 가벼운 점검(경고만)")
ok = bs.check_percentage_format([sun(PERCENTAGES=["17%", "30%", "23%", "10%", "20%"])])
check("정상 '17%' 들은 경고 없음", ok == [], f"→ {ok}")
warn = bs.check_percentage_format([sun(PERCENTAGES=["17", "30%", "23%", "10%", "20%"])])
check("'17'(% 빠짐)은 경고 1건", len(warn) == 1, f"→ {warn}")
check("3.5% 같은 소수도 허용",
      bs.check_percentage_format([sun(PERCENTAGES=["3.5%", "30%", "23%", "10%", "33.5%"])]) == [])
# 형식 오류는 '경고'일 뿐 validate_record(하드 에러)를 막지 않는다
try:
    bs.validate_record(0, sun(PERCENTAGES=["17", "30%", "23%", "10%", "20%"]))
    check("PERCENTAGES 형식 이상은 validate_record 를 막지 않음(경고 전용)", True)
except ValueError as exc:
    check("PERCENTAGES 형식 이상은 validate_record 를 막지 않음(경고 전용)", False, str(exc))


# ═══════════════════════════════════════════════════════════════════
print("[#6] 연도 매칭/병합 깨짐 가드 (오늘 난 사고 재발 방지)")
import csv as _csvmod
import tempfile
import top15


def _write_json(text: str) -> Path:
    f = tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8")
    f.write(text)
    f.close()
    return Path(f.name)

# (a) 한 객체에 키 중복 → load_records_strict 가 거부 (병합 깨짐)
expect_error("중복키 JSON 은 거부됨",
             lambda: bs.load_records_strict(
                 _write_json('[{"NUMBER":"38","NUMBER":"34"}]')),
             "중복")
# (b) 한 파일에 연도 두 개 섞임 → check_year_consistency 가 잡음
mixed = [sun(YEAR="2023.09(22시행)"), sun(YEAR="2026.11(25시행)")]
check("연도 섞이면 잡힘", len(bs.check_year_consistency(mixed)) >= 1)
# (c) YEAR 형식 오류 → 잡힘
check("YEAR 형식 오류 잡힘", len(bs.check_year_consistency([sun(YEAR="2023년 9월")])) >= 1)
# (d) 정상 단일 연도 → 통과
check("정상 단일 연도 통과", bs.check_year_consistency([sun(YEAR="2026.11(25시행)")]) == [])

# (e) CSV 대조: 틀린 정답이면 fatal
_csv_path = ROOT / "Documents" / "정답률별_정리_평가원.csv"
if _csv_path.exists():
    rows = list(_csvmod.DictReader(_csv_path.read_text(encoding="utf-8-sig").splitlines()))
    real = [r for r in rows if (r.get("시험") or "").startswith("2023.09")]
    one = sorted(real, key=lambda r: int(r["오답률순"]))[0]  # 1위 = 38번 정답2
    good = sun(NUMBER=one["번호"], YEAR=one["시험"], CORRECT_ANSWER=int(one["정답"]),
               CORRECT_RATE=one["정답률"],
               PERCENTAGES=[one[f"선지{i}"] for i in range(1, 6)])
    f0, _ = bs.crosscheck_csv([good], _csv_path)
    check("CSV 와 일치하면 통과(fatal 0)", f0 == [], f"→ {f0}")
    bad = copy.deepcopy(good)
    bad["CORRECT_ANSWER"] = 9 if int(one["정답"]) != 9 else 1
    fb, _ = bs.crosscheck_csv([bad], _csv_path)
    check("CSV 와 정답 다르면 fatal", len(fb) >= 1)
    # (f) top15 헬퍼: 학년도↔시행 모순이면 멈춤(exit 4)
    rc_bad = top15.main(["2022", "9", "--sihaeng", "2022"])   # 학년도2022는 21시행
    check("top15 모순(학년도2022/시행2022) 거부", rc_bad == 4, f"→ rc={rc_bad}")
    rc_ok = top15.main(["2023", "9", "--sihaeng", "2022"])    # 학년도2023 = 22시행 ✓
    check("top15 정상(학년도2023/시행2022) 통과", rc_ok == 0, f"→ rc={rc_ok}")
else:
    print("  · CSV 없어 CSV 대조 테스트 생략")

# (g) EBS 도 중복키 거부
expect_error("EBS 중복키 JSON 거부",
             lambda: json.loads('[{"RAW":"a","RAW":"b"}]', object_pairs_hook=be._no_dup_pairs),
             "중복")


# ═══════════════════════════════════════════════════════════════════
print("[#7] 원문 무결성 — 마커/라벨 가짜 누락경보 차단 + 진짜 누락은 잡음")
import integrity as ig

# docx 렌더본: 어법/어휘 지문은 '인라인 원문자가 제거되어' 박힌다(build 자동 밑줄).
_docx = ("In everyday life we use this term. "
         "Learn how to choose from alternatives and make a decision requires knowledge. "
         "According to the passage individuality is emphasized while conformity is reduced.")
_hay = ig._coverage_clean(_docx)

# (a) ★핵심 버그: RAW 에 인라인 원문자(②)가 있어도 docx(마커 없음)와 매칭 → 가짜경보 0
raw_inline = "② Learn how to choose from alternatives and make a decision requires knowledge."
check("어법 인라인 원문자: 가짜 누락경보 없음 (옛 버그)", ig.find_missing(raw_inline, _hay) == [],
      f"→ {ig.find_missing(raw_inline, _hay)}")

# (b) 문삽 선지줄 '① ② ③ ④ ⑤' 는 '문장'이 아님 → 건너뜀(경보 없음)
check("문삽 '①②③④⑤' 선지줄 건너뜀", ig.find_missing("① ② ③ ④ ⑤", _hay) == [],
      f"→ {ig.find_missing('① ② ③ ④ ⑤', _hay)}")

# (c) 요약 헤더 '(A) ……(B)' 도 '문장'이 아님 → 건너뜀
check("요약 '(A) ……(B)' 헤더 건너뜀", ig.find_missing("(A) ……(B)", _hay) == [],
      f"→ {ig.find_missing('(A) ……(B)', _hay)}")

# (d) ★안전장치 보존: 진짜 '문장'이 통째로 빠지면 여전히 잡힌다
dropped = "This entire sentence was dropped during analysis."
check("진짜 문장 누락은 잡힘(안전장치 보존)", ig.find_missing(dropped, _hay) == [dropped],
      f"→ {ig.find_missing(dropped, _hay)}")

# (e) 원문자 위치가 문장 중간에 있어도 문장이 같으면 통과
mid = "Learn how to choose ② from alternatives and make a decision requires knowledge."
check("원문자 위치 달라도 문장 같으면 통과", ig.find_missing(mid, _hay) == [],
      f"→ {ig.find_missing(mid, _hay)}")

# (g) ★문삽 사각지대: 주어진 문장(GIVEN)이 분석(SENTENCE_TRIPLES)에서 빠지면 잡힌다
giv_ok = {"GIVEN": "This inserted sentence belongs in the analysis.",
          "SENTENCE_TRIPLES": [["This inserted sentence belongs in the analysis.", "구문", "해석"],
                               ["Another passage sentence here.", "구문", "해석"]], "NUMBER": "38"}
giv_bad = {"GIVEN": "This inserted sentence belongs in the analysis.",
           "SENTENCE_TRIPLES": [["Another passage sentence here.", "구문", "해석"]], "NUMBER": "39"}
check("GIVEN 이 분석에 있으면 통과", ig.check_given_in_triples([giv_ok]) == [],
      f"→ {ig.check_given_in_triples([giv_ok])}")
check("GIVEN 이 분석에서 빠지면 잡힘(문삽 사각지대)",
      any("삽입문장 해설 누락" in m for m in ig.check_given_in_triples([giv_bad])),
      f"→ {ig.check_given_in_triples([giv_bad])}")

# (f) 검사기가 멈출 때 '검사기·RAW 고치지 말라'는 재발방지 안내가 붙는다
import tempfile as _tf
_bad_rec = {"RAW": "This sentence is absolutely not in the document at all.", "NUMBER": "99"}
with _tf.TemporaryDirectory() as _td:
    _o = Path(_td) / "x.docx"
    bs.build(bs.TEMPLATE, _o, [sun()])      # docx 엔 sun() 내용만 → _bad_rec 문장은 없음
    _iss = ig.check_records(_o, [_bad_rec])
check("누락 시 '검사기/RAW 고치지 말라' 안내 포함",
      any("고치지 말" in m for m in _iss), f"→ {_iss[-1] if _iss else '(없음)'}")


# ═══════════════════════════════════════════════════════════════════
print("[#8] 서식 구조 가드 — '모양 틀어지는' 사고 차단(어법/어휘·요약·순서/문삽)")
_gv_pa = ("People believe actions can ① be evaluated as right. To ② Learn how to choose "
          "is key. They ③ have discussions often, in gray areas ④ that create dilemmas, "
          "and they decide ⑤ whether to ask for help.")
_gv_ch = ["① be evaluated", "② Learn", "③ have discussions", "④ that create", "⑤ whether to ask"]


def _fmt(**over):
    r = sun(**over)
    r.pop("UNDERLINE", None)
    r.pop("SUB_QUESTIONS", None)
    return r

ok_gram = _fmt(QUESTION_TYPE="어법", CHOICES=_gv_ch, PASSAGE=_gv_pa)
check("어법 정상(선지 지문O·원문자5) 통과", bs.check_format_structure([ok_gram]) == [],
      f"→ {bs.check_format_structure([ok_gram])}")
_pa_nomark = _gv_pa
for _m in "①②③④⑤":
    _pa_nomark = _pa_nomark.replace(_m, "")
no_mark = _fmt(QUESTION_TYPE="어법", CHOICES=_gv_ch, PASSAGE=_pa_nomark)
check("어법 원문자 없음 → 잡힘", any("원문자" in m for m in bs.check_format_structure([no_mark])),
      f"→ {bs.check_format_structure([no_mark])}")
bad_ch = _fmt(QUESTION_TYPE="어법",
              CHOICES=["① totally different", "② another thing", "③ none here", "④ not in text", "⑤ absent word"],
              PASSAGE=_gv_pa)
check("어법 선지 지문불일치(빈칸식 오렌더) → 잡힘",
      any("빈칸식" in m for m in bs.check_format_structure([bad_ch])),
      f"→ {bs.check_format_structure([bad_ch])}")
check("요약 SUMMARY_SENTENCE 없음 → 잡힘",
      any("SUMMARY_SENTENCE" in m for m in bs.check_format_structure([sun(QUESTION_TYPE="요약", SUMMARY_SENTENCE="")])))
check("순서 GIVEN 없음 → 잡힘",
      any("GIVEN" in m for m in bs.check_format_structure([sun(QUESTION_TYPE="순서", GIVEN="")])))
# ★함축/밑줄의미(UNDERLINE 사용)는 '어휘'여도 어법/어휘 가드에서 제외 → 오탐 없음
hint = sun(QUESTION_TYPE="어휘", UNDERLINE="send us off into", CHOICES=["①", "②", "③", "④", "⑤"],
           PASSAGE="no marks here at all")
check("함축(어휘+UNDERLINE)은 오탐 없음", bs.check_format_structure([hint]) == [],
      f"→ {bs.check_format_structure([hint])}")


# ═══════════════════════════════════════════════════════════════════
print()
print(f"결과: {_PASS} 통과 / {_FAIL} 실패")
sys.exit(1 if _FAIL else 0)
