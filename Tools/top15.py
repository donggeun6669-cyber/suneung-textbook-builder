# -*- coding: utf-8 -*-
"""
top15.py — 수능 TOP15 결정적 추출기 (연도 매칭 오류 '원천 차단').

조교/AI 는 **학년도 + 월** 만 주면, CSV(Documents/정답률별_정리_평가원.csv)에서
그 시험의 TOP15(오답률순 상위 15)를 **CSV 값 그대로** 뽑아 JSON 으로 출력한다.
→ 정답·정답률·선지비율을 사람이 옮겨 적다 틀리는 사고를 없애고(=수치 지어내기 차단),
→ 학년도↔시행연도 혼동으로 엉뚱한 연도를 잡는 사고를 없앤다.

왜 off-by-one(2023학년도를 2022 로 잡는 류) 이 구조적으로 불가능한가:
  CSV '시험' 컬럼은 '학년도.월(시행연도시행)' 형식이다(점 앞 4자리 = 학년도).
  이 스크립트는 **'학년도.월' 접두로만** 행을 찾는다. 시행연도(괄호 안)는 매칭에
  쓰지 않으므로, '시행연도 ±1 산수' 자체가 필요 없고 → 틀릴 여지가 없다.

사용법:
    python Tools/top15.py 2023 9                # JSON 출력(기본)
    python Tools/top15.py 2023 9 --table        # 사람이 보기 좋은 표
    python Tools/top15.py 2023 11 --n 15         # 상위 N개(기본 15)
    python Tools/top15.py 2023 9 --csv <경로>    # CSV 경로 직접 지정

주의: 입력은 **학년도**다(시행연도 아님). 헷갈리면 "학년도 = 시행연도 + 1".
      예) 2022년 9월에 시행된 시험 → 학년도는 2023 → `top15.py 2023 9`.
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CSV = ROOT / "Documents" / "정답률별_정리_평가원.csv"

# 시험 문자열 'YYYY.MM(YY시행)' 에서 괄호 안 시행연도 2자리를 뽑는다.
_SIHAENG_RE = re.compile(r"\((\d{2})시행\)")


def decode_sihaeng(exam: str) -> int | None:
    """'2023.09(22시행)' → 2022 (시행연도). 못 읽으면 None."""
    m = _SIHAENG_RE.search(exam or "")
    return 2000 + int(m.group(1)) if m else None


def load_rows(csv_path: Path) -> list:
    """CSV(BOM 포함 가능)를 dict 행 리스트로 읽는다."""
    text = csv_path.read_text(encoding="utf-8-sig")
    return list(csv.DictReader(text.splitlines()))


def resolve_exam(rows: list, haknyeon: str, month: int) -> list:
    """'학년도.월' 접두로 시작하는 시험 행만 추린다(시행연도 무시 → off-by-one 불가)."""
    prefix = f"{haknyeon}.{month:02d}"
    return [r for r in rows if (r.get("시험") or "").startswith(prefix)]


def top_records(matched: list, n: int) -> list:
    """오답률순 정렬 후 상위 N개를, 빌드 JSON 에 그대로 쓸 필드로 변환."""
    def rank_key(r: dict) -> int:
        raw = (r.get("오답률순") or "").strip()
        return int(raw) if raw.isdigit() else 999

    matched = sorted(matched, key=rank_key)[:n]
    out = []
    for r in matched:
        ans = (r.get("정답") or "").strip()
        out.append({
            "RANK": (r.get("오답률순") or "").strip(),
            "NUMBER": (r.get("번호") or "").strip(),
            "CORRECT_ANSWER": int(ans) if ans.isdigit() else ans,
            "CORRECT_RATE": (r.get("정답률") or "").strip(),
            "PERCENTAGES": [(r.get(f"선지{i}") or "").strip() for i in range(1, 6)],
            "YEAR": (r.get("시험") or "").strip(),
        })
    return out


def long_passage_hints(recs: list) -> list:
    """장문(한 지문 공유) 묶음 안내 — 41·42 / 43·44·45 가 TOP 안에 들면 알린다."""
    nums = {r["NUMBER"] for r in recs}
    hints = []
    if {"41", "42"} <= nums:
        hints.append("41·42 는 같은 장문 지문 — 한 레코드로 합치되 두 문항의 정답/비율/근거를 모두 보존")
    if len({"43", "44", "45"} & nums) >= 2:
        hints.append("43·44·45 는 같은 장문 지문 — 한 레코드로 합치되 각 문항 정답/근거를 모두 보존")
    return hints


def main(argv=None) -> int:
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    ap = argparse.ArgumentParser(description="수능 TOP15 결정적 추출(학년도+월 → CSV 값 그대로)")
    ap.add_argument("haknyeon", help="학년도 4자리 (예: 2023). 시행연도 아님! (학년도=시행+1)")
    ap.add_argument("month", type=int, help="월 (6 / 9 / 11)")
    ap.add_argument("--sihaeng", type=int, default=None,
                    help="시행연도(조교가 함께 말한 값). 주면 학년도와 모순되는지 교차검증 "
                         "(학년도=시행+1). 오늘 같은 1년 어긋남을 여기서 잡는다.")
    ap.add_argument("--csv", default=str(DEFAULT_CSV), help="CSV 경로")
    ap.add_argument("--n", type=int, default=15, help="상위 N개 (기본 15)")
    ap.add_argument("--table", action="store_true", help="표 형식으로 보기")
    args = ap.parse_args(argv)

    csv_path = Path(args.csv)
    if not csv_path.exists():
        print(f"⛔ CSV 없음: {csv_path}", file=sys.stderr)
        return 2
    rows = load_rows(csv_path)
    matched = resolve_exam(rows, args.haknyeon, args.month)
    if not matched:
        exams = sorted({(r.get("시험") or "").strip() for r in rows if (r.get("시험") or "").strip()})
        print(f"⛔ '{args.haknyeon}.{args.month:02d}' 로 시작하는 시험이 CSV 에 없음.", file=sys.stderr)
        print("   혹시 시행연도를 학년도 자리에 넣지 않았나요? (학년도 = 시행연도 + 1)", file=sys.stderr)
        print(f"   CSV 에 있는 시험들: {', '.join(exams)}", file=sys.stderr)
        return 3

    recs = top_records(matched, args.n)
    exam = recs[0]["YEAR"]
    hints = long_passage_hints(recs)

    # ── 핵심 가드: 조교가 말한 '시행연도'와 '학년도'가 모순되면 멈춘다 ──
    #   조교는 보통 둘 다 말한다(예: "2022시행 (2023학년도)"). 이 redundancy 를
    #   교차검증에 쓴다. 오늘 사고('2022시행'을 학년도 2022 로 잡음)는 여기서 잡혔을 것.
    actual_sihaeng = decode_sihaeng(exam)
    if args.sihaeng is not None and actual_sihaeng is not None and args.sihaeng != actual_sihaeng:
        print(f"⛔ 연도 모순 — 멈춤.", file=sys.stderr)
        print(f"   학년도 {args.haknyeon} 의 {args.month}월 시험은 CSV 상 '{exam}' "
              f"= {actual_sihaeng}년 시행입니다.", file=sys.stderr)
        print(f"   그런데 조교가 말한 시행연도는 {args.sihaeng} 입니다.", file=sys.stderr)
        print(f"   → 시행 {args.sihaeng} 이 맞다면 학년도는 {args.sihaeng + 1} 입니다. "
              f"`top15.py {args.sihaeng + 1} {args.month}` 로 다시 실행하세요.", file=sys.stderr)
        return 4

    confirm = f"{exam}  =  학년도 {args.haknyeon} / {actual_sihaeng}년 시행 / {args.month}월"
    if args.table:
        print(f"시험 확인: {confirm}   TOP{len(recs)}")
        print("순위  번호  정답  정답률   선지1~5")
        for r in recs:
            print(f"{r['RANK']:>3}  {r['NUMBER']:>4}   {r['CORRECT_ANSWER']}   {r['CORRECT_RATE']:>4}   {'  '.join(r['PERCENTAGES'])}")
        for h in hints:
            print("힌트:", h)
    else:
        print(json.dumps(
            {"exam": exam, "exam_decoded": confirm, "haknyeon": args.haknyeon,
             "sihaeng": actual_sihaeng, "month": args.month,
             "top15": recs, "hints": hints},
            ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
