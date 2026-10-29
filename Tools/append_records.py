# -*- coding: utf-8 -*-
"""
append_records.py — 레코드를 '5개씩 안전하게' 누적하는 도구.

왜 필요한가:
  15문항을 한 번에 출력하면 (1) 길이 초과로 출력이 중간에 잘리고,
  (2) 기존 파일의 옛 데이터와 섞여 한 객체에 키가 두 벌씩 박힌다(병합 깨짐).
  이 도구를 쓰면 모델은 매번 '작은 청크(5개)'만 새로 출력하고, 이어 붙이는 일은
  코드가 한다 → '잘림'과 '병합 사고'가 구조적으로 불가능해진다.

흐름(조교/AI):
  1) 새로 시작 — 타깃을 빈 배열로 초기화(옛 데이터 오염원 제거):
       python Tools/append_records.py Tools/suneung_2023_09.json --reset
  2) 5개짜리 청크 파일을 만들어 이어 붙이기(원하는 만큼 반복):
       python Tools/append_records.py Tools/suneung_2023_09.json chunk.json
  3) 현재 상태 확인(청크 인자 없이):
       python Tools/append_records.py Tools/suneung_2023_09.json

안전장치:
  - 청크/타깃의 각 객체에 '키 중복'이 있으면 거부(병합 깨짐 차단).
  - 같은 NUMBER 가 이미 있으면 거부(같은 문항 두 번 차단).
  - 이어 붙인 뒤 항상 통째로 다시 써서 JSON 파싱을 보장.
  - 실패 시 타깃을 건드리지 않는다(원자적: 검증 통과 후에만 기록).
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def _no_dup_pairs(pairs: list) -> dict:
    seen: dict = {}
    for k, v in pairs:
        if k in seen:
            raise ValueError(
                f"한 객체에 키 '{k}' 가 중복 — 청크가 깨졌거나 옛 데이터와 섞였습니다. "
                f"청크는 '새 레코드만' 담아야 합니다."
            )
        seen[k] = v
    return seen


def _load_array(path: Path) -> list:
    if not path.exists():
        return []
    txt = path.read_text(encoding="utf-8").strip()
    if not txt:
        return []
    data = json.loads(txt, object_pairs_hook=_no_dup_pairs)
    if not isinstance(data, list):
        raise ValueError(f"{path} 가 JSON 배열이 아님")
    return data


def _load_chunk(src: str) -> list:
    raw = sys.stdin.read() if src == "-" else Path(src).read_text(encoding="utf-8")
    data = json.loads(raw, object_pairs_hook=_no_dup_pairs)
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list) or not data:
        raise ValueError("청크가 비었거나 JSON 배열이 아님")
    return data


def main(argv=None) -> int:
    for _s in (sys.stdout, sys.stderr):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    ap = argparse.ArgumentParser(description="레코드를 5개씩 안전하게 누적(잘림·병합 사고 방지)")
    ap.add_argument("target", help="누적 대상 JSON 파일(시험별, 예: Tools/suneung_2023_09.json)")
    ap.add_argument("chunk", nargs="?", help="이어 붙일 청크 JSON(배열) 파일. '-' 면 표준입력")
    ap.add_argument("--reset", action="store_true", help="타깃을 빈 배열 []로 초기화")
    ap.add_argument("--id", default="NUMBER", help="중복 판정 키(기본 NUMBER)")
    args = ap.parse_args(argv)
    target = Path(args.target)

    if args.reset:
        target.write_text("[]\n", encoding="utf-8")
        print(f"초기화: {target} = []")
        if not args.chunk:
            return 0

    try:
        existing = _load_array(target)
    except ValueError as exc:
        print(f"⛔ 타깃 읽기 실패: {exc}", file=sys.stderr)
        return 4

    if not args.chunk:
        ids = [r.get(args.id) for r in existing]
        print(f"현재 {target}: {len(existing)}개 — {ids}")
        return 0

    try:
        chunk = _load_chunk(args.chunk)
    except (ValueError, OSError) as exc:
        print(f"⛔ 청크 읽기/파싱 실패: {exc}", file=sys.stderr)
        return 4

    have = {str(r.get(args.id)) for r in existing}
    merged = list(existing)
    added = []
    for r in chunk:
        rid = str(r.get(args.id))
        if rid in have:
            print(f"⛔ {args.id}={rid} 이(가) 이미 있음 — 중복 추가 거부(청크에서 빼세요).",
                  file=sys.stderr)
            return 4
        have.add(rid)
        merged.append(r)
        added.append(rid)

    target.write_text(json.dumps(merged, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"✅ {len(added)}개 추가: {added}  → 총 {len(merged)}개 ({target})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
