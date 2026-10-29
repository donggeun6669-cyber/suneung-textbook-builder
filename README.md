# 수능 영어 교재 자동 제작 시스템 (공개용 코드)

> 조교가 수능 · EBS 영어 문제 원문을 붙여넣으면, 학원 서식이 적용된 완성 교재(docx)가 나옵니다.

영어학원에 **유료 납품되어 실제 교재 제작 · 인쇄에 쓰이고 있는** 시스템의 코드 부분입니다.

- **효과**: 15문제 1세트 제작 시간 **약 75분 → 20분 이내 (약 73% 단축)**
- **기술**: Python · `python-docx` · `lxml` (Word 문서를 XML 단위로 다뤄 서식을 지킨 채 내용만 채움)
- **역할**: 문제 정의 · 요구사항 정리 · 검증 체계 설계 · 납품 · 운영까지 단독으로 맡았습니다

> ⚠️ 학원의 교재 서식 파일(docx 템플릿), 기출 지문, 정답률 데이터, 만든 교재는 **이 저장소에 넣지 않았습니다.**
> 그래서 이 저장소만으로는 교재가 만들어지지 않습니다. 설계와 코드를 보여드리기 위한 저장소입니다.
>
> 📐 **코드 안의 서식 수치는 예시 값입니다.** 여백 · 표 폭 · 글꼴 · 글자 크기 · 색은 실제 학원 교재와 **다르게** 바꿔 두었습니다.
> 학원 서식을 보호하기 위해서이며, 실제 값은 비공개 템플릿과 함께 관리합니다. 수치가 든 설명 주석도 같은 이유로 뺐습니다.

## 동작 흐름

```
조교가 문제 원문 붙여넣기
  → 정답률 CSV에서 오답률 TOP15 추출          (top15.py)
  → 문항별 분석 결과를 JSON으로 구조화         (append_records.py, 5문항씩 나눠 누적)
  → 템플릿 docx에 채워 넣기                    (build_suneung.py · build_ebs.py)
  → 무결성 · 수치 · 구조 검증을 통과할 때만 docx 산출  (integrity.py)
```

## 신뢰성 설계 (이 프로젝트의 핵심)

AI가 만든 결과물을 **반복해서 써도 깨지지 않게** 하는 데 가장 많은 시간을 썼습니다.

| 막는 사고 | 방법 | 코드 |
| --- | --- | --- |
| 원문 문장 · 선지 누락 | 붙여넣은 원문(`RAW`)을 문장 단위로 쪼개 완성된 docx 안에 전부 있는지 대조. 하나라도 없으면 docx를 지우고 빌드 중단 | `integrity.py` |
| 지어낸 정답률 · 수치 | 정답 · 정답률 · 선지별 비율은 평가원 CSV 값만 허용. 다르면 빌드 거부 | `build_suneung.py` `crosscheck_csv` |
| 학년도 · 시행연도 혼동 | CSV를 「학년도.월」 접두로만 찾아 off-by-one 오류가 구조적으로 생길 수 없게 함 | `top15.py` |
| 형식이 깨진 입력 | 선지 5개 · 어휘 수 · 필수 항목 · 중복 키를 스키마로 강제 | `build_suneung.py` `validate_record` |
| 고친 버그의 재발 | 외부 감사(Codex)에서 나온 버그 5건을 최소 재현 테스트로 고정 | `test_validation_regressions.py` |

## 파일

| 파일 | 하는 일 |
| --- | --- |
| `Tools/build_suneung.py` | 수능형 교재 빌드 + 입력 검증 + CSV 대조 |
| `Tools/build_ebs.py` | EBS형 교재 빌드 |
| `Tools/integrity.py` | 원문 무결성 검사 · 문장별 해설 누락 검사 |
| `Tools/top15.py` | 정답률 CSV에서 시험별 오답률 TOP15 추출 |
| `Tools/append_records.py` | 분석 결과 JSON을 나눠서 안전하게 누적 |
| `Tools/test_validation_regressions.py` | 회귀 테스트 (실행에는 비공개 예시 입력 · CSV가 필요) |

```bash
pip install -r requirements.txt   # python-docx, lxml
```

## AI 활용

조교는 Claude Code에서 슬래시 커맨드로 이 도구들을 실행합니다.
구현은 Claude Code · Codex에 작업을 지시해 진행했고, 규칙 설계 · 검증 기준 · 납품 판단은 직접 했습니다.

---

**만든 사람** · 정동근 ([GitHub](https://github.com/donggeun6669-cyber) · donggeun6669@gmail.com)
