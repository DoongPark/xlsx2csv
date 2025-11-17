# xlsx2csv 성능 개선 프로젝트

## 📁 프로젝트 구조

```
xlsx2csv-performance/
├── 📘 docs/                             # 문서
│   ├── QUICKSTART.md                    # 1분 빠른 시작
│   ├── PROJECT_README.md                # 프로젝트 상세 설명
│   └── FINAL_REPORT.md                  # 종합 보고서 (2,172줄)
│
├── 🚀 src/                              # 최적화 구현
│   ├── xlsx2csv.py                      # 원본 (순차 처리)
│   ├── xlsx2csv_parallel.py             # 시트 병렬 처리 (3.75배) ⭐
│   ├── xlsx2csv_chunked.py              # 청크 병렬 처리 (2.20배) ⭐
│   └── xlsx2csv_hybrid.py               # 하이브리드 (자동 최적화) ⭐⭐
│
├── 🔧 benchmark/                        # 성능 측정
│   ├── benchmark.py                     # 기본 성능 측정
│   ├── compare_performance.py           # 순차 vs 병렬 비교
│   └── benchmark_chunked.py             # 청크 병렬 벤치마크
│
├── 🛠️ tools/                           # 데이터 생성
│   ├── generate_test_data.py            # 다중 시트 테스트 파일
│   ├── generate_large_sheet.py          # 대용량 단일 시트 파일
│   └── generate_hybrid_file.py          # 복합 테스트 파일
│
└── 📊 test_data/                        # 테스트 데이터
    ├── small_test.xlsx                  # 0.46 MB, 5시트
    ├── medium_test.xlsx                 # 6.74 MB, 10시트
    ├── small_single_50k.xlsx            # 5.23 MB, 50K행
    ├── medium_single_100k.xlsx          # 10.46 MB, 100K행
    ├── large_single_200k.xlsx           # 20.91 MB, 200K행
    └── hybrid_test.xlsx                 # 13.77 MB, 7시트 (복합)
```

## 🎯 사용 방법

### 빠른 시작 (권장)

```bash
# 자동 최적화 - 파일 분석 후 최적 전략 선택
python src/xlsx2csv_hybrid.py input.xlsx --output-dir output/
```

### 개별 전략 선택

```bash
# 다중 시트 파일 → 시트 병렬 처리
python src/xlsx2csv_parallel.py input.xlsx output/

# 대용량 단일 시트 → 청크 병렬 처리
python src/xlsx2csv_chunked.py input.xlsx output.csv --chunk-size 50000 --workers 4

# 원본 순차 처리
python src/xlsx2csv.py input.xlsx output.csv
```

## 📊 성능 결과

| 파일 유형                     | 원본    | 최적화 | 향상       |
| ----------------------------- | ------- | ------ | ---------- |
| 다중 시트 (10개, 6.74MB)      | 2.94초  | 0.78초 | **3.75배** |
| 대용량 단일 (200K행, 20.91MB) | 12.42초 | 5.64초 | **2.20배** |
| 복합 (7시트, 13.77MB)         | 7.50초  | 5.04초 | **1.49배** |

## 🔬 실험 요약

| #   | 실험                     | 결과        | 상태    |
| --- | ------------------------ | ----------- | ------- |
| 1   | 멀티프로세싱 (시트 병렬) | 3.75배      | ✅ 성공 |
| 2   | 멀티스레딩 (초기화)      | 0.98배      | ❌ 실패 |
| 3   | 청크 기반 병렬 (대용량)  | 2.20배      | ✅ 성공 |
| 4   | 하이브리드 시스템        | 자동 최적화 | ✅ 성공 |

## 📖 상세 문서

전체 실험 과정, 분석, 결과는 **[FINAL_REPORT.md](FINAL_REPORT.md)** 참조

- 4개 실험 상세 분석
- 성공과 실패 모두 기록
- 이론적 배경 및 실측 데이터
- 재현 가능한 벤치마크

## 🎓 핵심 학습

- ✅ **Amdahl의 법칙 검증**: 큰 병목(93%) 최적화 → 3.75배, 작은 병목(7%) → 효과 없음
- ✅ **오버헤드 측정**: 프로세스 생성 비용 < 이득 → 성공, 스레드 오버헤드 > 이득 → 실패
- ✅ **적응형 최적화**: 파일 특성 분석 후 자동으로 최적 전략 선택

## 📦 요구사항

```bash
Python 3.7+
openpyxl  # 테스트 데이터 생성용
psutil    # 벤치마크용 (선택)
```

## 🏆 프로젝트 성과

- **목표**: 2배 성능 향상
- **달성**: 3.75배 향상 (**187%** 달성)
- **코드**: 2,657줄 (구현 + 벤치마크 + 도구)
- **문서**: 2,124줄 (62KB)
- **테스트 데이터**: 6개 파일 (57MB)

---

**작성일**: 2025년 11월 17일  
**환경**: macOS ARM64, 8코어, Python 3.11.11
