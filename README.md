# xlsx2csv 성능 최적화 프로젝트

> 빅데이터 처리를 위한 xlsx2csv 병렬 처리 최적화 연구

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE.txt)
[![Performance](https://img.shields.io/badge/Performance-3.75x-brightgreen.svg)]()

## 🎯 프로젝트 개요

대용량 Excel 파일을 CSV로 변환할 때 **병렬 처리**를 통해 성능을 개선한 연구 프로젝트입니다.

### 성과

- 🏆 **3.75배** 성능 향상 달성 (목표 2배의 187%)
- 📊 4개 실험 수행 (성공 3개, 실패 1개)
- 📝 150페이지 상세 분석 문서

## 🚀 빠른 시작

```bash
# 자동 최적화 (권장)
python xlsx2csv_hybrid.py input.xlsx --output-dir output/
```

더 자세한 내용은 [QUICKSTART.md](QUICKSTART.md) 참조

## 📊 성능 비교

| 파일 유형                     | 원본    | 최적화 | 향상          |
| ----------------------------- | ------- | ------ | ------------- |
| 다중 시트 (10개, 6.74MB)      | 2.94초  | 0.78초 | **3.75배** ⭐ |
| 대용량 단일 (200K행, 20.91MB) | 12.42초 | 5.64초 | **2.20배**    |
| 복합 (7시트, 13.77MB)         | 7.50초  | 5.04초 | **1.49배**    |

## 🔬 구현 기술

### 1. 멀티프로세싱 (시트 병렬)

```python
# xlsx2csv_parallel.py
# 여러 시트를 독립된 프로세스에서 동시 처리
# → 3.75배 향상
```

### 2. 청크 기반 병렬 (대용량 시트)

```python
# xlsx2csv_chunked.py
# 대용량 단일 시트를 청크로 분할하여 병렬 처리
# → 2.20배 향상
```

### 3. 하이브리드 적응형 시스템

```python
# xlsx2csv_hybrid.py
# 파일 특성을 자동 분석하여 최적 전략 선택
# → 모든 파일에서 최고 성능
```

## 📖 문서

- **[QUICKSTART.md](QUICKSTART.md)** - 1분 빠른 시작
- **[PROJECT_README.md](PROJECT_README.md)** - 프로젝트 상세 소개
- **[FINAL_REPORT.md](FINAL_REPORT.md)** - 완전한 실험 보고서 (2,124줄)

## 🛠️ 설치 및 실행

### 요구사항

```bash
Python 3.7+
openpyxl  # 테스트 데이터 생성용
psutil    # 벤치마크용 (선택)
```

### 실행 예시

```bash
# 다중 시트 파일
python xlsx2csv_parallel.py input.xlsx output/

# 대용량 단일 시트
python xlsx2csv_chunked.py input.xlsx output.csv --chunk-size 50000

# 자동 최적화
python xlsx2csv_hybrid.py input.xlsx
```

### 성능 측정

```bash
# 테스트 데이터 생성
python generate_test_data.py

# 벤치마크 실행
python compare_performance.py test_data/medium_test.xlsx --runs 3
```

## 📁 프로젝트 구조

```
xlsx2csv-performance/
├── 📘 문서
│   ├── QUICKSTART.md
│   ├── PROJECT_README.md
│   └── FINAL_REPORT.md
│
├── 🚀 최적화 구현
│   ├── xlsx2csv.py              # 원본 (기준선)
│   ├── xlsx2csv_parallel.py     # 시트 병렬 (3.75배)
│   ├── xlsx2csv_chunked.py      # 청크 병렬 (2.20배)
│   └── xlsx2csv_hybrid.py       # 하이브리드 (자동)
│
├── 🔧 도구
│   ├── benchmark.py
│   ├── benchmark_chunked.py
│   ├── compare_performance.py
│   └── generate_*.py
│
└── 📊 테스트 데이터
    └── test_data/
```

## 🎓 핵심 학습

### 성공 사례

- ✅ **Amdahl의 법칙 검증**: 큰 병목(93%) 최적화 → 3.75배 향상
- ✅ **청크 기반 처리**: 대용량 단일 시트 처리 혁신
- ✅ **적응형 시스템**: 자동 최적화로 사용성 극대화

### 실패 사례 (학습)

- ❌ **멀티스레딩**: 작은 병목(7%) 최적화 → 효과 없음
  - 교훈: 오버헤드 측정의 중요성, GIL의 영향

## 📜 라이센스 및 출처

### 기반 프로젝트

이 프로젝트는 [dilshod/xlsx2csv](https://github.com/dilshod/xlsx2csv)를 Fork하여 **성능 최적화를 연구**한 프로젝트입니다.

```
Original work Copyright (c) Dilshod Temirkhodjaev
Original project: https://github.com/dilshod/xlsx2csv
License: MIT License (see LICENSE.txt)
```

### 추가된 기능

- ✨ 멀티프로세싱 시트 병렬 처리 (3.75배 성능 향상)
- ✨ 청크 기반 대용량 시트 처리 (2.20배 성능 향상)
- ✨ 하이브리드 적응형 자동 최적화 시스템
- ✨ 종합 벤치마크 및 성능 분석 도구

**원본 프로젝트에 감사드립니다!** 🙏

## � 재현 방법

### 1. 저장소 클론

```bash
git clone https://github.com/Doongpark/xlsx2csv.git
cd xlsx2csv
```

### 2. 테스트 데이터 생성 (선택)

```bash
python generate_test_data.py
python generate_large_sheet.py
```

### 3. 성능 측정

```bash
# 기본 벤치마크
python compare_performance.py test_data/medium_test.xlsx --runs 3

# 청크 병렬 벤치마크
python benchmark_chunked.py
```

### 4. 실제 사용

```bash
# 자동 최적화
python xlsx2csv_hybrid.py your_file.xlsx

# 또는 개별 전략
python xlsx2csv_parallel.py your_file.xlsx output/
```

## 📞 문의

- GitHub Issues: https://github.com/Doongpark/xlsx2csv/issues
- 과제 및 연구 목적 프로젝트

## 📈 벤치마크 환경

- **OS**: macOS 15.1 (Sequoia), ARM64
- **CPU**: Apple M1 Pro, 8코어 (Performance: 2, Efficiency: 6)
- **RAM**: 16 GB LPDDR5
- **Python**: 3.11.11
- **주요 라이브러리**: openpyxl 3.1.5, psutil 7.1.3

## 📊 상세 결과

전체 실험 과정과 분석은 **[FINAL_REPORT.md](FINAL_REPORT.md)** (2,176줄) 참조

---

**프로젝트 완료**: 2025년 11월 17일  
**목표 달성**: 187% (2배 목표 → 3.75배 달성) 🎉
