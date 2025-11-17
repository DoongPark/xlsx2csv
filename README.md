# xlsx2csv 성능 최적화 프로젝트# xlsx2csv 성능 최적화 프로젝트 - 팀 공유용# xlsx2csv 성능 최적화 프로젝트

> 빅데이터 처리를 위한 xlsx2csv 병렬 처리 최적화 연구

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)> 빅데이터 처리를 위한 xlsx2csv 병렬 처리 최적화 연구> 빅데이터 처리를 위한 xlsx2csv 병렬 처리 최적화 연구

[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE.txt)

[![Performance](https://img.shields.io/badge/Performance-3.75x-brightgreen.svg)]()

## 🎯 프로젝트 개요[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)

대용량 Excel 파일을 CSV로 변환할 때 **병렬 처리**를 통해 성능을 개선한 연구 프로젝트입니다.[![Performance](https://img.shields.io/badge/Performance-3.75x-brightgreen.svg)]()[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE.txt)

### 성과[![Performance](https://img.shields.io/badge/Performance-3.75x-brightgreen.svg)]()

- 🏆 **3.75배** 성능 향상 달성 (목표 2배의 187%)## 🎯 프로젝트 개요

- 📊 4개 실험 수행 (성공 3개, 실패 1개)

- 📝 2,172줄 상세 분석 문서## 🎯 프로젝트 개요

## 🚀 빠른 시작대용량 Excel 파일을 CSV로 변환할 때 **병렬 처리**를 통해 성능을 개선한 연구 프로젝트입니다.

````bash대용량 Excel 파일을 CSV로 변환할 때 **병렬 처리**를 통해 성능을 개선한 연구 프로젝트입니다.

# 1. 테스트 데이터 생성

python tools/generate_test_data.py### 주요 성과



# 2. 자동 최적화 실행 (권장)### 성과

python src/xlsx2csv_hybrid.py input.xlsx --output-dir output/

```- 🏆 **3.75배** 성능 향상 달성 (목표 2배의 187%)



더 자세한 내용은 [docs/QUICKSTART.md](docs/QUICKSTART.md) 참조- 📊 4개 실험 수행 (성공 3개, 실패 1개 - 학습 목적)- 🏆 **3.75배** 성능 향상 달성 (목표 2배의 187%)



## 📊 성능 비교- 📝 2,172줄 상세 분석 문서- 📊 4개 실험 수행 (성공 3개, 실패 1개)



| 파일 유형                     | 원본    | 최적화 | 향상          |- 📝 150페이지 상세 분석 문서

| ----------------------------- | ------- | ------ | ------------- |

| 다중 시트 (10개, 6.74MB)      | 2.94초  | 0.78초 | **3.75배** ⭐ |## 🚀 빠른 시작

| 대용량 단일 (200K행, 20.91MB) | 12.42초 | 5.64초 | **2.20배**    |

| 복합 (7시트, 13.77MB)         | 7.50초  | 5.04초 | **1.49배**    |## 🚀 빠른 시작



## 🔬 구현 기술### 1. 환경 설정



### 1. 멀티프로세싱 (시트 병렬)```bash

```python

# src/xlsx2csv_parallel.py```bash# 자동 최적화 (권장)

# 여러 시트를 독립된 프로세스에서 동시 처리

# → 3.75배 향상# Python 3.7 이상 필요python xlsx2csv_hybrid.py input.xlsx --output-dir output/

````

python --version```

### 2. 청크 기반 병렬 (대용량 시트)

````python

# src/xlsx2csv_chunked.py

# 대용량 단일 시트를 청크로 분할하여 병렬 처리# 필요한 라이브러리 (선택사항)더 자세한 내용은 [QUICKSTART.md](QUICKSTART.md) 참조

# → 2.20배 향상

```pip install openpyxl psutil



### 3. 하이브리드 적응형 시스템```## 📊 성능 비교

```python

# src/xlsx2csv_hybrid.py

# 파일 특성을 자동 분석하여 최적 전략 선택

# → 모든 파일에서 최고 성능### 2. 테스트 데이터 생성| 파일 유형                     | 원본    | 최적화 | 향상          |

````

| ----------------------------- | ------- | ------ | ------------- |

## 📖 문서

````bash| 다중 시트 (10개, 6.74MB)      | 2.94초  | 0.78초 | **3.75배** ⭐ |

- **[docs/QUICKSTART.md](docs/QUICKSTART.md)** - 1분 빠른 시작

- **[docs/PROJECT_README.md](docs/PROJECT_README.md)** - 프로젝트 상세 소개# 다양한 크기의 테스트 파일 생성| 대용량 단일 (200K행, 20.91MB) | 12.42초 | 5.64초 | **2.20배**    |

- **[docs/FINAL_REPORT.md](docs/FINAL_REPORT.md)** - 완전한 실험 보고서 (2,172줄)

python generate_test_data.py| 복합 (7시트, 13.77MB)         | 7.50초  | 5.04초 | **1.49배**    |

## 🛠️ 설치 및 실행



### 요구사항

```bash# 대용량 단일 시트 파일 (선택)## 🔬 구현 기술

Python 3.7+

openpyxl  # 테스트 데이터 생성용python generate_large_sheet.py

psutil    # 벤치마크용 (선택)

```### 1. 멀티프로세싱 (시트 병렬)



### 실행 예시# 복합 구조 파일 (선택)

```bash

# 다중 시트 파일python generate_hybrid_file.py```python

python src/xlsx2csv_parallel.py input.xlsx output/

```# xlsx2csv_parallel.py

# 대용량 단일 시트

python src/xlsx2csv_chunked.py input.xlsx output.csv --chunk-size 50000# 여러 시트를 독립된 프로세스에서 동시 처리



# 자동 최적화### 3. 실행# → 3.75배 향상

python src/xlsx2csv_hybrid.py input.xlsx

````

### 성능 측정```bash

````bash

# 테스트 데이터 생성# 자동 최적화 (권장) - 파일 분석 후 최적 전략 선택### 2. 청크 기반 병렬 (대용량 시트)

python tools/generate_test_data.py

python xlsx2csv_hybrid.py input.xlsx --output-dir output/

# 벤치마크 실행

python benchmark/compare_performance.py test_data/medium_test.xlsx --runs 3```python

````

# 또는 개별 전략 선택# xlsx2csv_chunked.py

## 📁 프로젝트 구조

python xlsx2csv_parallel.py input.xlsx output/ # 다중 시트# 대용량 단일 시트를 청크로 분할하여 병렬 처리

```````

xlsx2csv-performance/python xlsx2csv_chunked.py input.xlsx output.csv  # 대용량 단일 시트# → 2.20배 향상

├── 📘 docs/                       # 문서

│   ├── QUICKSTART.md``````

│   ├── PROJECT_README.md

│   └── FINAL_REPORT.md

│

├── 🚀 src/                        # 최적화 구현## 📊 성능 비교### 3. 하이브리드 적응형 시스템

│   ├── xlsx2csv.py                # 원본 (기준선)

│   ├── xlsx2csv_parallel.py       # 시트 병렬 (3.75배)

│   ├── xlsx2csv_chunked.py        # 청크 병렬 (2.20배)

│   └── xlsx2csv_hybrid.py         # 하이브리드 (자동)| 파일 유형                     | 원본    | 최적화 | 향상          |```python

│

├── 🔧 benchmark/                  # 성능 측정| ----------------------------- | ------- | ------ | ------------- |# xlsx2csv_hybrid.py

│   ├── benchmark.py

│   ├── compare_performance.py| 다중 시트 (10개, 6.74MB)      | 2.94초  | 0.78초 | **3.75배** ⭐ |# 파일 특성을 자동 분석하여 최적 전략 선택

│   └── benchmark_chunked.py

│| 대용량 단일 (200K행, 20.91MB) | 12.42초 | 5.64초 | **2.20배**    |# → 모든 파일에서 최고 성능

├── 🛠️ tools/                     # 데이터 생성

│   ├── generate_test_data.py| 복합 (7시트, 13.77MB)         | 7.50초  | 5.04초 | **1.49배**    |```

│   ├── generate_large_sheet.py

│   └── generate_hybrid_file.py

│

└── 📊 test_data/                  # 테스트 데이터## 📁 프로젝트 구조## 📖 문서

```````

## 🎓 핵심 학습

````- **[QUICKSTART.md](QUICKSTART.md)** - 1분 빠른 시작

### 성공 사례

- ✅ **Amdahl의 법칙 검증**: 큰 병목(93%) 최적화 → 3.75배 향상xlsx2csv-team/- **[PROJECT_README.md](PROJECT_README.md)** - 프로젝트 상세 소개

- ✅ **청크 기반 처리**: 대용량 단일 시트 처리 혁신

- ✅ **적응형 시스템**: 자동 최적화로 사용성 극대화├── 📘 문서- **[FINAL_REPORT.md](FINAL_REPORT.md)** - 완전한 실험 보고서 (2,124줄)



### 실패 사례 (학습)│   ├── README.md                    # 프로젝트 소개 (이 파일)

- ❌ **멀티스레딩**: 작은 병목(7%) 최적화 → 효과 없음

  - 교훈: 오버헤드 측정의 중요성, GIL의 영향│   ├── QUICKSTART.md                # 1분 빠른 시작## 🛠️ 설치 및 실행



## 📜 라이센스 및 출처│   ├── PROJECT_README.md            # 상세 프로젝트 설명



### 기반 프로젝트│   └── FINAL_REPORT.md              # 완전한 실험 보고서 (2,172줄)### 요구사항



이 프로젝트는 [dilshod/xlsx2csv](https://github.com/dilshod/xlsx2csv)를 Fork하여 **성능 최적화를 연구**한 프로젝트입니다.│



```├── 🚀 최적화 구현 (핵심)```bash

Original work Copyright (c) Dilshod Temirkhodjaev

Modified work Copyright (c) 2025 Performance Optimization Team│   ├── xlsx2csv.py                  # 원본 (기준선)Python 3.7+



Original project: https://github.com/dilshod/xlsx2csv│   ├── xlsx2csv_parallel.py         # 시트 병렬 (3.75배) ⭐openpyxl  # 테스트 데이터 생성용

License: MIT License (see LICENSE.txt)

```│   ├── xlsx2csv_chunked.py          # 청크 병렬 (2.20배) ⭐psutil    # 벤치마크용 (선택)



### 추가된 기능│   └── xlsx2csv_hybrid.py           # 하이브리드 (자동 최적화) ⭐⭐```

- ✨ 멀티프로세싱 시트 병렬 처리 (3.75배 성능 향상)

- ✨ 청크 기반 대용량 시트 처리 (2.20배 성능 향상)│

- ✨ 하이브리드 적응형 자동 최적화 시스템

- ✨ 종합 벤치마크 및 성능 분석 도구├── 🔧 도구### 실행 예시



**원본 프로젝트에 감사드립니다!** 🙏│   ├── benchmark.py                 # 성능 측정



## 🔄 재현 방법│   ├── compare_performance.py       # 순차 vs 병렬 비교```bash



### 1. 저장소 클론│   ├── benchmark_chunked.py         # 청크 벤치마크# 다중 시트 파일

```bash

git clone https://github.com/Doongpark/xlsx2csv.git│   ├── generate_test_data.py        # 테스트 파일 생성python xlsx2csv_parallel.py input.xlsx output/

cd xlsx2csv

```│   ├── generate_large_sheet.py      # 대용량 파일 생성



### 2. 테스트 데이터 생성 (선택)│   └── generate_hybrid_file.py      # 복합 파일 생성# 대용량 단일 시트

```bash

python tools/generate_test_data.py│python xlsx2csv_chunked.py input.xlsx output.csv --chunk-size 50000

python tools/generate_large_sheet.py

```└── 📊 테스트 데이터 (test_data/)



### 3. 성능 측정    └── (generate 스크립트로 생성)# 자동 최적화

```bash

# 기본 벤치마크```python xlsx2csv_hybrid.py input.xlsx

python benchmark/compare_performance.py test_data/medium_test.xlsx --runs 3

````

# 청크 병렬 벤치마크

python benchmark/benchmark_chunked.py## 🔬 구현 기술

````

### 성능 측정

### 4. 실제 사용

```bash### 1. 멀티프로세싱 (시트 병렬)

# 자동 최적화

python src/xlsx2csv_hybrid.py your_file.xlsx- 여러 시트를 독립된 프로세스에서 동시 처리```bash



# 또는 개별 전략- **3.75배 향상** (6.74MB, 10 시트)# 테스트 데이터 생성

python src/xlsx2csv_parallel.py your_file.xlsx output/

```python generate_test_data.py



## 📞 문의### 2. 청크 기반 병렬 (대용량 시트)



- GitHub Issues: https://github.com/Doongpark/xlsx2csv/issues- 대용량 단일 시트를 청크로 분할하여 병렬 처리# 벤치마크 실행

- 과제 및 연구 목적 프로젝트

- **2.20배 향상** (20.91MB, 200K 행)python compare_performance.py test_data/medium_test.xlsx --runs 3

## 📈 벤치마크 환경

````

- **OS**: macOS 15.1 (Sequoia), ARM64

- **CPU**: Apple M2, 8코어 (Performance: 4, Efficiency: 4)### 3. 하이브리드 적응형 시스템

- **RAM**: 16 GB LPDDR5

- **Python**: 3.11.11- 파일 특성을 자동 분석하여 최적 전략 선택## 📁 프로젝트 구조

- **주요 라이브러리**: openpyxl 3.1.5, psutil 7.1.3

- 모든 파일에서 최고 성능 보장

## 📊 상세 결과

````

전체 실험 과정과 분석은 **[docs/FINAL_REPORT.md](docs/FINAL_REPORT.md)** (2,172줄) 참조

## 📖 문서xlsx2csv-performance/

---

├── 📘 문서

**프로젝트 완료**: 2025년 11월 17일

**목표 달성**: 187% (2배 목표 → 3.75배 달성) 🎉- **[QUICKSTART.md](QUICKSTART.md)** - 1분 빠른 시작│   ├── QUICKSTART.md


- **[PROJECT_README.md](PROJECT_README.md)** - 프로젝트 상세 소개│   ├── PROJECT_README.md

- **[FINAL_REPORT.md](FINAL_REPORT.md)** - 완전한 실험 보고서 (2,172줄)│   └── FINAL_REPORT.md

│

## 🎓 학습 내용├── 🚀 최적화 구현

│   ├── xlsx2csv.py              # 원본 (기준선)

### 성공 사례│   ├── xlsx2csv_parallel.py     # 시트 병렬 (3.75배)

- ✅ **Amdahl의 법칙 검증**: 큰 병목(93%) 최적화 → 3.75배 향상│   ├── xlsx2csv_chunked.py      # 청크 병렬 (2.20배)

- ✅ **청크 기반 처리**: 대용량 단일 시트 처리 혁신│   └── xlsx2csv_hybrid.py       # 하이브리드 (자동)

- ✅ **적응형 시스템**: 자동 최적화로 사용성 극대화│

├── 🔧 도구

### 실패 사례 (학습)│   ├── benchmark.py

- ❌ **멀티스레딩**: 작은 병목(7%) 최적화 → 효과 없음│   ├── compare_performance.py

  - 교훈: 오버헤드 측정의 중요성, GIL의 영향│   └── generate_*.py

│

## 🛠️ 성능 측정└── 📊 테스트 데이터

    └── test_data/

```bash```

# 기본 벤치마크

python compare_performance.py test_data/medium_test.xlsx --runs 3## 🎓 핵심 학습



# 청크 병렬 벤치마크### 성공 사례

python benchmark_chunked.py

- ✅ **Amdahl의 법칙 검증**: 큰 병목(93%) 최적화 → 3.75배 향상

# 개별 파일 벤치마크- ✅ **청크 기반 처리**: 대용량 단일 시트 처리 혁신

python benchmark.py test_data/your_file.xlsx --runs 5- ✅ **적응형 시스템**: 자동 최적화로 사용성 극대화

````

### 실패 사례 (학습)

## 📦 요구사항

- ❌ **멀티스레딩**: 작은 병목(7%) 최적화 → 효과 없음

```````bash - 교훈: 오버헤드 측정의 중요성, GIL의 영향

Python 3.7+

openpyxl  # 테스트 데이터 생성용## 📜 라이센스 및 출처

psutil    # 벤치마크용 (선택)

```### 기반 프로젝트



## 🏆 프로젝트 성과이 프로젝트는 [dilshod/xlsx2csv](https://github.com/dilshod/xlsx2csv)를 Fork하여 **성능 최적화를 연구**한 프로젝트입니다.



- **목표**: 2배 성능 향상```

- **달성**: 3.75배 향상 (**187%** 달성)Original work Copyright (c) Dilshod Temirkhodjaev

- **코드**: 2,657줄 (구현) + 800줄 (도구)Modified work Copyright (c) 2025 Performance Optimization Team

- **문서**: 2,172줄 (FINAL_REPORT.md)

- **실험**: 4개 (성공 3, 실패 1)Original project: https://github.com/dilshod/xlsx2csv

License: MIT License (see LICENSE.txt)

## 📈 실험 환경```



- **OS**: macOS 15.1 (Sequoia), ARM64### 추가된 기능

- **CPU**: Apple M2, 8코어 (Performance: 4, Efficiency: 4)- ✨ 멀티프로세싱 시트 병렬 처리 (3.75배 성능 향상)

- **RAM**: 16 GB LPDDR5- ✨ 청크 기반 대용량 시트 처리 (2.20배 성능 향상)

- **Python**: 3.11.11- ✨ 하이브리드 적응형 자동 최적화 시스템

- **주요 라이브러리**: openpyxl 3.1.5, psutil 7.1.3- ✨ 종합 벤치마크 및 성능 분석 도구



## 👥 팀원 기여 방법**원본 프로젝트에 감사드립니다!** 🙏



### 1. 저장소 클론## � 재현 방법

```bash

git clone <조직-레포지토리-URL>### 1. 저장소 클론

cd xlsx2csv-team```bash

```git clone https://github.com/Doongpark/xlsx2csv.git

cd xlsx2csv

### 2. 테스트 데이터 생성```

```bash

python generate_test_data.py### 2. 테스트 데이터 생성 (선택)

``````bash

python generate_test_data.py

### 3. 실험 실행python generate_large_sheet.py

```bash```

# 성능 비교

python compare_performance.py test_data/medium_test.xlsx --runs 3### 3. 성능 측정

```bash

# 개별 실행# 기본 벤치마크

python xlsx2csv_hybrid.py test_data/medium_test.xlsxpython compare_performance.py test_data/medium_test.xlsx --runs 3

```````

# 청크 병렬 벤치마크

### 4. 변경사항 커밋python benchmark_chunked.py

`bash`

git add .

git commit -m "설명: 변경 내용"### 4. 실제 사용

git push origin main```bash

```# 자동 최적화

python xlsx2csv_hybrid.py your_file.xlsx

## 📜 라이센스

# 또는 개별 전략

이 프로젝트는 [dilshod/xlsx2csv](https://github.com/dilshod/xlsx2csv)를 기반으로 성능 최적화를 연구한 프로젝트입니다.python xlsx2csv_parallel.py your_file.xlsx output/

```

```

Original work Copyright (c) Dilshod Temirkhodjaev## 📞 문의

Modified work Copyright (c) 2025 Performance Optimization Team

- GitHub Issues: https://github.com/Doongpark/xlsx2csv/issues

License: MIT License (see LICENSE.txt)- 과제 및 연구 목적 프로젝트

```

## 📈 벤치마크 환경

## 📞 문의

- **OS**: macOS 15.1 (Sequoia), ARM64

- 프로젝트 관련 질문은 Issues 탭에 등록- **CPU**: Apple M2, 8코어 (Performance: 4, Efficiency: 4)

- 팀 내부 논의는 조직 채널 활용- **RAM**: 16 GB LPDDR5

- **Python**: 3.11.11

---- **주요 라이브러리**: openpyxl 3.1.5, psutil 7.1.3

**프로젝트 완료**: 2025년 11월 17일 ## 📊 상세 결과

**목표 달성**: 187% (2배 목표 → 3.75배 달성) 🎉

전체 실험 과정과 분석은 **[FINAL_REPORT.md](FINAL_REPORT.md)** (2,176줄) 참조

---

**프로젝트 완료**: 2025년 11월 17일  
**목표 달성**: 187% (2배 목표 → 3.75배 달성) 🎉
