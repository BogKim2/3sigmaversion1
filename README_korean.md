# DCR 포맷 변환기

<p align="center">
  <img src="https://upload.wikimedia.org/wikipedia/en/thumb/4/4e/Sumitomo_Electric_Industries_logo.svg/200px-Sumitomo_Electric_Industries_logo.svg.png" alt="Sumitomo Electric Industries" width="200"/>
</p>

<p align="center">
  <strong>DCR 포맷 변환 및 3-시그마 LSL/USL 계산 도구</strong><br>
  <em>버전 1.0 | 개발자: 김상우</em>
</p>

---

## 📋 목차

1. [개요](#개요)
2. [주요 기능](#주요-기능)
3. [시스템 요구사항](#시스템-요구사항)
4. [설치 방법](#설치-방법)
5. [빠른 시작 가이드](#빠른-시작-가이드)
6. [상세 사용 설명서](#상세-사용-설명서)
7. [프로그램 구조](#프로그램-구조)
8. [파일 구조](#파일-구조)
9. [입출력 파일 명세](#입출력-파일-명세)
10. [문제 해결](#문제-해결)
11. [자주 묻는 질문](#자주-묻는-질문)
12. [라이선스 및 크레딧](#라이선스-및-크레딧)

---

## 개요

**DCR 포맷 변환기**는 스미토모 전기공업(Sumitomo Electric Industries)의 품질 관리 및 측정 데이터 처리 워크플로우를 위해 설계된 전문 데스크톱 애플리케이션입니다. 이 도구는 다음과 같은 복잡한 프로세스를 자동화합니다:

- NET 파일과 Excel 데이터를 표준화된 DCR 포맷으로 변환
- TDR/임피던스 분석을 포함한 Form Measurement Result 데이터 처리
- 3-시그마 LSL(하한 규격 한계) 및 USL(상한 규격 한계) 통계 계산

이 애플리케이션은 수시간의 수동 Excel 작업을 제거하고 일관되고 오류 없는 데이터 처리를 보장합니다.

---

## 주요 기능

### 탭 1: DCR 포맷 생성
- ✅ `.NET` 파일 파싱 (네트워크 토폴로지 파일)
- ✅ Excel에서 벤더 사양 추출
- ✅ DE Requirement 시트 생성
- ✅ 수식이 포함된 Input Check Pin 시트 생성
- ✅ Judge (check pin) 평가 시트 생성
- ✅ 최종 DCR 시트 생성

### 탭 2: Form Measurement Result 파일 생성
- ✅ DK 파일 처리 (DK 1.5, DK 1.6, ... DK CENTER)
- ✅ TDR (Time Domain Reflectometry) 데이터 추출
- ✅ Impedance NET resistance 값 채우기
- ✅ Dimension 데이터 처리 (회로 폭/두께)
- ✅ LSL/Center/USL 사양 적용
- ✅ 조건부 서식이 포함된 Min/Max/Average 자동 계산

### 탭 3: LSL USL 계산
- ✅ 병합된 측정 파일 처리
- ✅ 종합 통계 계산:
  - Min, Max, Average, Median, Stdev
  - IQR (사분위수 범위)
  - 1st Quartile - 4×IQR, 3rd Quartile + 4×IQR
  - AverageIfs, StdevIfs (이상치 필터링)
  - LSL = A - 3B, USL = A + 3B
- ✅ 시각화 플롯 생성 (PNG 파일)

### 추가 기능
- 🎨 스미토모 브랜딩이 적용된 Material Design UI
- 📊 실시간 진행 로깅
- 💾 자동 설정 저장
- 📁 정리된 출력 디렉토리 구조
- 📈 Matplotlib 기반 데이터 시각화
- 🔄 전체 자동 실행 기능

---

## 시스템 요구사항

| 구성요소 | 최소 사양 | 권장 사양 |
|----------|-----------|-----------|
| **운영체제** | Windows 10 (64비트) | Windows 11 (64비트) |
| **RAM** | 4 GB | 8 GB 이상 |
| **디스크 공간** | 500 MB | 1 GB |
| **Python** | 3.12 이상 | 3.12 이상 |
| **디스플레이** | 1280×720 | 1920×1080 |

### 단독 실행 파일의 경우
- Python 설치 불필요
- `DCR_Converter.exe`를 다운로드하여 바로 실행

---

## 설치 방법

### 옵션 A: 단독 실행 파일 실행 (일반 사용자 권장)

1. `dist/` 폴더에서 `DCR_Converter.exe` 다운로드
2. `files.json` 설정 파일을 같은 디렉토리에 복사
3. 더블클릭하여 실행

### 옵션 B: 소스에서 실행 (개발자용)

#### 1단계: UV 패키지 관리자 설치
```bash
# Windows (PowerShell)
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# 또는 pip 사용
pip install uv
```

#### 2단계: 프로젝트 복제 및 설정
```bash
# 프로젝트 디렉토리로 이동
cd G:\dback\00_sumitomo_sangwoo\pythonprogram\0105python

# 가상환경 생성 및 의존성 설치
uv sync

# 가상환경 활성화
.venv\Scripts\activate
```

#### 3단계: 애플리케이션 실행
```bash
python main.py
```

### 옵션 C: 소스에서 실행 파일 빌드

```bash
# 가상환경이 활성화된 상태에서 프로젝트 디렉토리에서 실행
python build_exe.py
```

실행 파일이 `dist/` 폴더에 생성됩니다.

---

## 빠른 시작 가이드

### 🚀 5분 빠른 시작

1. **실행** - 애플리케이션 실행 (`DCR_Converter.exe` 또는 `python main.py`)

2. **작업자 이름 입력** - 창 상단에서 이름 입력

3. **탭 1 - DCR 포맷:**
   - "Browse" 클릭하여 `.NET` 파일 선택
   - "Browse" 클릭하여 vendorspec Excel 파일 선택
   - "Browse" 클릭하여 partpin Excel 파일 선택
   - "Execute" 버튼 클릭

4. **탭 2 - Form Measurement:**
   - "Browse" 클릭하여 etching 디렉토리 선택 (DK 파일 포함)
   - "Browse" 클릭하여 dimension 파일 선택
   - "Browse" 클릭하여 LSLUSL 파일 선택
   - "Execute" 버튼 클릭

5. **탭 3 - Calculate LSL USL:**
   - "Browse" 클릭하여 merged 파일 선택
   - "Execute" 버튼 클릭

6. **또는 "Auto Execute All" 클릭**하여 모든 탭을 순차적으로 실행!

7. **출력 확인** - `output/` 폴더에서 결과 파일 확인

---

## 상세 사용 설명서

### 애플리케이션 시작

애플리케이션을 실행하면 다음과 같은 전문적인 인터페이스가 표시됩니다:
- **헤더**: 스미토모 전기 로고와 프로그램 버전
- **작업자 입력**: 여기에 이름 입력 (필수)
- **세 개의 탭**: 각각 다른 처리 단계용
- **Auto Execute All 버튼**: 모든 탭을 순서대로 실행

### 탭 이해하기

#### 탭 1: DCR 포맷 생성

이 탭은 원시 네트워크 및 사양 파일을 표준화된 DCR 포맷으로 변환합니다.

**필요한 입력 파일:**

| 파일 유형 | 설명 | 예시 |
|-----------|------|------|
| NET 파일 | 네트워크 토폴로지 정의 | `7S3493-994110-6CRMEV.NET` |
| Vendorspec 파일 | 벤더 사양 Excel | `099-51612-03-Sumitomo.xlsx` |
| Partpin 파일 | 부품-핀 매핑 Excel | `6CRMEV-P2-POR_回路図 - FPC r1.xlsx` |

**처리 단계:**
1. **Vendor 시트 생성** - 벤더 사양 복사
2. **DE Requirement 시트 생성** - 4-wire pair 정보 추출
3. **Input Check Pin 시트 생성** - 핀 검사 구조 생성
4. **int_med.xlsx 생성** - 중간 4W 그룹화 파일
5. **Input Check Pin Final 생성** - 모든 핀 데이터 병합
6. **Judge (check pin) 생성** - 판정 수식 추가
7. **DCR 시트 생성** - 최종 DCR 포맷 출력
8. **Cover Page 추가** - 메타데이터 및 추적성 추가

**출력:**
- `DCR_format_yamaha_{작업자}_{날짜}.xlsx`

#### 탭 2: Form Measurement Result 파일 생성

이 탭은 에칭 테스트의 측정 데이터를 처리합니다.

**필요한 입력 파일:**

| 파일 유형 | 설명 | 예시 |
|-----------|------|------|
| Etching 디렉토리 | DK*.xls 파일이 포함된 폴더 | `etching/` |
| Dimension 파일 | 회로 폭/두께 데이터 | `7E3493-00003.xlsx` |
| LSLUSL 파일 | LSL/USL 사양 값 | `LSLUSL.xlsx` |

**예상되는 DK 파일:**
- `DK1.5.xls`, `DK1.6.xls`, `DK1.7.xls`, `DK1.9.xls`
- `DK2.1.xls`, `DK2.3.xls`, `DK2.4.xls`, `DK2.5.xls`
- `DK CENTER.xls`

**처리 단계:**
1. **Form Measurement 파일 생성** - 템플릿 복사
2. **Impedance 데이터 채우기** - DK 파일에서 TDR 값 읽기
3. **Dimension 데이터 채우기** - 회로 폭/두께 추가
4. **LSL/USL 데이터 채우기** - 사양 한계 적용
5. **Cover Page 추가** - 메타데이터 추가

**출력:**
- `Form_measurement_result_{작업자}_{날짜}.xlsx`

#### 탭 3: LSL USL 계산

이 탭은 종합적인 통계 분석을 수행합니다.

**필요한 입력 파일:**

| 파일 유형 | 설명 | 예시 |
|-----------|------|------|
| Merged 파일 | 통합된 측정 데이터 | `merged_file.xlsx` |
| DCR 파일 | 탭 1의 출력 (자동 선택) | 자동 감지 |

**처리 단계:**
1. **Merged 데이터 읽기** - Method=3 데이터 필터링
2. **Cal_merged 시트 생성** - NET 개수별 재구성
3. **Sap xep 시트 생성** - 데이터 매트릭스 전치
4. **tinh LCLUCL 시트 생성** - 모든 통계 계산
5. **Calculate USL LSL 시트 생성** - DCR 데이터와 함께 최종 요약
6. **플롯 생성** - 시각화 PNG 파일 생성
7. **Cover Page 추가** - 메타데이터 추가

**사용된 통계 공식:**
```
Min = MIN(데이터_범위)
Max = MAX(데이터_범위)
Average = AVERAGE(데이터_범위)
Median = MEDIAN(데이터_범위)
Stdev = STDEV(데이터_범위)
IQR = QUARTILE(데이터, 3) - QUARTILE(데이터, 1)
1stQ-4IQR = MAX(0, QUARTILE(데이터, 1) - 4×IQR)
3rdQ+4IQR = QUARTILE(데이터, 3) + 4×IQR
AverageIfs = AVERAGEIFS(데이터, ">1stQ-4IQR", "<3rdQ+4IQR")
StdevIfs = 필터링된 데이터의 STDEV
LSL = AverageIfs - 3×StdevIfs
USL = AverageIfs + 3×StdevIfs
```

**출력:**
- `Calculate_3Sigma_LSLUSL_final.xlsx`
- `output/plots/` - PNG 시각화 파일

### Auto Execute All 사용하기

**"Auto Execute All"** 버튼을 클릭하면:
1. 탭 간 자동 전환
2. 세 가지 처리 단계 모두 실행
3. 모든 출력 파일 생성
4. 종합 로그 파일 생성

**참고:** 이 기능을 사용하기 전에 모든 입력 파일이 선택되어 있는지 확인하세요.

### 출력 파일 이해하기

모든 출력은 `output/` 디렉토리에 저장됩니다:

```
output/
├── DCR_format_yamaha_{작업자}_{날짜}.xlsx
├── Form_measurement_result_{작업자}_{날짜}.xlsx
├── Calculate_3Sigma_LSLUSL_final.xlsx
├── log_{작업자}_{날짜}.dat
└── plots/
    ├── TDR_BoxPlot_{타임스탬프}.png
    ├── Dimension_BarChart_{타임스탬프}.png
    ├── LSLUSL_Control_{타임스탬프}.png
    └── LSLUSL_Histogram_NET{n}_{타임스탬프}.png
```

### 설정 파일 (files.json)

애플리케이션은 파일 선택을 기억합니다:

```json
{
    "net_file": "파일경로/file.NET",
    "xlsx_file": "파일경로/vendorspec.xlsx",
    "partpin_file": "파일경로/partpin.xlsx",
    "etching_directory": "파일경로/etching/",
    "dimension_file": "파일경로/dimension.xlsx",
    "lslusl_file": "파일경로/LSLUSL.xlsx",
    "merged_file": "파일경로/merged_file.xlsx",
    "operator": "사용자 이름"
}
```

---

## 프로그램 구조

### 상위 수준 아키텍처

```
┌─────────────────────────────────────────────────────────────┐
│                      main.py (진입점)                       │
└─────────────────────────┬───────────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────────┐
│                   ui/main_window.py                         │
│  ┌─────────────┬─────────────┬─────────────┬──────────────┐ │
│  │   탭 1      │   탭 2      │   탭 3      │ 전체 자동    │ │
│  │ DCR 포맷   │Form Measure │ LSL/USL 계산│    실행      │ │
│  └──────┬──────┴──────┬──────┴──────┬──────┴──────────────┘ │
└─────────┼─────────────┼─────────────┼───────────────────────┘
          │             │             │
┌─────────▼─────────────▼─────────────▼───────────────────────┐
│                    logic/ (비즈니스 로직)                    │
│  ┌────────────────┐ ┌────────────────┐ ┌──────────────────┐ │
│  │ makevendor.py  │ │make_form_      │ │calculate_lsl_   │ │
│  │ make_de_req... │ │measurement.py  │ │usl.py           │ │
│  │ make_input_... │ │                │ │                  │ │
│  │ make_int_med.. │ │                │ │                  │ │
│  │ make_judge_... │ │                │ │                  │ │
│  │ make_dcr.py    │ │                │ │                  │ │
│  └────────────────┘ └────────────────┘ └──────────────────┘ │
│  ┌────────────────┐ ┌────────────────┐ ┌──────────────────┐ │
│  │ file_reader.py │ │ cover_page.py  │ │ visualizer.py    │ │
│  │config_manager  │ │                │ │ (matplotlib)     │ │
│  └────────────────┘ └────────────────┘ └──────────────────┘ │
└─────────────────────────────────────────────────────────────┘
          │
┌─────────▼───────────────────────────────────────────────────┐
│              외부 라이브러리 (의존성)                         │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌────────────────┐  │
│  │ PySide6  │ │ openpyxl │ │ pandas   │ │ matplotlib     │  │
│  │ (Qt GUI) │ │ (Excel)  │ │ (데이터) │ │ (플롯)         │  │
│  └──────────┘ └──────────┘ └──────────┘ └────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

### 모듈 설명

#### 진입점
| 파일 | 설명 |
|------|------|
| `main.py` | 애플리케이션 진입점, QApplication과 MainWindow 초기화 |

#### UI 레이어 (`ui/`)
| 파일 | 설명 |
|------|------|
| `main_window.py` | 탭, 버튼, 이벤트 핸들러가 있는 메인 GUI 창 |

#### 로직 레이어 (`logic/`)
| 파일 | 설명 | 라인 수 |
|------|------|---------|
| `makevendor.py` | 벤더 사양 시트 복사 | ~100 |
| `make_de_requirement.py` | partpin 데이터에서 DE requirement 생성 | ~200 |
| `make_input_check_pin.py` | input check pin 구조 생성 | ~300 |
| `make_int_med.py` | 4W 그룹용 NET 파일 처리 | ~200 |
| `make_judge_check_pin.py` | 판정 수식 생성 | ~150 |
| `make_dcr.py` | 최종 DCR 시트 생성 | ~250 |
| `make_form_measurement.py` | form measurement 데이터 처리 | ~400 |
| `calculate_lsl_usl.py` | 통계 계산 | ~500 |
| `file_reader.py` | NET 및 Excel 파일 읽기 | ~100 |
| `config_manager.py` | JSON 설정 저장/로드 | ~50 |
| `cover_page.py` | 표지 메타데이터 추가 | ~100 |
| `visualizer.py` | matplotlib 플롯 생성 | ~200 |

### 데이터 흐름

```
입력 파일                      처리                          출력 파일
────────                      ────                          ────────

┌──────────┐                                              ┌──────────────────┐
│ .NET     │──┐                                           │DCR_format_yamaha_│
│ 파일     │  │    ┌─────────────────────────────┐        │{작업자}_{날짜}.xlsx│
└──────────┘  ├───►│                             │───────►└──────────────────┘
┌──────────┐  │    │        탭 1 로직            │
│vendorspec│──┤    │    (7단계 순차 처리)        │        ┌──────────────────┐
│ .xlsx    │  │    │                             │        │  int_med.xlsx    │
└──────────┘  │    └─────────────────────────────┘        └──────────────────┘
┌──────────┐  │
│ partpin  │──┘
│ .xlsx    │
└──────────┘

┌──────────┐                                              ┌──────────────────┐
│DK*.xls   │──┐                                           │Form_measurement_ │
│(복수파일)│  │    ┌─────────────────────────────┐        │result_{작업자}_   │
└──────────┘  ├───►│                             │───────►│{날짜}.xlsx       │
┌──────────┐  │    │        탭 2 로직            │        └──────────────────┘
│dimension │──┤    │    (4단계 순차 처리)        │
│ .xlsx    │  │    │                             │        ┌──────────────────┐
└──────────┘  │    └─────────────────────────────┘        │ plots/*.png      │
┌──────────┐  │                                           └──────────────────┘
│ LSLUSL   │──┘
│ .xlsx    │
└──────────┘

┌──────────┐                                              ┌──────────────────┐
│merged_   │──┐                                           │Calculate_3Sigma_ │
│file.xlsx │  │    ┌─────────────────────────────┐        │LSLUSL_final.xlsx │
└──────────┘  ├───►│                             │───────►└──────────────────┘
┌──────────┐  │    │        탭 3 로직            │
│DCR 출력  │──┘    │  (통계 분석)                │        ┌──────────────────┐
│(탭1에서) │       │                             │        │ plots/*.png      │
└──────────┘       └─────────────────────────────┘        └──────────────────┘
```

---

## 파일 구조

```
0105python/
│
├── main.py                    # 애플리케이션 진입점
├── build_exe.py               # PyInstaller 빌드 스크립트
├── pyproject.toml             # 프로젝트 의존성
├── files.json                 # 사용자 설정
│
├── ui/                        # 사용자 인터페이스
│   ├── __init__.py
│   └── main_window.py         # 메인 윈도우 구현
│
├── logic/                     # 비즈니스 로직
│   ├── __init__.py
│   ├── makevendor.py          # Vendor 시트 생성
│   ├── make_de_requirement.py # DE requirement 처리
│   ├── make_input_check_pin.py# Input check pin 생성
│   ├── make_int_med.py        # 중간 파일 생성
│   ├── make_judge_check_pin.py# Judge 시트 생성
│   ├── make_dcr.py            # DCR 시트 생성
│   ├── make_form_measurement.py# Form measurement 처리
│   ├── calculate_lsl_usl.py   # LSL/USL 계산
│   ├── file_reader.py         # 파일 읽기 유틸리티
│   ├── config_manager.py      # 설정 관리
│   ├── cover_page.py          # 표지 생성
│   └── visualizer.py          # 플롯 생성
│
├── output/                    # 생성된 출력 파일
│   ├── *.xlsx                 # Excel 출력
│   ├── *.dat                  # 로그 파일
│   └── plots/                 # PNG 시각화
│
├── dist/                      # 빌드된 실행 파일
│   └── DCR_Converter.exe
│
├── 301_3sigma CAL/            # 샘플 입력 파일
│   ├── etching/               # DK 파일 디렉토리
│   ├── *.xlsx                 # 참조 Excel 파일
│   └── *.NET                  # 네트워크 파일
│
└── Form measurement result files_form.xlsx  # 템플릿 파일
```

---

## 입출력 파일 명세

### 입력 파일 형식

#### .NET 파일 형식
```
#CONT
...연속성 데이터...
%END

#4W
#Gr01
EXR4W 1,2,2049,2050
EXR4W 3,4,2051,2052
...
%END
```

#### DK Excel 파일 (Form kq 시트)
| 열 | 내용 |
|----|------|
| STT | 시퀀스 번호 |
| TDR | TDR 측정값 |
| ... | 기타 측정값 |

### 출력 파일 형식

#### DCR_format_yamaha.xlsx 시트
1. **Cover Page** - 메타데이터
2. **vendor** - 벤더 사양
3. **DE requirement** - 설계 엔지니어링 요구사항
4. **input check pin interm** - 중간 핀 데이터
5. **input check pin** - 최종 핀 검사 데이터
6. **Judge(check pin)** - 판정 결과
7. **DCR** - 최종 DCR 포맷

#### Calculate_3Sigma_LSLUSL.xlsx 시트
1. **Cover Page** - 메타데이터
2. **merged_file** - 원시 필터링 데이터
3. **Cal_merged** - 전치된 데이터
4. **Sap xep** - 재구성된 데이터
5. **tinh LCLUCL** - 통계 계산
6. **Calculate USL LSL** - 최종 요약

---

## 문제 해결

### 일반적인 문제

#### 1. "Template file not found" 오류
**원인:** 템플릿 파일 `Form measurement result files_form.xlsx`가 없습니다.

**해결책:**
- 소스 실행: 템플릿 파일이 프로젝트 루트에 있는지 확인
- 실행 파일: 템플릿을 .exe와 같은 폴더에 배치

#### 2. 시작 시 애플리케이션 충돌
**원인:** 의존성 누락 또는 설치 손상.

**해결책:**
```bash
# 의존성 재설치
uv sync --reinstall
```

#### 3. "Invalid column index" 오류
**원인:** Excel 파일의 열이 너무 많음 (>16384).

**해결책:** 이는 일반적으로 데이터 처리 오류를 나타냅니다. merged_file.xlsx에 문제가 있는지 확인하세요.

#### 4. 빈 출력 파일
**원인:** 입력 파일이 예상치 못한 형식일 수 있음.

**해결책:**
- 입력 파일 형식이 예상 구조와 일치하는지 확인
- 상세 오류 메시지는 로그 파일 확인

#### 5. 일본어 문자가 올바르게 표시되지 않음
**원인:** .NET 파일의 인코딩 문제.

**해결책:** 애플리케이션은 자동 감지를 위해 `chardet`를 사용합니다. 파일이 UTF-8 또는 Shift-JIS 인코딩으로 저장되었는지 확인하세요.

### 로그 파일 위치
자세한 처리 정보는 `output/log_{작업자}_{날짜}.dat`의 로그 파일을 확인하세요.

---

## 자주 묻는 질문

**Q: Mac이나 Linux에서 실행할 수 있나요?**
A: Python 소스 코드는 크로스 플랫폼이지만, GUI는 Windows에 최적화되어 있습니다. PySide6는 모든 플랫폼에서 작동합니다.

**Q: 처리 시간은 얼마나 걸리나요?**
A: 탭 1과 2는 몇 초 정도 걸립니다. 탭 3은 merged_file.xlsx의 크기에 따라 1-5분이 걸릴 수 있습니다 (100,000행 이상일 수 있음).

**Q: 출력 형식을 사용자 정의할 수 있나요?**
A: 네, `logic/` 디렉토리의 파일을 수정하세요. 각 모듈은 특정 출력 형식을 처리합니다.

**Q: 모든 파일을 선택하지 않고 실행하면 어떻게 되나요?**
A: 애플리케이션은 누락된 입력이 있는 단계를 건너뛰고 사용 가능한 데이터로 계속 진행합니다.

**Q: 애플리케이션을 업데이트하려면 어떻게 하나요?**
A: `dist/DCR_Converter.exe`를 새 버전으로 교체하거나, 소스 업데이트의 경우 `git pull`을 사용하세요.

---

## 라이선스 및 크레딧

### 크레딧

- **개발자:** 김상우
- **AI 지원:** Claude Opus 4.5 (Anthropic), Gemini (Google)
- **회사:** 스미토모 전기공업 주식회사 (Sumitomo Electric Industries, Ltd.)
- **버전:** 1.0

### 의존성

| 라이브러리 | 버전 | 라이선스 |
|------------|------|----------|
| PySide6 | 6.10.1+ | LGPL |
| openpyxl | 3.1.5+ | MIT |
| pandas | 2.3.3+ | BSD |
| matplotlib | 3.9.0+ | PSF |
| xlrd | 2.0.2+ | BSD |
| chardet | 5.2.0+ | LGPL |
| PyInstaller | 6.17.0+ | GPL |

### 연락처

버그 리포트 또는 기능 요청은 개발팀에 문의해 주세요.

---

<p align="center">
  <em>© 2026 스미토모 전기공업 주식회사. 모든 권리 보유.</em>
</p>
