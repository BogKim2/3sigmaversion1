# DCR Format Converter

<p align="center">
  <img src="https://upload.wikimedia.org/wikipedia/en/thumb/4/4e/Sumitomo_Electric_Industries_logo.svg/200px-Sumitomo_Electric_Industries_logo.svg.png" alt="Sumitomo Electric Industries" width="200"/>
</p>

<p align="center">
  <strong>DCR Format Conversion & 3-Sigma LSL/USL Calculation Tool</strong><br>
  <em>Version 1.1 | Programmed by Sangwoo Kim</em>
</p>

---

**Language / 言語 / 언어:**
[English](#english) | [日本語](#日本語) | [한국어](#한국어)

---

# English

## Table of Contents

1. [Overview](#overview)
2. [Features](#features)
3. [System Requirements](#system-requirements)
4. [Installation](#installation)
5. [Quick Start Guide](#quick-start-guide)
6. [Detailed User Guide](#detailed-user-guide)
7. [Output Structure](#output-structure)
8. [File Structure](#file-structure)
9. [Troubleshooting](#troubleshooting)

---

## Overview

**DCR Format Converter** is a professional desktop application designed for Sumitomo Electric Industries' quality control and measurement data processing workflows. It automates:

- Converting NET files and Excel data into standardized DCR format
- Processing Form Measurement Result data with TDR/Impedance analysis
- Calculating 3-Sigma LSL/USL (Lower/Upper Specification Limit) statistics
- Generating statistical visualization charts (PNG)

---

## Features

### Common Settings (Header Area)
| Field | Description |
|-------|-------------|
| **Operator Name** | Your name (required for execution) |
| **Item Name** | Product item name (e.g., `FPC`) |
| **Item Code** | Product item code (e.g., `7S3493`) |
| **Output Directory** | Base output directory (default: `{app_dir}/output/`) |
| **Output Folder Preview** | Shows real-time preview: `{ItemName}_{ItemCode}` |
| **Auto Execute All** | Runs all 3 tabs sequentially |

### Tab 1: Make DCR Format
- Parse `.NET` files (network topology)
- Extract vendor specifications from Excel
- Generate: Vendor, DE Requirement, Input Check Pin, Judge (check pin), DCR sheets
- Auto-generate Cover Page
- Statistical plots: Vendor Spec, Spec Range, Part Distribution, Judge Results, DCR Analysis

### Tab 2: Make Form Measurement Result
- **Two etching modes:**
  - **Auto Mode** - Scan a directory for all DK files automatically
  - **Manual Mode** - Select individual DK files (useful when 1-2 results are out of spec and you want to choose which conditions to include)
- Fill Impedance/TDR data from DK files
- Fill Dimension data (Circuit Width / Thickness)
- Fill LSL/Center/USL specification values
- Statistical plots: TDR Box Plot, Violin Plot, Dimension Trend, Distribution

### Tab 3: Calculate LSL USL
- 3-Sigma calculation from merged measurement data
- Auto-reference DCR file from Tab 1 output
- Statistical plots: Control Chart, Cpk Analysis, Pass/Fail Ratio, Histograms

---

## System Requirements

| Component | Requirement |
|-----------|-------------|
| OS | Windows 10 / 11 |
| Python | 3.12 or higher |
| RAM | 4 GB minimum |
| Display | 1280 x 1024 minimum |

### Dependencies
```
PySide6 >= 6.10.1
openpyxl >= 3.1.5
pandas >= 2.3.3
matplotlib >= 3.9.0
scipy >= 1.11.0
xlrd >= 2.0.2
chardet >= 5.2.0
```

---

## Installation

### Option A: Run from Source
```bash
# Clone repository
git clone https://github.com/BogKim2/3sigmaversion1.git
cd 3sigmaversion1

# Install dependencies (using uv or pip)
uv sync
# or
pip install -r requirements.txt

# Run the application
python main.py
```

### Option B: Run as Executable
- Download `DCR_Converter.exe` from the dist folder
- Place it alongside the template file `Form measurement result files_form.xlsx`
- Double-click to run

---

## Quick Start Guide

1. **Launch** the application (`python main.py` or `DCR_Converter.exe`)
2. **Fill in Common Settings:**
   - Enter your **Operator Name**
   - Enter **Item Name** and **Item Code** (output folder will be `{ItemName}_{ItemCode}`)
   - Optionally select a custom **Output Directory**
3. **Tab 1** - Select NET, Vendorspec, and Partpin files, then click **Execute**
4. **Tab 2** - Choose etching mode (Auto/Manual), select Dimension and LSLUSL files, then click **Execute**
5. **Tab 3** - Select Merged file, then click **Execute**

> **Tip:** Use **Auto Execute All** to run all three tabs sequentially with one click.

---

## Detailed User Guide

### Common Settings

These settings are shared across all tabs and appear at the top of the window.

- **Operator Name**: Required before any execution. Used in output filenames.
- **Item Name / Item Code**: Combined to create the output subfolder name. For example, if Item Name = `FPC` and Item Code = `7S3493`, output goes to `FPC_7S3493/`.
- **Output Directory**: Click **Browse** to change the base output directory. Default is `{app_dir}/output/`.

### Tab 1: Make DCR Format

**Input files required:**
| File | Format | Description |
|------|--------|-------------|
| NET File | `.NET` | Network topology file |
| Vendorspec File | `.xlsx` | Vendor specification data |
| Partpin File | `.xlsx` | Part & pin assignment data |

**Steps executed:**
1. Make Vendor Sheet
2. Make DE Requirement Sheet
3. Make Input Check Pin Sheet
4. Create int_med.xlsx (intermediate file)
5. Create final Input Check Pin sheet
6. Create Judge (check pin) sheet
7. Create DCR sheet
8. Add Cover Page
9. Generate statistical plots

**Output:** `DCR_format_yamaha_{Operator}_{Date}.xlsx`

### Tab 2: Make Form Measurement Result

**Etching File Mode:**

| Mode | When to Use |
|------|-------------|
| **Auto (Directory Scan)** | Normal case: all DK files in a directory should be processed |
| **Manual (Select Files)** | When some conditions are out of spec and you want to select only valid DK files |

- **Auto Mode**: Select an etching directory. All files matching `DK*.xls` will be processed.
- **Manual Mode**: Click **Add Files** to select individual DK files. Use **Remove Selected** or **Clear All** to manage the list.

**Additional input files:**
| File | Description |
|------|-------------|
| Dimension File | Cross-section measurement data (e.g., `7E3493-00003.xlsx`) |
| Dimension Sheet | Select the sheet to use from the dropdown |
| LSLUSL File | LSL/Center/USL specification values |

**Output:** `Form_measurement_result_{Operator}_{Date}.xlsx`

### Tab 3: Calculate LSL USL

**Input files:**
| File | Description |
|------|-------------|
| Merged File | Combined measurement data (`merged_file.xlsx`) |
| DCR File | Auto-selected from Tab 1 output |

**Output:** `Calculate_3Sigma_LSLUSL_final.xlsx`

---

## Output Structure

All output files are saved in a single organized folder:

```
{Output Directory}/
  {ItemName}_{ItemCode}/
    DCR_format_yamaha_{Operator}_{Date}.xlsx
    Form_measurement_result_{Operator}_{Date}.xlsx
    Calculate_3Sigma_LSLUSL_final.xlsx
    log_{Operator}_{Date}.dat
    plots/
      DCR_VendorSpec_{Operator}_{Date}.png
      DCR_PartDist_{Operator}_{Date}.png
      Form_TDR_BoxPlot_{Operator}_{Date}.png
      Form_TDR_Violin_{Operator}_{Date}.png
      LSLUSL_Control_{Operator}_{Date}.png
      LSLUSL_Cpk_{Operator}_{Date}.png
      ... (and more)
```

---

## File Structure

```
0105python/
  main.py                 # Entry point
  pyproject.toml          # Dependencies
  build_exe.py            # EXE build script
  Form measurement result files_form.xlsx  # Template
  DCR format base new form - yamaha.xlsx   # DCR template
  int_med.xlsx            # Intermediate template
  ui/
    __init__.py
    main_window.py        # Main window UI (PySide6)
  logic/
    __init__.py
    config_manager.py     # Config save/load (JSON)
    file_reader.py        # NET/XLSX file readers
    makevendor.py         # Vendor sheet generation
    make_de_requirement.py
    make_input_check_pin.py
    make_int_med.py
    make_judge_check_pin.py
    make_dcr.py
    make_form_measurement.py  # Form measurement + etching
    calculate_lsl_usl.py      # 3-sigma calculation
    cover_page.py             # Cover page generation
    visualizer.py             # Chart generation (matplotlib)
  data/
    files.json            # Default config
```

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "Please enter operator name" | Enter your name in the Operator Name field |
| Template file not found | Place `Form measurement result files_form.xlsx` next to the executable |
| No DK files found | Ensure DK files are named `DK*.xls` (e.g., `DK1.5.xls`) |
| DCR file not found (Tab 3) | Execute Tab 1 first to generate the DCR file |
| Plots not generated | Ensure matplotlib and scipy are installed |
| Permission error on output | Check that the output directory is writable |

---
---

# 日本語

## 目次

1. [概要](#概要)
2. [機能](#機能)
3. [システム要件](#システム要件)
4. [インストール](#インストール)
5. [クイックスタート](#クイックスタート)
6. [詳細ユーザーガイド](#詳細ユーザーガイド)
7. [出力構造](#出力構造)
8. [トラブルシューティング](#トラブルシューティング-1)

---

## 概要

**DCR Format Converter**は、住友電気工業の品質管理および測定データ処理ワークフロー向けに設計されたデスクトップアプリケーションです。以下の作業を自動化します：

- NETファイルとExcelデータの標準DCRフォーマットへの変換
- TDR/インピーダンス分析によるForm Measurement Resultデータの処理
- 3シグマ LSL/USL（下限規格値/上限規格値）統計の計算
- 統計的可視化チャート（PNG）の生成

---

## 機能

### 共通設定（ヘッダーエリア）
| フィールド | 説明 |
|-----------|------|
| **Operator Name** | 作業者名（実行前に必須入力） |
| **Item Name** | 製品アイテム名（例：`FPC`） |
| **Item Code** | 製品アイテムコード（例：`7S3493`） |
| **Output Directory** | 出力基本ディレクトリ（デフォルト：`{app_dir}/output/`） |
| **出力フォルダプレビュー** | リアルタイムプレビュー表示：`{ItemName}_{ItemCode}` |
| **Auto Execute All** | 3つのタブを順番に実行 |

### Tab 1: Make DCR Format
- `.NET`ファイル（ネットワークトポロジ）の解析
- Excelからベンダー仕様を抽出
- Vendor、DE Requirement、Input Check Pin、Judge、DCRシートを生成
- カバーページ自動生成
- 統計プロット生成

### Tab 2: Make Form Measurement Result
- **2つのエッチングモード：**
  - **自動モード** — ディレクトリ内のすべてのDKファイルを自動スキャン
  - **手動モード** — 個別のDKファイルを選択（1〜2件の結果がスペック外で、含める条件を選択したい場合に便利）
- DKファイルからインピーダンス/TDRデータを入力
- Dimensionデータの入力（回路幅/厚さ）
- LSL/Center/USL仕様値の入力
- 統計プロット生成

### Tab 3: Calculate LSL USL
- マージされた測定データからの3シグマ計算
- Tab 1出力からのDCRファイル自動参照
- 統計プロット生成（管理図、Cpk分析、合否比率、ヒストグラム）

---

## システム要件

| 項目 | 要件 |
|------|------|
| OS | Windows 10 / 11 |
| Python | 3.12 以上 |
| RAM | 4 GB 以上 |
| ディスプレイ | 1280 x 1024 以上 |

---

## インストール

### 方法A：ソースから実行
```bash
git clone https://github.com/BogKim2/3sigmaversion1.git
cd 3sigmaversion1

# 依存関係のインストール
pip install pyside6 openpyxl pandas matplotlib scipy xlrd chardet

# アプリケーション実行
python main.py
```

### 方法B：実行ファイルから実行
- `DCR_Converter.exe`をダウンロード
- テンプレートファイル `Form measurement result files_form.xlsx` を同じフォルダに配置
- ダブルクリックで実行

---

## クイックスタート

1. アプリケーションを**起動**
2. **共通設定を入力：**
   - **Operator Name**（作業者名）を入力
   - **Item Name**と**Item Code**を入力（出力フォルダ名：`{ItemName}_{ItemCode}`）
   - 必要に応じて**Output Directory**を選択
3. **Tab 1** — NET、Vendorspec、Partpinファイルを選択し、**Execute**をクリック
4. **Tab 2** — エッチングモード（自動/手動）を選択、DimensionおよびLSLUSLファイルを選択し、**Execute**をクリック
5. **Tab 3** — Mergedファイルを選択し、**Execute**をクリック

> **ヒント：** **Auto Execute All**ボタンで、ワンクリックで3つのタブを順番に実行できます。

---

## 詳細ユーザーガイド

### 共通設定

すべてのタブで共有される設定で、ウィンドウ上部に表示されます。

- **Operator Name**：実行前に必須。出力ファイル名に使用されます。
- **Item Name / Item Code**：出力サブフォルダ名の作成に使用。例：Item Name = `FPC`、Item Code = `7S3493`の場合、出力先は`FPC_7S3493/`。
- **Output Directory**：**Browse**をクリックして出力基本ディレクトリを変更。

### Tab 2: エッチングモードの選択

| モード | 使用場面 |
|--------|---------|
| **自動（ディレクトリスキャン）** | 通常：ディレクトリ内のすべてのDKファイルを処理 |
| **手動（ファイル選択）** | 一部の条件がスペック外で、有効なDKファイルのみ選択したい場合 |

- **自動モード**：エッチングディレクトリを選択。`DK*.xls`に一致するすべてのファイルが処理されます。
- **手動モード**：**Add Files**で個別のDKファイルを選択。**Remove Selected**や**Clear All**でリストを管理。

---

## 出力構造

すべての出力ファイルは整理されたフォルダに保存されます：

```
{出力ディレクトリ}/
  {ItemName}_{ItemCode}/
    DCR_format_yamaha_{Operator}_{Date}.xlsx
    Form_measurement_result_{Operator}_{Date}.xlsx
    Calculate_3Sigma_LSLUSL_final.xlsx
    log_{Operator}_{Date}.dat
    plots/
      DCR_VendorSpec_{Operator}_{Date}.png
      Form_TDR_BoxPlot_{Operator}_{Date}.png
      LSLUSL_Control_{Operator}_{Date}.png
      ...（その他）
```

---

## トラブルシューティング

| 問題 | 解決方法 |
|------|---------|
| "Please enter operator name" | Operator Nameフィールドに名前を入力 |
| テンプレートファイルが見つからない | `Form measurement result files_form.xlsx`を実行ファイルと同じ場所に配置 |
| DKファイルが見つからない | DKファイル名が`DK*.xls`形式であることを確認 |
| DCRファイルが見つからない（Tab 3） | 先にTab 1を実行してDCRファイルを生成 |

---
---

# 한국어

## 목차

1. [개요](#개요)
2. [기능](#기능-1)
3. [시스템 요구사항](#시스템-요구사항)
4. [설치 방법](#설치-방법)
5. [빠른 시작 가이드](#빠른-시작-가이드)
6. [상세 사용 가이드](#상세-사용-가이드)
7. [출력 구조](#출력-구조)
8. [문제 해결](#문제-해결)

---

## 개요

**DCR Format Converter**는 스미토모 전기공업의 품질 관리 및 측정 데이터 처리 워크플로우를 위해 설계된 데스크톱 애플리케이션입니다. 다음 작업을 자동화합니다:

- NET 파일 및 Excel 데이터를 표준 DCR 포맷으로 변환
- TDR/임피던스 분석을 통한 Form Measurement Result 데이터 처리
- 3시그마 LSL/USL (하한/상한 규격값) 통계 계산
- 통계 시각화 차트 (PNG) 생성

---

## 기능

### 공통 설정 (헤더 영역)
| 필드 | 설명 |
|------|------|
| **Operator Name** | 작업자 이름 (실행 전 필수 입력) |
| **Item Name** | 제품 아이템 이름 (예: `FPC`) |
| **Item Code** | 제품 아이템 코드 (예: `7S3493`) |
| **Output Directory** | 출력 기본 디렉토리 (기본값: `{app_dir}/output/`) |
| **출력 폴더 미리보기** | 실시간 미리보기 표시: `{ItemName}_{ItemCode}` |
| **Auto Execute All** | 3개 탭을 순차적으로 실행 |

### Tab 1: Make DCR Format
- `.NET` 파일 (네트워크 토폴로지) 파싱
- Excel에서 벤더 스펙 추출
- Vendor, DE Requirement, Input Check Pin, Judge, DCR 시트 생성
- 커버 페이지 자동 생성
- 통계 플롯 생성 (Vendor Spec, Spec Range, Part Distribution, Judge Results, DCR Analysis)

### Tab 2: Make Form Measurement Result
- **2가지 에칭 모드:**
  - **자동 모드** — 디렉토리 내 모든 DK 파일을 자동 스캔
  - **수동 모드** — 개별 DK 파일 선택 (일부 조건이 스펙 외일 때, 포함할 조건을 직접 선택 가능)
- DK 파일에서 임피던스/TDR 데이터 입력
- Dimension 데이터 입력 (회로 폭 / 두께)
- LSL/Center/USL 규격값 입력
- 통계 플롯 생성 (TDR Box Plot, Violin Plot, Dimension Trend, Distribution)

### Tab 3: Calculate LSL USL
- 병합된 측정 데이터에서 3시그마 계산
- Tab 1 출력에서 DCR 파일 자동 참조
- 통계 플롯 생성 (Control Chart, Cpk 분석, Pass/Fail 비율, 히스토그램)

---

## 시스템 요구사항

| 항목 | 요구사항 |
|------|---------|
| OS | Windows 10 / 11 |
| Python | 3.12 이상 |
| RAM | 4 GB 이상 |
| 디스플레이 | 1280 x 1024 이상 |

### 의존성 패키지
```
PySide6 >= 6.10.1
openpyxl >= 3.1.5
pandas >= 2.3.3
matplotlib >= 3.9.0
scipy >= 1.11.0
xlrd >= 2.0.2
chardet >= 5.2.0
```

---

## 설치 방법

### 방법 A: 소스에서 실행
```bash
# 저장소 클론
git clone https://github.com/BogKim2/3sigmaversion1.git
cd 3sigmaversion1

# 의존성 설치 (uv 또는 pip 사용)
uv sync
# 또는
pip install pyside6 openpyxl pandas matplotlib scipy xlrd chardet

# 애플리케이션 실행
python main.py
```

### 방법 B: 실행 파일로 실행
- `DCR_Converter.exe` 다운로드
- 템플릿 파일 `Form measurement result files_form.xlsx`를 같은 폴더에 배치
- 더블 클릭으로 실행

---

## 빠른 시작 가이드

1. 애플리케이션 **실행** (`python main.py` 또는 `DCR_Converter.exe`)
2. **공통 설정 입력:**
   - **Operator Name** (작업자 이름) 입력
   - **Item Name**과 **Item Code** 입력 (출력 폴더명: `{ItemName}_{ItemCode}`)
   - 필요시 **Output Directory** 선택
3. **Tab 1** — NET, Vendorspec, Partpin 파일 선택 후 **Execute** 클릭
4. **Tab 2** — 에칭 모드(자동/수동) 선택, Dimension 및 LSLUSL 파일 선택 후 **Execute** 클릭
5. **Tab 3** — Merged 파일 선택 후 **Execute** 클릭

> **팁:** **Auto Execute All** 버튼으로 한 번에 3개 탭을 순차 실행할 수 있습니다.

---

## 상세 사용 가이드

### 공통 설정

모든 탭에서 공유되는 설정으로, 윈도우 상단에 표시됩니다.

- **Operator Name**: 실행 전 필수 입력. 출력 파일명에 사용됩니다.
- **Item Name / Item Code**: 출력 하위 폴더명 생성에 사용. 예: Item Name = `FPC`, Item Code = `7S3493`이면 출력 경로는 `FPC_7S3493/`.
- **Output Directory**: **Browse** 클릭하여 출력 기본 디렉토리 변경 가능. 기본값은 `{app_dir}/output/`.

### Tab 1: Make DCR Format

**필요한 입력 파일:**
| 파일 | 형식 | 설명 |
|------|------|------|
| NET File | `.NET` | 네트워크 토폴로지 파일 |
| Vendorspec File | `.xlsx` | 벤더 스펙 데이터 |
| Partpin File | `.xlsx` | 파트 & 핀 배정 데이터 |

**실행 단계:**
1. Vendor 시트 생성
2. DE Requirement 시트 생성
3. Input Check Pin 시트 생성
4. int_med.xlsx 생성 (중간 파일)
5. 최종 Input Check Pin 시트 생성
6. Judge (check pin) 시트 생성
7. DCR 시트 생성
8. 커버 페이지 추가
9. 통계 플롯 생성

**출력:** `DCR_format_yamaha_{Operator}_{Date}.xlsx`

### Tab 2: Make Form Measurement Result

**에칭 파일 모드:**

| 모드 | 사용 시점 |
|------|----------|
| **자동 (디렉토리 스캔)** | 일반적인 경우: 디렉토리 내 모든 DK 파일을 처리 |
| **수동 (파일 선택)** | 일부 조건이 스펙 외일 때, 유효한 DK 파일만 선택하고 싶은 경우 |

- **자동 모드**: 에칭 디렉토리를 선택합니다. `DK*.xls` 패턴에 맞는 모든 파일이 처리됩니다.
- **수동 모드**: **Add Files**로 개별 DK 파일을 선택합니다. **Remove Selected** 또는 **Clear All**로 목록을 관리합니다.

**추가 입력 파일:**
| 파일 | 설명 |
|------|------|
| Dimension File | 단면 측정 데이터 (예: `7E3493-00003.xlsx`) |
| Dimension Sheet | 드롭다운에서 사용할 시트 선택 |
| LSLUSL File | LSL/Center/USL 규격값 |

**출력:** `Form_measurement_result_{Operator}_{Date}.xlsx`

### Tab 3: Calculate LSL USL

**입력 파일:**
| 파일 | 설명 |
|------|------|
| Merged File | 병합된 측정 데이터 (`merged_file.xlsx`) |
| DCR File | Tab 1 출력에서 자동 선택 |

**출력:** `Calculate_3Sigma_LSLUSL_final.xlsx`

---

## 출력 구조

모든 출력 파일은 하나의 정리된 폴더에 저장됩니다:

```
{출력 디렉토리}/
  {ItemName}_{ItemCode}/
    DCR_format_yamaha_{Operator}_{Date}.xlsx
    Form_measurement_result_{Operator}_{Date}.xlsx
    Calculate_3Sigma_LSLUSL_final.xlsx
    log_{Operator}_{Date}.dat
    plots/
      DCR_VendorSpec_{Operator}_{Date}.png
      DCR_PartDist_{Operator}_{Date}.png
      Form_TDR_BoxPlot_{Operator}_{Date}.png
      Form_TDR_Violin_{Operator}_{Date}.png
      LSLUSL_Control_{Operator}_{Date}.png
      LSLUSL_Cpk_{Operator}_{Date}.png
      ... (기타)
```

---

## 문제 해결

| 문제 | 해결 방법 |
|------|----------|
| "Please enter operator name" 메시지 | Operator Name 필드에 이름 입력 |
| 템플릿 파일을 찾을 수 없음 | `Form measurement result files_form.xlsx`를 실행 파일과 같은 위치에 배치 |
| DK 파일을 찾을 수 없음 | DK 파일명이 `DK*.xls` 형식인지 확인 (예: `DK1.5.xls`) |
| DCR 파일을 찾을 수 없음 (Tab 3) | Tab 1을 먼저 실행하여 DCR 파일 생성 |
| 플롯이 생성되지 않음 | matplotlib과 scipy가 설치되어 있는지 확인 |
| 출력 디렉토리 권한 오류 | 출력 디렉토리에 쓰기 권한이 있는지 확인 |

---

## License & Credits

- **Version**: 1.1
- **Programmer**: Sangwoo Kim
- **Organization**: Sumitomo Electric Industries
- **Acknowledgments**: Lots of help from Opus4.5 and Gemini
