# aoiro

残高試算表と総勘定元帳を「超シンプルな青色申告、教えてもらいました」の仕訳帳に追加しました。

## 概要

このツールは、「超シンプルな青色申告、教えてもらいました」の仕訳帳Excelファイルから仕訳データを読み込み、月別残高試算表と総勘定元帳を自動生成するPythonプログラムです。

## 機能

- **仕訳帳データの読み込み**: Excelファイルから仕訳データを自動読み込み
- **月別残高試算表の生成**: 各科目の月別借方合計・貸方合計・残高を集計
- **総勘定元帳の生成**: 各科目の取引履歴を時系列で一覧表示
- **科目コード対応**: 科目コード表に基づいて科目名を自動取得

## 必要要件

- uv (Pythonパッケージマネージャー)

> **注:** Python 3.13以上と必要なライブラリ（openpyxl、pandas）は、`uv sync`実行時に自動的にインストールされます。

## インストール

### 1. uvのインストール

uvがまだインストールされていない場合は、以下の方法でインストールしてください。

**Windows (PowerShell):**
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

**macOS/Linux:**
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

**その他の方法:**
- 公式ドキュメント: https://docs.astral.sh/uv/getting-started/installation/

### 2. 依存関係のインストール

プロジェクトのルートディレクトリで以下のコマンドを実行（Pythonと必要なライブラリが自動的にインストールされます）:

```bash
uv sync
```

## 使用方法

### 基本的な使い方

デフォルトファイル(`簡単仕訳帳2026.xlsx`)を使用する場合:

```bash
uv run python main.py
```

### ファイルを指定する場合

任意のExcelファイルを指定できます:

```bash
uv run python main.py path/to/your/仕訳帳.xlsx
```

実行後、指定したExcelファイルに「月別残高試算表」と「総勘定元帳」のシートが追加されます。

### ヘルプの表示

```bash
uv run python main.py --help
```

## 出力

- **月別残高試算表シート**: 月・科目・借方合計・貸方合計・残高
- **総勘定元帳シート**: 科目・日付・摘要・相手科目・借方・貸方・残高

## ライセンス

MIT
