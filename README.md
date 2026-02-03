# MacroRemover - Excelマクロ除去ツール

マクロ付きExcelファイル（.xlsm）からマクロを除去し、安全な.xlsxファイルとして保存するデスクトップアプリケーションです。

## 機能

- .xlsmファイルをアップロードしてマクロを除去
- .xlsxファイルとしてダウンロード
- 書式・レイアウト・数式を完全保持
- ドラッグ&ドロップ対応のUI
- pywebviewによるデスクトップアプリ化
- pywebview未インストール時はブラウザで起動

## 必要環境

- Python 3.8以上
- Windows / macOS / Linux

## インストール

```bash
# リポジトリをクローン
cd macro-remover

# 依存関係をインストール
pip install -r requirements.txt
```

## 使い方

### 開発環境での起動

```bash
python app.py
```

pywebviewがインストールされている場合はデスクトップアプリとして起動します。
インストールされていない場合はブラウザで `http://127.0.0.1:5001` が開きます。

### 操作方法

1. アプリケーションを起動
2. .xlsmファイルをドラッグ&ドロップ（またはクリックして選択）
3. 「マクロを除去してダウンロード」ボタンをクリック
4. マクロが除去された.xlsxファイルがダウンロードされます

## Windowsでのビルド（exe化）

```bash
# build.batを実行
build.bat
```

ビルドが完了すると `dist\MacroRemover.exe` が生成されます。

### 手動でビルドする場合

```bash
# 仮想環境を作成・有効化
python -m venv venv
venv\Scripts\activate

# 依存関係をインストール
pip install -r requirements.txt

# PyInstallerでビルド
pyinstaller --name "MacroRemover" --windowed --onefile --add-data "templates;templates" --hidden-import "webview" --hidden-import "webview.platforms.winforms" --hidden-import "webview.platforms.edgechromium" --clean --noconfirm app.py
```

## macOSでのビルド

```bash
# 仮想環境を作成・有効化
python -m venv venv
source venv/bin/activate

# 依存関係をインストール
pip install -r requirements.txt

# PyInstallerでビルド
pyinstaller --name "MacroRemover" --windowed --onefile --add-data "templates:templates" --hidden-import "webview" --hidden-import "webview.platforms.cocoa" --clean --noconfirm app.py
```

## ファイル構成

```
macro-remover/
├── app.py              # メインアプリケーション
├── templates/
│   └── index.html      # UIテンプレート
├── requirements.txt    # 依存関係
├── build.bat          # Windowsビルドスクリプト
└── README.md          # このファイル
```

## 技術仕様

- **Webフレームワーク**: Flask
- **Excel処理**: openpyxl
- **デスクトップ化**: pywebview
- **ビルドツール**: PyInstaller
- **最大ファイルサイズ**: 100MB

## ライセンス

MIT License
