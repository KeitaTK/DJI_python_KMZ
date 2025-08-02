# Windowsでpythonをexe化する方法

1. 仮想環境を作成する

```powershell
python -m venv venv-dronePC
```

2. 仮想環境を有効化する

- パワーシェルの管理者権限で
  
```powershell    
Set-ExecutionPolicy RemoteSigned
```
を実行してから

```powershell    
.\venv-dronePC\Scripts\activate.ps1
```
プロンプトに `(venv-dronePC)` が付いていることを確認します。
（もし付いていなければ前回案内の方法で仮想環境を有効化してください。）

3. **必要なライブラリを一括インストール**

このスクリプトで使用されている外部ライブラリをすべてインストールします：

```powershell
pip install pyinstaller tkinterdnd2 lxml pyperclip
```

**各ライブラリの説明：**
- **pyinstaller**: exe化に必要
- **tkinterdnd2**: ドラッグ&ドロップ機能を提供
- **lxml**: XML/HTMLパーサー（高速で機能豊富）
- **pyperclip**: クリップボード操作（コピー&ペースト）

4. **個別インストール確認（オプション）**

個別に確認したい場合：

```powershell
pip show pyinstaller
pip show tkinterdnd2
pip show lxml
pip show pyperclip
```

1. **全インストール済みパッケージ確認**

```powershell
pip list
```

6. Pythonスクリプトを実行して動作確認

```powershell
python GUI57.py
```

— ここで `ModuleNotFoundError` が出なければOKです。

7. **exe化実行**

先に動作確認した上で、PyInstaller で exe 化します。

### **基本的なexe化（アイコンなし）**
```powershell
pyinstaller --onefile --windowed --collect-all tkinterdnd2 GUI57.py
```

### **アイコンを使用したexe化**
同じディレクトリに `app.ico` がある場合、以下のようにアイコンを指定できます：

```powershell
pyinstaller --onefile --windowed --icon=app.ico --collect-all tkinterdnd2 GUI57.py
```

## **オプション説明**
- `--onefile`: 単一の実行ファイルを作成
- `--windowed`: コンソールウィンドウを非表示（GUI用）
- `--icon=app.ico`: 指定したアイコンファイルを使用
- `--collect-all tkinterdnd2`: tkinterdnd2の全ファイルを含める

## **使用ライブラリ分類**

### **標準ライブラリ（インストール不要）**
- `os`, `shutil`, `zipfile`, `glob`, `threading`
- `tkinter`（および`ttk`, `messagebox`, `scrolledtext`）
- `datetime`

### **外部ライブラリ（インストール必要）**
- `tkinterdnd2`: ドラッグ&ドロップ機能
- `lxml`: XML解析
- `pyperclip`: クリップボード操作
- `pyinstaller`: exe化ツール

## **成果物の確認**
exe化が完了すると、`dist` フォルダ内に実行ファイルが作成されます：
```
dist/GUI57.exe
```

以上で、必要なライブラリがすべてインストールされ、カスタムアイコン付きのドラッグ&ドロップ機能が動作するexeファイルが作成されます。