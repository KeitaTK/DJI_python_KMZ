# Windowsでpythonをexe化する方法

1. 仮想環境を作成する

```powershell
python -m venv venv
```

2. 仮想環境を有効化する

- パワーシェルの管理者権限で
  
```powershell    
Set-ExecutionPolicy RemoteSigned
```
を実行してから

```powershell    
.\venv\Scripts\activate.ps1
```
プロンプトに `(venv)` が付いていることを確認します。
（もし付いていなければ前回案内の方法で仮想環境を有効化してください。）

3. **PyInstallerをインストール**

exe化に必要なPyInstallerをインストールします：

```powershell
pip install pyinstaller
```

インストール確認：
```powershell
pyinstaller --version
```

4. `tkinterdnd2` をインストール

```powershell
pip install tkinterdnd2
```

*代替パッケージとして、ARM環境などで互換性が高い* `tkinterdnd2-universal` *を試したい場合は*

```powershell
pip install tkinterdnd2-universal
```

5. インストール確認

```powershell
pip show tkinterdnd2
```

または

```powershell
pip show tkinterdnd2-universal
```

*いずれかが表示されれば成功です。*

6. 必要な他ライブラリをインストール

```powershell
pip install lxml pyperclip
```

**一括インストールする場合：**
```powershell
pip install pyinstaller tkinterdnd2 lxml pyperclip
```

7. Pythonスクリプトを実行して動作確認

```powershell
python GUI11.py
```

— ここで `ModuleNotFoundError` が出なければOKです。

8. **exe化実行**

先に動作確認した上で、PyInstaller で exe 化します。

### **基本的なexe化（アイコンなし）**
```powershell
pyinstaller --onefile --windowed --collect-all tkinterdnd2 GUI11.py
```

### **アイコンを使用したexe化**
同じディレクトリに `app.ico` がある場合、以下のようにアイコンを指定できます：

```powershell
pyinstaller --onefile --windowed --icon=app.ico --collect-all tkinterdnd2 GUI11.py
```

### **tkinterdnd2-universalを使用する場合**
```powershell
pyinstaller --onefile --windowed --icon=app.ico --collect-all tkinterdnd2-universal GUI11.py
```

## **オプション説明**
- `--onefile`: 単一の実行ファイルを作成
- `--windowed`: コンソールウィンドウを非表示（GUI用）
- `--icon=app.ico`: 指定したアイコンファイルを使用
- `--collect-all tkinterdnd2`: tkinterdnd2の全ファイルを含める

## **成果物の確認**
exe化が完了すると、`dist` フォルダ内に実行ファイルが作成されます：
```
dist/GUI11.exe
```

以上で、`No module named 'tkinterdnd2'` エラーは解消され、カスタムアイコン付きのドラッグ＆ドロップ機能が動作するexeファイルが作成されます。

⁂

: http://hachisue.blog65.fc2.com/blog-entry-861.html

: https://juu7g.hatenablog.com/entry/Python/csv/viewer

: https://pypi.org/project/tkinterdnd2/

: https://zenn.dev/takudooon/articles/b66f913e9b59b2

: https://qiita.com/kmn_qt/items/96b054edc2233c76ff19

: https://office54.net/python/tkinter/file-drag-drop

: https://maywork.net/computer/python_tkinter_drag_and_drop/

: https://ja.stackoverflow.com/questions/92126/tkinterdndをexe化できない

: https://ameblo.jp/kamekame0912/entry-12869964793.html

: https://pypi.org/project/tkinterdnd2-universal/

: https://rcmdnk.com/blog/2024/02/29/computer-python/

: https://qiita.com/bassan/items/0094379024a3e88d4d23