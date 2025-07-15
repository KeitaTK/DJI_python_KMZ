

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
3. `tkinterdnd2` をインストール

```powershell
pip install tkinterdnd2
```

*代替パッケージとして、ARM環境などで互換性が高い* `tkinterdnd2-universal` *を試したい場合は*

```powershell
pip install tkinterdnd2-universal
```

4. インストール確認

```powershell
pip show tkinterdnd2
```

または

```powershell
pip show tkinterdnd2-universal
```

*いずれかが表示されれば成功です。*
5. 必要な他ライブラリを再インストール（念のため）

```powershell
pip install lxml
```

6. Pythonスクリプトを実行して動作確認

```powershell
python GUI11.py
```

— ここで `ModuleNotFoundError` が出なければOKです。
7. （exe化済みの場合）再ビルド
先に動作確認した上で、PyInstaller で exe 化し直します。tkinterdnd2 を含めるには：

```powershell
pyinstaller --onefile --windowed --collect-all tkinterdnd2 GUI11.py
```

あるいは universal を使うなら

```powershell
pyinstaller --onefile --windowed --collect-all tkinterdnd2-universal GUI11.py
```


以上で、`No module named 'tkinterdnd2'` エラーは解消され、ドラッグ＆ドロップ機能が動作するはずです。

<div style="text-align: center">⁂</div>

[^1]: http://hachisue.blog65.fc2.com/blog-entry-861.html

[^2]: https://juu7g.hatenablog.com/entry/Python/csv/viewer

[^3]: https://pypi.org/project/tkinterdnd2/

[^4]: https://zenn.dev/takudooon/articles/b66f913e9b59b2

[^5]: https://qiita.com/kmn_qt/items/96b054edc2233c76ff19

[^6]: https://office54.net/python/tkinter/file-drag-drop

[^7]: https://maywork.net/computer/python_tkinter_drag_and_drop/

[^8]: https://ja.stackoverflow.com/questions/92126/tkinterdndをexe化できない

[^9]: https://ameblo.jp/kamekame0912/entry-12869964793.html

[^10]: https://pypi.org/project/tkinterdnd2-universal/

[^11]: https://rcmdnk.com/blog/2024/02/29/computer-python/

[^12]: https://qiita.com/bassan/items/0094379024a3e88d4d23

