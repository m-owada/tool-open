# OPEN

## 概要

名前を入力してファイルを開くツールです。

## 使い方

- テキストボックスにファイル名を入力してEnterキーを押すと、指定したアプリケーションでファイルを開きます。
- チェックボックスをONにすることで、常に手前に表示することができます。
- 存在しないファイル名を入力した場合、前方一致するファイル名の一覧をリストボックスに表示します。
- ファイルをリストボックスから選択して、ダブルクリック又はEnterキーを押すことで開くことができます。
- ファイル名が同じファイルが複数存在する場合、リストボックスを右クリックすることで選択することができます。

### 【コンパイル方法】

- 「open.cs」と同じ場所に「compile.bat」を保存して実行してください。

## 設定

- ツールを起動すると設定ファイル「config.xml」が生成されます。必要に応じて設定内容を編集してください。

### 【設定ファイル】

```
<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<config>
  <x>10</x>
  <y>10</y>
  <width>220</width>
  <height>160</height>
  <topmost>True</topmost>
  <path>C:\Users\Default\Documents</path>
  <editor>C:\Windows\notepad.exe</editor>
  <extension>txt</extension>
  <lookahead>False</lookahead>
  <trim>False</trim>
</config>
```

### 【設定内容】

|タグ     |説明                                                                |
|---------|--------------------------------------------------------------------|
|x        |ウィンドウの位置（X座標）                                           |
|y        |ウィンドウの位置（Y座標）                                           |
|width    |ウィンドウのサイズ（幅）                                            |
|height   |ウィンドウのサイズ（高さ）                                          |
|topmost  |常に手前に表示するか（True / False）                                |
|path     |参照するフォルダパスを指定（複数可）                                |
|editor   |ファイルを開くアプリケーションを指定（*1, *2）                      |
|extension|対象とする拡張子を指定（省略可）                                    |
|lookahead|アプリケーション起動時にファイルの一覧を先読みするか（True / False）|
|trim     |入力したファイル名の前後の空白を削除するか（True / False）          |

*1 editorタグの設定を省略した場合は既定のアプリケーションでファイルを開きます。  
*2 特定の拡張子のファイルのみ別のアプリケーションを指定したい場合は、editorタグのtarget属性に対象となる拡張子を指定します。
```
<!-- 拡張子bmpのファイルの場合のみ使用するアプリケーションを指定 -->
<editor target="bmp">C:\Windows\System32\mspaint.exe</editor>
<!-- 上記以外の拡張子のファイルを開く場合に使用するアプリケーションを指定 -->
<editor>C:\Windows\notepad.exe</editor>
```
