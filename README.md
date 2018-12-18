# batvbs

## 説明
Windows 環境におけるポータビリティの高いコンソールアプリケーションを実現するために役立つ、いくつかのバッチファイルおよび VBScript プログラムのサンプルを提示します。拡張子を除いて同名のファイルらはひとつのプログラムの構成部品です。

## 概要
|BAT|VBS|概要|
|---|---|---|
|[drop.bat](./drop.bat)|[drop.vbs](./drop.vbs)|drop.bat に対してドロップされた1つ以上のファイルを drop.vbs が処理します。各行の先頭に行番号が前置され標準出力へ出力されます。|
|[logging.bat](./logging.bat)|-|任意のコマンドの出力を、実行時のタイムスタンプを後置した名前を持つログファイルへ出力します。複数回実行される可能性のあるプログラムの実行ログを上書きせずに個別に保存する必要がある場合に有効なイディオムです。また、ログファイル名の先頭部分にはバッチファイル名が使用されます。これはログファイルと実行ファイルの対応関係を自明にするために役立ちます。|
|[listhere.bat](./listhere.bat)|-|バッチファイルが格納されていてるディレクトリと同階層に存在する全てのファイルおよびディレクトリの名前を標準入力に出力します。dir コマンドのオプション /a を使用して対象の属性（ファイル・ディレクトリ・アーカイブ・隠しファイル等）を設定することができます。|
