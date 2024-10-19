# 爆速議事録アプリケーション

このアプリケーションは、音声ファイルから議事録を自動生成するツールです。

## 必要条件

- Python 3.7以上
- FFmpeg
- FFprobe

## セットアップ手順

1. リポジトリをクローンまたはダウンロードします。

2. 必要なPythonパッケージをインストールします：
   ```
   pip install -r requirements.txt
   ```

3. FFmpegとFFprobeをインストールします：

   - Windowsの場合：
     1. [FFmpeg公式サイト](https://ffmpeg.org/download.html)からFFmpegをダウンロードします。
     2. ダウンロードしたzipファイルを解凍し、中のffmpeg.exeとffprobe.exeをアプリケーションと同じディレクトリに配置します。

   - macOSの場合：
     1. Homebrewがインストールされていない場合は、まずHomebrewをインストールします。
     2. ターミナルで以下のコマンドを実行します：
        ```
        brew install ffmpeg
        ```

   - Linuxの場合：
     1. ターミナルで以下のコマンドを実行します：
        ```
        sudo apt-get update
        sudo apt-get install ffmpeg
        ```

4. 環境変数を設定します：
   - `環境変数.env`ファイルを作成し、以下の内容を記入します：
     ```
     GEMINI_API_KEY_1=your_api_key_here
     GEMINI_API_KEY_2=your_api_key_here
     ...
     GEMINI_API_KEY_10=your_api_key_here
     ```
   - 各`your_api_key_here`を実際のGemini APIキーに置き換えてください。

5. テンプレートファイルを準備します：
   - `テンプレート.docx`ファイルをアプリケーションと同じディレクトリに配置します。

6. Documentsフォルダの確認：
   - ユーザーのDocumentsフォルダが存在することを確認してください。
   - Windowsの場合：通常は`C:\Users\[ユーザー名]\Documents`にあります。
   - macOSの場合：通常は`/Users/[ユーザー名]/Documents`にあります。
   - Linuxの場合：通常は`/home/[ユーザー名]/Documents`にあります。
   - もしDocumentsフォルダが存在しない場合は、手動で作成してください。

## 使用方法

1. アプリケーションを起動します：
   ```
   python minutes_app.py
   ```

2. GUIウィンドウが開きます。

3. 音声ファイル処理：
   - 「音声ファイルを選択する」ボタンをクリックし、処理したい音声ファイル（.mp3または.wav）を選択します。
   - 「音声ファイルを処理する」ボタンをクリックして処理を開始します。

4. Excelファイル処理：
   - 「Excelファイルを選択する」ボタンをクリックし、処理したいExcelファイル（.xlsx）を選択します。
   - 「Excelファイルを処理する」ボタンをクリックして処理を開始します。

5. 処理が完了すると、結果ファイルがDocumentsフォルダに保存されます。

## 注意事項

- 大きな音声ファイルの処理には時間がかかる場合があります。
- APIキーの使用量に注意してください。
- 処理中はアプリケーションを閉じないでください。

## トラブルシューティング

問題が発生した場合は、以下を確認してください：

- すべての必要なパッケージがインストールされているか
- FFmpegとFFprobeが正しくインストールされ、パスが通っているか
- 環境変数ファイルが正しく設定されているか
- テンプレートファイルが正しい場所に配置されているか
- Documentsフォルダが存在し、アクセス可能であるか

それでも問題が解決しない場合は、ログファイル（Documents/app_log.txt）を確認し、エラーメッセージを参照してください。
