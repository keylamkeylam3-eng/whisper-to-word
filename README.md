# MP3 語音轉文字 Word 產生器

這是一個使用 Streamlit 與 OpenAI Whisper 製作的語音轉文字工具。

## 功能

- 上傳 MP3、WAV、M4A、OGG 或 FLAC 音訊檔
- 使用 Whisper 進行語音辨識
- 自動加入標點符號
- 依時間順序顯示辨識結果
- 產生並下載 Word（.docx）檔
- 亦可下載 TXT 純文字檔

## 使用方法

1. 開啟 Streamlit 網頁。
2. 選擇 Whisper 模型與語言。
3. 上傳 MP3 檔案。
4. 按「開始轉換為 Word 文件」。
5. 完成後按「下載 Word 檔」。

## 部署

本專案可部署於 Streamlit Community Cloud。

主程式檔案：

```text
app.py
```
