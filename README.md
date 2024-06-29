PTT Crawler 是一個允許用戶從 PTT（批踢踢）網站爬取特定版面文章的應用程式。用戶可以選擇要爬取的版面、指定爬取的頁數,並設置關鍵字篩選文章。最終結果會保存為 Excel 和 JSON 格式的文件。

## 主要功能

- 爬取選定 PTT 版面的文章
- 根據關鍵字篩選文章 
- 按人氣排序文章
- 將結果保存為 Excel 和 JSON 文件

## 安裝與使用

1. [下載](https://github.com/shoz87/PTT-Crawler/releases/tag/Crawler) 並解壓該應用程式。
2. 確保您的系統上安裝了所需的依賴庫,包括 Python 3.6+、requests、beautifulsoup4、openpyxl、pandas 和 tkinter。
3. 執行 `ptt_crawler.exe` 文件即可啟動應用程式。
4. 在 GUI 中選擇要爬取的版面、輸入頁數和關鍵字,然後點擊"開始爬取"按鈕。
5. 爬取完成後,結果會自動保存到您的下載目錄中,包括 Excel 和 JSON 格式的文件。

## 項目結構

- `ptt_crawler.exe`：主可執行文件
- `requirements.txt`：Python 依賴包列表
- `README.md`：說明文件
