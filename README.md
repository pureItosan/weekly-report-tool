# 週報整合工具 · Weekly Report Hub

部門週報的「匯入 → 人員對應 → 疊加比對 → Word 匯出」單頁工具。
純前端（HTML/CSS/JS），資料存在瀏覽器 localStorage，沒有後端、可離線使用、也可放到 GitHub Pages。

## 怎麼啟動
1. 在這個資料夾開終端機，執行：
   ```
   python -m http.server 8848
   ```
2. 瀏覽器打開： http://localhost:8848/index.html
3. （或直接把 `index.html` 丟到 GitHub Pages 也可以。）

## 使用流程
1. **成員名單**（左側）：貼上或上傳清單，一行一位。可寫別名：`Raveendra: Ravi, Rav`。
   名單會存進瀏覽器，下次打開還在；匯入報告不會洗掉名單。
   - **依報告自動建立成員**：勾選「匯入報告時自動把負責人加入名單」（預設開），匯入後報告裡的 Owner 會自動成為成員，沒當 Owner 的人才會是 Pending。也可手動按「＋ 從報告補齊成員」。
   - 名單一變動會**自動重新分派**所有任務。
2. **匯入週報檔**（右上）：支援 PPTX / DOCX / XLSX / TXT / CSV，可一次多檔。
   - PPTX 會解析 Reporter / Project / Current Job and Issue / Risk / Due / Owner / Next Week，
     並抽出投影片圖片關聯到對應任務。
   - 任務分派以 **Owner 優先**；多人（`Eric/Isha`、`Rick, Jin, John`、`Sam / Adrian`）會分到每個人底下並標 **Shared owner**；對不到的人放 **Unassigned**。
3. **疊加比對**：下週再匯入新檔，不會覆蓋舊資料，會用 project/owner/描述產生穩定 key，
   標示 **New / Updated / Unchanged**，Updated 會顯示 **Weekly delta** 並保留前次內容。
4. **成員工作台**（右上）：成員自行選名字、填 project/risk/due/progress/complexity/本週/問題/下週，
   有提示句型按鈕（Debug issue、Verification…），可上傳 issue 圖片（縮圖點擊放大）。
5. **篩選**（成員區上方工具列）：關鍵字搜尋、依**專案**篩選、**成員快選**跳到某人、「只顯示有任務的成員」隱藏 Pending。
6. **匯出 Word**：右上「匯出全部 Word」，或在任務詳情裡「匯出此成員 Word」。
   格式為精簡專業週報（`*成員名` / `This week: [ Project - % | … ]` / 編號工作 + Status/Analysis/Shared owner/Attachments / Next week），沒有內容的成員也會輸出 `Pending input`。

## 檔案
- `index.html` / `styles.css` / `app.js` — 主程式
- `vendor/` — JSZip、SheetJS（本地內嵌，離線可用）
- `real.pptx` — 你的範例週報（測試用，可刪）
- `make_test_pptx.py` / `test.pptx` — 合成測試檔（可刪）

## 已驗證（用你的 real.pptx，43 筆任務 / 69 張圖）
- 解析成員與任務、Reporter 與多人 Owner 分派、Shared owner、Unassigned
- 模糊對應：`Ravi → Raveendra`
- 圖片抽出並關聯任務、Word 內嵌圖片
- 疊加判斷 New / Updated / Unchanged 與 Weekly delta
- 重新匯入報告不會洗掉成員名單
