# 同業送件明細自動產出網頁

這是一個免費、免登入的小型內部網頁。每週先手動下載「申報案件彙總表」Excel，再從網頁上傳；系統會依教學檔規則篩選上市 / 上櫃與指定案件類別，並嘗試從公開資訊觀測站 / TWSE 公開資料補股票簡稱、CB/ECB 商品簡稱與本次籌資計畫，產出每週 Excel 並提供下載，Email 範本則直接顯示在頁面上可複製。

## 使用方式：本機

1. 在公司可上網的 Windows 電腦上打開 `啟動正式版.bat`。
2. 瀏覽器會自動開啟 `http://localhost:8796`。
3. 選擇本週的「申報案件彙總表」Excel，按「用上傳檔產出」。
4. 產出的檔案會放在 `reports` 資料夾，也會出現在網頁下載清單。

若要讓同事一起用，把這台電腦的內部 IP 分享給同事，例如：

```text
http://公司電腦IP:8796
```

## 使用方式：GitHub + Render 免費部署

這個專案可以直接放到 GitHub，再用 Render 免費 Web Service 部署，這樣同事用 Render 給的網址即可，不需要開著你的電腦。

1. 建議建立一個 GitHub private repository。
2. 把此資料夾內容推到 GitHub repository 根目錄。
3. 到 Render 新增 Blueprint 或 Web Service，連到該 GitHub repository。
4. Render 會讀取 `render.yaml`，使用 `python server.py` 啟動。
5. 部署完成後，把 Render 網址給同事使用。

注意：免費主機的檔案儲存不是永久資料庫，產出的 Excel 應在頁面產出後直接下載；每週重新上傳來源檔即可。

## 產出內容

Excel 會包含：

- 115 年統計頁
- 115 年本次籌資計畫頁
- 本週新增藍字標示
- CB/ECB 第幾次名稱
- 本次籌資計畫原因
- Email 範本顯示在網頁上

若公司網路擋住公開資訊觀測站，或 MOPS 查詢頁改版，網頁會顯示「MOPS 待確認」清單，不會把查不到的資料靜默當成正確答案。

週期預設為最近一個週四往前推 6 天，也就是教學檔裡的週五到週四。

## 成本與帳號

- 不需要登入個人帳號。
- 不需要 Google、Microsoft 365、Notion、雲端資料庫或付費服務。
- 只需要一台可上網的公司電腦與免費 Python。

## 注意事項

篩選邏輯在 `server.py` 的 `COMPANY_TYPES`、`CASE_KEYWORDS` 和日期欄位區段；MOPS 查詢不到時的備援對照在 `SECURITY_SHORT_NAMES`、`BOND_SHORT_NAMES`、`MOPS_ENRICHMENTS`。

## 自檢

部署前可執行：

```bash
python scripts/self_check.py
```

若要用實際來源檔完整測試：

```bash
SAMPLE_SOURCE_XLSX=/path/to/申報案件彙總表.xlsx python scripts/self_check.py
```
