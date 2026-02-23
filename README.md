# Excel 合併 UI 工具

## 功能
- 選擇資料夾並記住上一次路徑。
- 讀取資料夾內所有 `.xlsx` 與 `.csv`。
- 以第一行為 Header。
- 依欄位 `SN`, `CHNumber`, `TESTRESULT` 驗證：
  - 同一個 `SN` 在單一檔案內必須有 8 筆。
  - `CHNumber` 必須是 1~8。
  - `TESTRESULT` 必須全為 `PASS`。
- 合併所有合格資料後輸出 `merged_output.xlsx`。
- `CHNumber` 會依檔名自動附加 `_RT`, `_LT`, `_HT`（若無法辨識則為 `_UNKN`）。
- 若未安裝 `openpyxl`，會退回輸出 `merged_output.csv`。

## 執行
```bash
python app.py
```

> 若要完整支援 `.xlsx` 讀寫，請先安裝：
>
> ```bash
> pip install openpyxl
> ```
