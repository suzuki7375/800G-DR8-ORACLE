# Excel 合併 UI 工具

## 功能
- 選擇資料夾並記住上一次路徑。
- 讀取資料夾內所有 `.xlsx` 與 `.csv`。
- 以第一行為 Header。
- 依欄位 `TESTSN`（或 `SN`）, `CHNumber`, `TESTRESULT` 驗證：
  - 每個檔案中同一個 `TESTSN` 必須涵蓋 `CHNumber` 1~8。
  - 若同一個 `CHNumber` 有多筆 `PASS`，會以 `TESTDATE` 最新的一筆為準。
  - 任一 `CHNumber` 沒有 PASS（或缺號）即略過該 `TESTSN`。
- 合併後會再檢查同一個 `TESTSN` 是否剛好 24 筆（例如 RT/LT/HT 各 8 筆），只有符合者才會輸出（不足 24 筆不輸出）。
- 若沒有任何 `TESTSN` 達到 24 筆，會顯示錯誤訊息。
- 合併所有合格資料後輸出 `merged_output.xlsx`。
- 輸出排序為同一 `TESTSN` 下依序：`1~8_RT`、`1~8_LT`、`1~8_HT`。
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
