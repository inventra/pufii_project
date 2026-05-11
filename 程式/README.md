# 帕妃程式碼產出說明

## 程式位置

`/Users/openclaw/創旭/帕妃/程式/generate_live_table.py`

## 建議執行方式：一鍵重跑全套 + 統一待補資料

```bash
/opt/homebrew/bin/python3 /Users/openclaw/創旭/帕妃/程式/run_all_pafei_outputs.py --date 0414 --year 2026
```

這支會：
1. 確認 `商品資料大檔.xlsx` alias/symlink 存在。
2. 清空 `輸出檔/` 後重跑直播、91、官網收音機整包全部產出。
3. 讀取所有 `.audit.json`。
4. 把所有缺漏統一彙整到：
   `/Users/openclaw/創旭/帕妃/輸出檔/待補資料.txt`

## 單支直播表格執行方式

```bash
/opt/homebrew/bin/python3 /Users/openclaw/創旭/帕妃/程式/generate_live_table.py --date 0414
```

會產出：

`/Users/openclaw/創旭/帕妃/輸出檔/0414直播表格.xlsx`

同時產出紀錄：

`/Users/openclaw/創旭/帕妃/輸出檔/0414直播表格.xlsx.audit.json`

## 如何知道檔案是程式碼產出的

每次程式執行會產生 `.audit.json`，裡面包含：

- 執行時間
- 使用的 script 路徑
- 使用的日期參數
- 輸入檔路徑與 SHA256
- 輸出檔路徑與 SHA256
- 筆數
- 還沒確認的規則
- 找不到或待確認欄位

如果要驗證，可以：

1. 看輸出檔旁邊是否有 `.audit.json`
2. 看 audit 裡面的 `script` 是否指到 `generate_live_table.py`
3. 重新執行同一支程式、同一批輸入檔、同一日期
4. 比對輸出檔 SHA256 或內容是否一致

## 目前程式已整理的規則

### 已確認
- 檔案名稱固定 `{日期}直播表格.xlsx`，同檔名迭代更新。
- 分頁：`{日期}`、`對款`、`姐`、`RAY`。
- `標數` 空白。
- 原價：`ray價錢` L 欄。
- 直播價：`ray價錢` O 欄。
- RAY 成本候選：`ray價錢` K 欄。
- 尺寸表用流水號去 `尺寸表&試穿報告.xlsx` 各尺寸分頁找。
- `上衣+下身` 找到時，第二段下身尺寸要多讀 Q-W 欄。
- 尺寸欄位名稱原則保留來源欄名，不自行把 `大腿` 改成 `大腿寬`、`下擺` 改成 `下擺寬`。

### 待確認，所以程式先不定案
- 顏色來源。
- 主表成本是否固定 500 或人工異動。
- 推薦尺寸若尺寸表空白但正確檔有值，來源在哪。
- P33315 不同分頁品名差異。


# 91 新品上架檔案程式

## 程式位置
`/Users/openclaw/創旭/帕妃/程式/generate_91_listing.py`

## 執行方式
```bash
/opt/homebrew/bin/python3 /Users/openclaw/創旭/帕妃/程式/generate_91_listing.py --date 0414 --year 2026
```

## 輸入
- `/Users/openclaw/創旭/帕妃/輸入檔/商品資料大檔.xlsx`
- `/Users/openclaw/創旭/帕妃/輸入檔/尺寸表&試穿報告.xlsx`

## 輸出
- `/Users/openclaw/創旭/帕妃/輸出檔/91-0414新品上架的檔案.xlsx`
- `/Users/openclaw/創旭/帕妃/輸出檔/91-0414新品上架的檔案.xlsx.audit.json`

## 目前產出結果
- 分頁：`商品資料`
- 欄位數：50 欄
- 最後欄：AX `隱賣商品`
- 篩選條件：`商品資料(大檔)` C 欄 `用途 = 上架`
- 0414 目前篩出 1 筆。
- Q-AE 商品選項/商品圖檔欄位尚未確認，先留空。


## 商品料號規則
- `商品料號` 欄位取 `商品名稱` 最後一組 `【】` 裡面的文字。
- 範例：`PUFII- ...【CP33034】` → `CP33034`。
- 已寫入程式 `generate_91_listing.py`，並重新產出 `91-0414新品上架的檔案.xlsx`。


## 2026-05-09 程式草稿與試跑
- 程式：`/Users/openclaw/創旭/帕妃/程式/generate_website_radio.py`
- 已試跑：`--date 0414`
- 輸出資料夾：`/Users/openclaw/創旭/帕妃/輸出檔/0414官網上架匯入收音機的檔案`
- 輸出檔：`/Users/openclaw/創旭/帕妃/輸出檔/0414官網上架匯入收音機的檔案/0414官網上架匯入收音機的檔案.xlsx`
- audit：`/Users/openclaw/創旭/帕妃/輸出檔/0414官網上架匯入收音機的檔案/0414官網上架匯入收音機的檔案.xlsx.audit.json`
- 結果：暫沿用 `商品資料(大檔)` C 欄 `用途=上架`，目前篩出 1 筆。
- 欄位確認：共 15 欄，最後一欄 O `原價標`。
