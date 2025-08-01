# Excel Splitter for SharePoint

這個工具可以將 Excel 檔案依據 Approver 欄位自動拆分成多個子檔案，方便在 SharePoint 上分享給不同的審核者。

## 功能特色

- 自動讀取 Excel 中的 Approver 欄位
- 為每個 Approver 建立獨立資料夾
- 保留 Excel 格式和篩選功能（AutoFilter）
- 支援 SharePoint 同步機制

## 安裝

1. 確保已安裝 Python 3.9 或更高版本
2. 安裝必要套件：

```bash
pip install -r requirements.txt
```

## 使用方式

### 基本用法

```bash
python splitter.py <Excel檔案路徑>
```

### 範例

```bash
# 本機檔案
python splitter.py /Users/name/Documents/master.xlsx

# SharePoint 掛載路徑
python splitter.py "/Volumes/SharePoint/Sites/MyTeam/Documents/approval_list.xlsx"
```

## 執行流程

1. 程式會讀取指定的 Excel 檔案
2. 找出所有唯一的 Approver
3. 在原始檔案的同一層目錄下建立子資料夾（以 Approver 名稱命名）
4. 複製原始檔案到每個子資料夾
5. 套用 AutoFilter 篩選，只顯示該 Approver 的資料

## 檔案結構範例

執行前：
```
Documents/
└── master.xlsx
```

執行後：
```
Documents/
├── master.xlsx
├── 張三/
│   └── master.xlsx (只顯示張三的資料)
├── 李四/
│   └── master.xlsx (只顯示李四的資料)
└── 王五/
    └── master.xlsx (只顯示王五的資料)
```

## SharePoint 整合

### 方法 1：手動分享
1. 執行完程式後，將整個資料夾結構上傳到 SharePoint
2. 分享各子資料夾給對應的 Approver
3. 當 Approver 在自己的檔案中進行修改，SharePoint 會自動同步回母檔

### 方法 2：Power Automate 自動化（推薦）
適用於有 Power Automate Premium 授權的組織，無需 Azure AD 應用程式註冊。

1. 使用 `excel_splitter_onedrive_fixed.ipynb` 處理檔案
2. 程式會自動產生 Power Automate 所需的檔案：
   - `power_automate_input.xlsx` - 包含所有審查者資訊
   - `power_automate_instructions.html` - 詳細設定說明
3. 匯入提供的 `power_automate_sharepoint_permissions.json` 到 Power Automate
4. 上傳 Excel 檔案到 OneDrive/SharePoint，Power Automate 會自動處理權限設定

詳細說明請參考 [Power Automate 整合指南](POWER_AUTOMATE_INTEGRATION_GUIDE.md)

### 方法 3：直接 API 整合
需要 Azure AD 應用程式註冊，適用於可以註冊應用程式的環境。

使用 `excel_splitter_interface_sharepoint_enhanced.ipynb`，提供：
- 瀏覽器 SSO 認證
- 自動查詢使用者 Email
- 一鍵分享到 SharePoint
- 即時進度追蹤

詳細說明請參考 [SharePoint 整合指南](SHAREPOINT_INTEGRATION_GUIDE.md)

## 注意事項

- Excel 檔案必須包含名為 "Approver" 的欄位（大小寫需完全相符）
- 確保有足夠的磁碟空間（會複製多份檔案）
- 如果 Approver 名稱包含特殊字元，可能會影響資料夾建立
- 建議先用小檔案測試，確認運作正常後再處理大檔案

## 常見問題

### Q: 找不到 Approver 欄位？
A: 請確認 Excel 中的欄位名稱是否完全為 "Approver"（注意大小寫）

### Q: 可以修改篩選的欄位名稱嗎？
A: 可以修改 `splitter.py` 中的 'Approver' 字串為你需要的欄位名稱

### Q: 支援其他檔案格式嗎？
A: 目前只支援 .xlsx 格式，不支援 .xls 或 .csv

## 授權

MIT License