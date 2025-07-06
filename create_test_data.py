#!/usr/bin/env python3
"""建立測試用的 Excel 檔案"""

import pandas as pd
from datetime import datetime, timedelta
import random

# 測試資料
approvers = ["張三", "李四", "王五", "趙六", "陳七"]
departments = ["IT部", "財務部", "人資部", "業務部", "行政部"]
statuses = ["待審核", "進行中", "已完成"]

# 建立測試資料
data = []
for i in range(50):
    data.append({
        "ID": f"REQ-{i+1:04d}",
        "申請人": f"員工{random.randint(1, 20)}",
        "部門": random.choice(departments),
        "Approver": random.choice(approvers),
        "申請日期": datetime.now() - timedelta(days=random.randint(0, 30)),
        "金額": random.randint(1000, 50000),
        "狀態": random.choice(statuses),
        "備註": f"測試備註 {i+1}"
    })

# 建立 DataFrame 並儲存
df = pd.DataFrame(data)
df.to_excel("test_data.xlsx", index=False, engine='openpyxl')
print("測試檔案 test_data.xlsx 已建立！")