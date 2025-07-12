#!/usr/bin/env python3
"""
實戰痛點測試 - 用真實檔案操作測試各種會爆炸的情況
"""

import os
import tempfile
import pandas as pd
import time
import tracemalloc
from pathlib import Path

def create_problematic_test_files():
    """建立各種有問題的測試檔案"""
    print("建立測試用的地雷檔案...")
    
    test_dir = "/Users/davidshih/projects/excel/test_files"
    os.makedirs(test_dir, exist_ok=True)
    
    # 1. 正常檔案
    normal_data = pd.DataFrame({
        'Reviewer': ['張三', '李四', '王五'] * 10,
        'Data': range(30)
    })
    normal_path = os.path.join(test_dir, "正常檔案.xlsx")
    normal_data.to_excel(normal_path, index=False)
    print(f"✓ 建立正常檔案: {normal_path}")
    
    # 2. 空檔案
    empty_path = os.path.join(test_dir, "空檔案.xlsx")
    pd.DataFrame().to_excel(empty_path, index=False)
    print(f"✓ 建立空檔案: {empty_path}")
    
    # 3. 超大檔案
    if not os.path.exists(os.path.join(test_dir, "超大檔案.xlsx")):
        print("建立超大檔案... (這會花點時間)")
        big_data = pd.DataFrame({
            'Reviewer': [f'評審員_{i%100}' for i in range(50000)],
            'Data': range(50000),
            'Extra1': ['測試資料'] * 50000,
            'Extra2': [f'更多資料_{i}' for i in range(50000)]
        })
        big_path = os.path.join(test_dir, "超大檔案.xlsx")
        big_data.to_excel(big_path, index=False)
        print(f"✓ 建立超大檔案: {big_path} ({os.path.getsize(big_path)/1024/1024:.1f}MB)")
    
    # 4. 特殊字元檔案
    special_data = pd.DataFrame({
        'Reviewer': ["O'Brien", "José María", "李明 (Ming)", "Smith & Co.", "Test/User"],
        'Data': range(5)
    })
    special_path = os.path.join(test_dir, "特殊字元!@#$%.xlsx")
    special_data.to_excel(special_path, index=False)
    print(f"✓ 建立特殊字元檔案: {special_path}")
    
    # 5. 假Excel檔案 (實際是文字)
    fake_path = os.path.join(test_dir, "假Excel.xlsx")
    with open(fake_path, 'w', encoding='utf-8') as f:
        f.write("這不是Excel檔案,只是改了副檔名")
    print(f"✓ 建立假Excel檔案: {fake_path}")
    
    return test_dir

def test_file_reading_robustness():
    """測試檔案讀取的堅固性"""
    print("\n測試檔案讀取堅固性...")
    print("=" * 50)
    
    test_dir = create_problematic_test_files()
    
    test_files = [
        ("正常檔案.xlsx", "應該正常"),
        ("空檔案.xlsx", "應該警告空資料"),
        ("超大檔案.xlsx", "應該有效能警告"),
        ("特殊字元!@#$%.xlsx", "應該處理特殊字元"),
        ("假Excel.xlsx", "應該偵測到假檔案"),
        ("不存在.xlsx", "應該提示檔案不存在")
    ]
    
    for filename, expected in test_files:
        file_path = os.path.join(test_dir, filename)
        print(f"\n測試: {filename} - {expected}")
        
        try:
            start_time = time.time()
            
            if not os.path.exists(file_path):
                print(f"  ❌ 檔案不存在: {file_path}")
                continue
            
            # 檢查檔案大小
            file_size = os.path.getsize(file_path)
            if file_size > 100 * 1024 * 1024:  # 100MB
                print(f"  ⚠️ 大檔案警告: {file_size/1024/1024:.1f}MB")
            
            # 嘗試讀取
            df = pd.read_excel(file_path, engine='openpyxl')
            
            elapsed = time.time() - start_time
            print(f"  ✓ 讀取成功: {df.shape[0]}行 x {df.shape[1]}欄 ({elapsed:.2f}秒)")
            
            # 檢查是否為空
            if df.empty:
                print(f"  ⚠️ 檔案為空")
            
            # 檢查是否有Reviewer欄
            if 'Reviewer' not in df.columns:
                print(f"  ⚠️ 找不到Reviewer欄位")
                print(f"    可用欄位: {', '.join(df.columns)}")
            else:
                reviewers = df['Reviewer'].dropna().unique()
                print(f"  ✓ 找到 {len(reviewers)} 個評審員")
                
                # 檢查特殊字元
                special_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|', '#']
                problematic_reviewers = []
                for reviewer in reviewers:
                    reviewer_str = str(reviewer)
                    if any(char in reviewer_str for char in special_chars):
                        problematic_reviewers.append(reviewer_str)
                
                if problematic_reviewers:
                    print(f"  ⚠️ 有 {len(problematic_reviewers)} 個評審員名稱含特殊字元")
                    for name in problematic_reviewers[:3]:  # 只顯示前3個
                        print(f"    - {name}")
            
        except Exception as e:
            print(f"  ❌ 讀取失敗: {type(e).__name__}: {str(e)[:100]}")

def test_memory_usage_monitoring():
    """測試記憶體使用監控"""
    print("\n\n測試記憶體使用監控...")
    print("=" * 50)
    
    tracemalloc.start()
    
    try:
        # 模擬處理大量評審員
        print("模擬建立大量評審員資料...")
        
        reviewer_data = {}
        for i in range(1000):
            reviewer_name = f"評審員_{i}"
            reviewer_data[reviewer_name] = {
                'email': f"reviewer{i}@company.com",
                'selected': True,
                'status': 'ready',
                'data': [f"資料_{j}" for j in range(100)]  # 模擬一些資料
            }
        
        current, peak = tracemalloc.get_traced_memory()
        print(f"當前記憶體使用: {current / 1024 / 1024:.1f}MB")
        print(f"高峰記憶體使用: {peak / 1024 / 1024:.1f}MB")
        
        # 清理資料
        del reviewer_data
        
        current_after, _ = tracemalloc.get_traced_memory()
        print(f"清理後記憶體使用: {current_after / 1024 / 1024:.1f}MB")
        
        if current_after < current * 0.8:
            print("✓ 記憶體清理成功")
        else:
            print("⚠️ 可能有記憶體洩漏")
            
    finally:
        tracemalloc.stop()

def test_folder_name_sanitization():
    """測試資料夾名稱清理"""
    print("\n\n測試資料夾名稱清理...")
    print("=" * 50)
    
    problematic_names = [
        "正常名稱",
        "John O'Brien",
        "María García-López",
        "Smith & Jones Co.",
        "Test/User\\Name",
        "File:Name*With?Special\"Chars<>|",
        "超級長的名稱" * 20,  # 超過255字元
        "   前後有空格   ",
        "#Hashtag%Name",
        "李明 (Li Ming) 測試"
    ]
    
    for name in problematic_names:
        print(f"\n原始名稱: '{name}' (長度: {len(name)})")
        
        # 模擬清理函數
        sanitized = sanitize_folder_name_test(name)
        print(f"清理後: '{sanitized}' (長度: {len(sanitized)})")
        
        # 檢查是否適合作為資料夾名稱
        if len(sanitized) > 255:
            print("  ❌ 名稱太長")
        elif not sanitized.strip():
            print("  ❌ 清理後變成空白")
        else:
            print("  ✓ 適合作為資料夾名稱")

def sanitize_folder_name_test(name):
    """測試用的資料夾名稱清理函數"""
    replacements = {
        '/': '_', '\\': '_', ':': '_', '*': '_', '?': '_',
        '"': '_', '<': '_', '>': '_', '|': '_', '#': '_', '%': '_'
    }
    
    sanitized = name.strip()
    for char, replacement in replacements.items():
        sanitized = sanitized.replace(char, replacement)
    
    if len(sanitized) > 255:
        sanitized = sanitized[:255].rstrip()
    
    return sanitized

def test_api_timeout_simulation():
    """模擬API逾時情況"""
    print("\n\n測試API逾時模擬...")
    print("=" * 50)
    
    timeout_scenarios = [
        (0.1, "正常回應"),
        (1.0, "稍慢回應"),
        (5.0, "慢速回應"),
        (30.0, "極慢回應"),
        (60.0, "超逾時回應")
    ]
    
    for delay, desc in timeout_scenarios:
        print(f"\n模擬 {desc} (延遲 {delay}秒)")
        
        start_time = time.time()
        try:
            # 模擬API呼叫
            simulate_api_call(delay, timeout=10.0)
            elapsed = time.time() - start_time
            print(f"  ✓ API呼叫成功 ({elapsed:.2f}秒)")
            
        except TimeoutError:
            elapsed = time.time() - start_time
            print(f"  ❌ API逾時 ({elapsed:.2f}秒)")
        except Exception as e:
            elapsed = time.time() - start_time
            print(f"  ❌ API錯誤: {e} ({elapsed:.2f}秒)")

def simulate_api_call(delay, timeout=30.0):
    """模擬API呼叫"""
    if delay > timeout:
        raise TimeoutError(f"逾時: {delay}秒 > {timeout}秒")
    
    # 模擬處理時間
    time.sleep(min(delay, 0.1))  # 實際只等待短時間避免測試太慢
    
    if delay > 30:
        raise Exception("模擬API錯誤")
    
    return {"status": "success", "delay": delay}

def run_real_world_tests():
    """執行所有實戰測試"""
    print("🔥 實戰痛點測試開始")
    print("=" * 70)
    print("這些都是真實環境會踩到的地雷！\n")
    
    tests = [
        test_file_reading_robustness,
        test_memory_usage_monitoring,
        test_folder_name_sanitization,
        test_api_timeout_simulation
    ]
    
    for test in tests:
        try:
            test()
        except Exception as e:
            print(f"\n💥 {test.__name__} 測試爆炸: {e}")
            print("Stack trace:")
            import traceback
            traceback.print_exc()
    
    print("\n\n" + "=" * 70)
    print("🎯 實戰測試完成")
    print("=" * 70)
    print("\n現在你的SharePoint整合應該更能應付現實世界的各種鳥事了！")
    print("記住：程式碼能跑不代表能在生產環境存活 😅")

if __name__ == "__main__":
    run_real_world_tests()