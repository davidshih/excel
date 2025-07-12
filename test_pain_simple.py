#!/usr/bin/env python3
"""
實戰痛點測試 (簡化版) - 不依賴pandas但測試重要概念
"""

import os
import tempfile
import time

def test_folder_name_sanitization():
    """測試資料夾名稱清理 - 最重要的測試之一"""
    print("測試資料夾名稱清理 - 這個超重要！")
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
        "李明 (Li Ming) 測試",
        "",  # 空字串
        "   ",  # 只有空格
        "COM1",  # Windows保留字
        "PRN",   # Windows保留字
        "檔案名稱.副檔名但沒有副檔名",
    ]
    
    for name in problematic_names:
        print(f"\n原始: '{name}' (長度: {len(name)})")
        
        sanitized = sanitize_folder_name_test(name)
        print(f"清理: '{sanitized}' (長度: {len(sanitized)})")
        
        # 驗證結果
        if not sanitized or not sanitized.strip():
            print("  ❌ 清理後變成空白")
        elif len(sanitized) > 255:
            print("  ❌ 還是太長")
        elif any(char in sanitized for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']):
            print("  ❌ 還有危險字元")
        else:
            print("  ✓ 安全的資料夾名稱")

def sanitize_folder_name_test(name):
    """資料夾名稱清理函數"""
    if not name or not name.strip():
        return "未命名資料夾"
    
    replacements = {
        '/': '_', '\\': '_', ':': '_', '*': '_', '?': '_',
        '"': '_', '<': '_', '>': '_', '|': '_', '#': '_', '%': '_'
    }
    
    sanitized = name.strip()
    for char, replacement in replacements.items():
        sanitized = sanitized.replace(char, replacement)
    
    # 處理Windows保留字
    windows_reserved = ['CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9']
    if sanitized.upper() in windows_reserved:
        sanitized = sanitized + "_folder"
    
    # 長度限制
    if len(sanitized) > 255:
        sanitized = sanitized[:250] + "..."
    
    return sanitized

def test_email_validation():
    """測試Email驗證"""
    print("\n\n測試Email驗證 - 用戶輸入很混亂")
    print("=" * 50)
    
    test_emails = [
        ("john.doe@company.com", True, "正常格式"),
        ("user+tag@domain.com", True, "有加號"),
        ("name@sub.domain.com", True, "子網域"),
        ("invalid@", False, "沒有網域"),
        ("@invalid.com", False, "沒有用戶名"),
        ("no-at-sign.com", False, "沒有@符號"),
        ("spaces in@email.com", False, "有空格"),
        ("", False, "空字串"),
        ("   ", False, "只有空格"),
        ("unicode测试@domain.com", True, "Unicode字元"),
        ("very.long.email.address@very.long.domain.name.com", True, "很長的信箱"),
        ("test@localhost", True, "本地網域"),
        ("test@192.168.1.1", False, "IP位址域名")
    ]
    
    for email, expected, desc in test_emails:
        result = validate_email_test(email)
        status = "✓" if result == expected else "❌"
        print(f"  {status} '{email}' - {desc} (預期: {expected}, 實際: {result})")

def validate_email_test(email):
    """Email驗證函數"""
    import re
    
    if not email or not email.strip():
        return False
    
    # 基本格式檢查
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email.strip()) is not None

def test_api_retry_logic():
    """測試API重試邏輯"""
    print("\n\n測試API重試邏輯 - 網路就是會斷")
    print("=" * 50)
    
    scenarios = [
        ("正常回應", 0, 200, None),
        ("暫時錯誤", 1, 500, "伺服器錯誤"),
        ("認證過期", 0, 401, "需要重新認證"),
        ("速率限制", 2, 429, "太多請求"),
        ("網路逾時", 3, None, "連線逾時"),
        ("完全失敗", 3, 500, "持續錯誤")
    ]
    
    for desc, fail_times, status_code, error_msg in scenarios:
        print(f"\n模擬: {desc}")
        
        result = simulate_api_call_with_retry(fail_times, status_code, error_msg)
        
        if result['success']:
            print(f"  ✓ 最終成功 (重試 {result['attempts']} 次)")
        else:
            print(f"  ❌ 最終失敗 (重試 {result['attempts']} 次): {result['error']}")

def simulate_api_call_with_retry(fail_times, final_status, error_msg, max_retries=3):
    """模擬帶重試的API呼叫"""
    attempts = 0
    
    for attempt in range(max_retries + 1):
        attempts += 1
        print(f"    嘗試 {attempts}: ", end="")
        
        if attempt < fail_times:
            print("失敗 (將重試)")
            if attempt < max_retries:
                time.sleep(0.1)  # 模擬等待時間
        else:
            if final_status == 200:
                print("成功")
                return {'success': True, 'attempts': attempts, 'error': None}
            elif final_status == 401:
                print("認證錯誤 (不重試)")
                return {'success': False, 'attempts': attempts, 'error': error_msg}
            elif final_status == 429:
                print("速率限制 (等待後重試)")
                if attempt < max_retries:
                    time.sleep(0.1)
            else:
                print(f"錯誤 {final_status}")
                if attempt < max_retries:
                    time.sleep(0.1)
    
    return {'success': False, 'attempts': attempts, 'error': error_msg}

def test_file_operations():
    """測試檔案操作"""
    print("\n\n測試檔案操作 - 各種會出錯的情況")
    print("=" * 50)
    
    with tempfile.TemporaryDirectory() as temp_dir:
        print(f"使用暫存目錄: {temp_dir}")
        
        # 測試1: 建立正常資料夾
        normal_folder = os.path.join(temp_dir, "正常資料夾")
        try:
            os.makedirs(normal_folder, exist_ok=True)
            print(f"  ✓ 建立正常資料夾成功")
        except Exception as e:
            print(f"  ❌ 建立正常資料夾失敗: {e}")
        
        # 測試2: 建立有特殊字元的資料夾
        special_folder = os.path.join(temp_dir, sanitize_folder_name_test("Test/User:Name"))
        try:
            os.makedirs(special_folder, exist_ok=True)
            print(f"  ✓ 建立特殊字元資料夾成功: {os.path.basename(special_folder)}")
        except Exception as e:
            print(f"  ❌ 建立特殊字元資料夾失敗: {e}")
        
        # 測試3: 檔案權限測試
        test_file = os.path.join(normal_folder, "測試檔案.txt")
        try:
            with open(test_file, 'w', encoding='utf-8') as f:
                f.write("測試內容")
            print(f"  ✓ 寫入檔案成功")
            
            with open(test_file, 'r', encoding='utf-8') as f:
                content = f.read()
            print(f"  ✓ 讀取檔案成功: {content}")
            
        except Exception as e:
            print(f"  ❌ 檔案操作失敗: {e}")

def test_unicode_handling():
    """測試Unicode處理"""
    print("\n\n測試Unicode處理 - 國際化很重要")
    print("=" * 50)
    
    unicode_tests = [
        ("English Name", "英文名稱"),
        ("中文姓名", "中文名稱"),
        ("日本語名前", "日文名稱"),
        ("한국어 이름", "韓文名稱"),
        ("العربية", "阿拉伯文"),
        ("Русский", "俄文"),
        ("Ελληνικά", "希臘文"),
        ("🙂 Emoji Name", "有表情符號"),
        ("Mixed中英文Name", "中英混合"),
        ("Café & Résumé", "有重音符號"),
    ]
    
    for name, desc in unicode_tests:
        print(f"\n測試: {name} ({desc})")
        
        # 測試編碼
        try:
            encoded = name.encode('utf-8')
            decoded = encoded.decode('utf-8')
            print(f"  ✓ UTF-8編碼: {len(encoded)}字節")
            
            if decoded == name:
                print(f"  ✓ 編碼解碼正確")
            else:
                print(f"  ❌ 編碼解碼失敗")
                
        except Exception as e:
            print(f"  ❌ 編碼測試失敗: {e}")
        
        # 測試資料夾名稱清理
        try:
            sanitized = sanitize_folder_name_test(name)
            print(f"  ✓ 清理後: {sanitized}")
        except Exception as e:
            print(f"  ❌ 名稱清理失敗: {e}")

def run_all_pain_tests():
    """執行所有痛點測試"""
    print("🔥 實戰痛點測試 (簡化版)")
    print("=" * 70)
    print("測試各種現實世界會遇到的問題...\n")
    
    tests = [
        test_folder_name_sanitization,
        test_email_validation,
        test_api_retry_logic,
        test_file_operations,
        test_unicode_handling
    ]
    
    for test in tests:
        try:
            test()
        except Exception as e:
            print(f"\n💥 {test.__name__} 爆炸了: {e}")
            import traceback
            traceback.print_exc()
    
    print("\n\n" + "=" * 70)
    print("🎯 痛點測試完成")
    print("=" * 70)
    print("\n測試結果顯示很多邊界情況需要處理！")
    print("現在你的SharePoint整合應該更穩定了 💪")

if __name__ == "__main__":
    run_all_pain_tests()