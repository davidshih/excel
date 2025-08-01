#!/usr/bin/env python3
"""
修正版 Excel COM 處理程式
解決 COM 物件釋放和工作表遺失問題

主要修正：
1. 正確釋放 COM 物件
2. 複製所有工作表（包含資料驗證來源）
3. 僅對主工作表套用篩選
4. 加強錯誤處理
"""

import os
import sys
import time
import gc
from typing import Optional, Tuple, List

# Windows 檢查
if sys.platform != 'win32':
    print("❌ 此程式僅支援 Windows 系統")
    sys.exit(1)

try:
    import win32com.client
    import pywintypes
except ImportError:
    print("❌ 請先安裝 pywin32: pip install pywin32")
    sys.exit(1)


class ExcelCOMProcessor:
    """Excel COM 處理器 - 修正版"""
    
    def __init__(self):
        self.excel = None
        self.workbooks = []  # 追蹤開啟的工作簿
    
    def start_excel(self, visible=False):
        """啟動 Excel 應用程式"""
        try:
            # 確保沒有殘留的 Excel 程序
            self.cleanup()
            
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = visible
            self.excel.DisplayAlerts = False
            self.excel.ScreenUpdating = False  # 提升效能
            self.excel.Calculation = -4135  # xlCalculationManual
            
            print("✓ Excel COM 應用程式已啟動")
            return True
            
        except Exception as e:
            print(f"❌ 啟動 Excel 失敗: {e}")
            return False
    
    def process_reviewer_excel_com_fixed(self, file_path: str, reviewer: str, 
                                       column_name: str, output_folder: str) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        修正版 COM 處理函數
        
        主要修正：
        1. 複製所有工作表
        2. 僅對主要工作表套用篩選
        3. 正確處理資料驗證來源工作表
        """
        if not self.excel:
            print("❌ Excel 應用程式未啟動")
            return False, None, None
            
        wb_source = None
        wb_dest = None
        
        try:
            reviewer_name = self.sanitize_folder_name(str(reviewer).strip())
            reviewer_folder = os.path.join(output_folder, reviewer_name)
            os.makedirs(reviewer_folder, exist_ok=True)
            
            base_name = os.path.basename(file_path)
            name_without_ext = os.path.splitext(base_name)[0]
            new_filename = f"{name_without_ext} - {reviewer_name}.xlsx"
            dst_path = os.path.join(reviewer_folder, new_filename)
            
            # 開啟來源檔案
            print(f"  📖 開啟來源檔案...")
            wb_source = self.excel.Workbooks.Open(os.path.abspath(file_path))
            self.workbooks.append(wb_source)
            
            # 🔧 關鍵修正：複製整個工作簿（包含所有工作表）
            print(f"  📋 複製整個工作簿...")
            wb_source.SaveCopyAs(os.path.abspath(dst_path))
            
            # 關閉來源檔案
            wb_source.Close(SaveChanges=False)
            self.workbooks.remove(wb_source)
            wb_source = None
            
            # 開啟複製的檔案進行處理
            print(f"  📝 開啟複製檔案進行處理...")
            wb_dest = self.excel.Workbooks.Open(os.path.abspath(dst_path))
            self.workbooks.append(wb_dest)
            
            # 🔧 關鍵修正：只處理主要工作表（通常是第一個）
            main_ws = wb_dest.Worksheets(1)  # 或根據名稱選擇
            
            print(f"  🔍 尋找欄位 '{column_name}'...")
            col_idx = self.find_column_com(main_ws, column_name)
            if not col_idx:
                raise ValueError(f"找不到欄位 '{column_name}'")
            
            # 🔧 修正：使用 AutoFilter 隱藏非相關列（保留所有資料）
            print(f"  🎯 套用篩選條件...")
            self.apply_filter_com(main_ws, col_idx, str(reviewer))
            
            # 顯示工作簿資訊
            print(f"  📊 工作簿包含 {wb_dest.Worksheets.Count} 個工作表")
            for i in range(1, wb_dest.Worksheets.Count + 1):
                ws_name = wb_dest.Worksheets(i).Name
                print(f"    - 工作表 {i}: {ws_name}")
            
            # 儲存變更
            print(f"  💾 儲存檔案...")
            wb_dest.Save()
            wb_dest.Close()
            self.workbooks.remove(wb_dest)
            wb_dest = None
            
            print(f"  ✅ COM 處理完成: {new_filename}")
            return True, reviewer_folder, new_filename
            
        except Exception as e:
            print(f"  ❌ COM 處理失敗 {reviewer}: {str(e)}")
            
            # 清理工作簿
            if wb_source:
                try:
                    wb_source.Close(SaveChanges=False)
                    if wb_source in self.workbooks:
                        self.workbooks.remove(wb_source)
                except:
                    pass
                    
            if wb_dest:
                try:
                    wb_dest.Close(SaveChanges=False)
                    if wb_dest in self.workbooks:
                        self.workbooks.remove(wb_dest)
                except:
                    pass
            
            return False, None, None
    
    def find_column_com(self, worksheet, column_name: str) -> Optional[int]:
        """在工作表中尋找欄位（COM 版本）"""
        try:
            used_range = worksheet.UsedRange
            first_row = used_range.Rows(1)
            
            for col in range(1, first_row.Columns.Count + 1):
                cell_value = first_row.Cells(1, col).Value
                if cell_value == column_name:
                    return col
            
            return None
            
        except Exception as e:
            print(f"    ⚠️ 尋找欄位時發生錯誤: {e}")
            return None
    
    def apply_filter_com(self, worksheet, col_idx: int, reviewer: str):
        """套用自動篩選（COM 版本）"""
        try:
            # 清除現有篩選
            worksheet.AutoFilterMode = False
            
            # 設定新的自動篩選
            used_range = worksheet.UsedRange
            used_range.AutoFilter(Field=col_idx, Criteria1=reviewer)
            
            print(f"    ✓ 已套用篩選條件: {reviewer}")
            
        except Exception as e:
            print(f"    ⚠️ 套用篩選時發生錯誤: {e}")
    
    def sanitize_folder_name(self, name: str) -> str:
        """清理資料夾名稱"""
        invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|', '#', '%']
        sanitized = name.strip()
        
        for char in invalid_chars:
            sanitized = sanitized.replace(char, '_')
        
        if len(sanitized) > 255:
            sanitized = sanitized[:255].rstrip()
        
        return sanitized
    
    def process_multiple_reviewers(self, file_path: str, reviewers: List[str], 
                                 column_name: str, output_folder: str) -> Tuple[int, int]:
        """
        處理多個審查者（修正版）
        
        關鍵修正：
        1. 重複使用同一個 Excel 實例
        2. 每次處理後進行清理
        3. 避免 COM 物件累積
        """
        if not self.start_excel():
            return 0, len(reviewers)
        
        processed = 0
        failed = 0
        
        try:
            for i, reviewer in enumerate(reviewers):
                print(f"\n📝 處理審查者 {i+1}/{len(reviewers)}: {reviewer}")
                
                success, folder_path, filename = self.process_reviewer_excel_com_fixed(
                    file_path, reviewer, column_name, output_folder
                )
                
                if success:
                    processed += 1
                else:
                    failed += 1
                
                # 🔧 關鍵修正：每次處理後進行清理
                self.partial_cleanup()
                
                # 每 5 個檔案後進行一次垃圾收集
                if (i + 1) % 5 == 0:
                    print("  🧹 執行垃圾收集...")
                    gc.collect()
                    time.sleep(0.5)  # 給系統一點時間
        
        finally:
            # 完全清理
            self.cleanup()
        
        return processed, failed
    
    def partial_cleanup(self):
        """部分清理：關閉工作簿但保留 Excel 實例"""
        # 關閉所有追蹤的工作簿
        for wb in self.workbooks[:]:  # 複製列表避免修改問題
            try:
                wb.Close(SaveChanges=False)
                self.workbooks.remove(wb)
            except:
                pass
        
        # 垃圾收集
        gc.collect()
    
    def cleanup(self):
        """完全清理 Excel COM 物件"""
        print("🧹 清理 Excel COM 物件...")
        
        # 關閉所有工作簿
        if self.workbooks:
            for wb in self.workbooks[:]:
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass
            self.workbooks.clear()
        
        # 關閉 Excel 應用程式
        if self.excel:
            try:
                # 恢復設定
                self.excel.ScreenUpdating = True
                self.excel.Calculation = -4105  # xlCalculationAutomatic
                self.excel.DisplayAlerts = True
                
                # 退出 Excel
                self.excel.Quit()
            except:
                pass
            finally:
                self.excel = None
        
        # 強制垃圾收集
        gc.collect()
        time.sleep(1)  # 給系統時間清理
        
        print("✓ Excel COM 清理完成")


def demo_usage():
    """使用範例"""
    print("📋 Excel COM 處理器使用範例")
    print("=" * 50)
    
    # 模擬參數
    file_path = "test.xlsx"  # 請替換為實際路徑
    reviewers = ["張三", "李四", "王五"]
    column_name = "Reviewer"
    output_folder = "output"
    
    # 建立處理器
    processor = ExcelCOMProcessor()
    
    try:
        # 處理多個審查者
        processed, failed = processor.process_multiple_reviewers(
            file_path, reviewers, column_name, output_folder
        )
        
        print(f"\n✅ 處理完成！")
        print(f"📊 成功: {processed}/{len(reviewers)}")
        print(f"❌ 失敗: {failed}")
        
    except Exception as e:
        print(f"❌ 處理過程發生錯誤: {e}")
    
    finally:
        # 確保清理
        processor.cleanup()


if __name__ == "__main__":
    demo_usage()