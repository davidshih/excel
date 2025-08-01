#!/usr/bin/env python3
"""
完整的 Excel COM 解決方案
解決所有已知問題：COM 物件釋放、工作表遺失、資料驗證

整合功能：
1. 正確的 COM 物件管理
2. 完整的工作表複製
3. 智慧篩選（只對資料工作表）
4. 資料驗證保護
5. 詳細的診斷和日誌
"""

import os
import sys
import time
import gc
import shutil
from typing import Dict, List, Tuple, Optional, Set
from datetime import datetime
import traceback

# Windows 檢查
if sys.platform != 'win32':
    print("❌ 此程式僅支援 Windows 系統")
    exit(1)

try:
    import win32com.client
    import pywintypes
    from win32com.client import constants
except ImportError:
    print("❌ 請先安裝 pywin32: pip install pywin32")
    exit(1)


class ExcelCOMManager:
    """完整的 Excel COM 管理器"""
    
    def __init__(self, visible=False, enable_logging=True):
        self.excel = None
        self.workbooks = []
        self.visible = visible
        self.enable_logging = enable_logging
        self.log_file = None
        
        if self.enable_logging:
            log_filename = f"excel_com_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            self.log_file = open(log_filename, 'w', encoding='utf-8')
            self.log(f"Excel COM 管理器啟動 - 日誌檔案: {log_filename}")
    
    def log(self, message: str):
        """記錄日誌"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        log_msg = f"[{timestamp}] {message}"
        print(log_msg)
        
        if self.log_file:
            self.log_file.write(log_msg + "\n")
            self.log_file.flush()
    
    def start_excel(self) -> bool:
        """啟動 Excel 應用程式"""
        try:
            # 確保清理舊的實例
            self.cleanup()
            
            self.log("🚀 啟動 Excel COM 應用程式...")
            self.excel = win32com.client.Dispatch("Excel.Application")
            
            # 設定 Excel 參數
            self.excel.Visible = self.visible
            self.excel.DisplayAlerts = False
            self.excel.ScreenUpdating = False
            self.excel.EnableEvents = False
            self.excel.Calculation = constants.xlCalculationManual
            
            self.log("✅ Excel COM 應用程式啟動成功")
            return True
            
        except Exception as e:
            self.log(f"❌ 啟動 Excel 失敗: {e}")
            return False
    
    def analyze_workbook_structure(self, file_path: str) -> Dict:
        """分析工作簿結構"""
        self.log(f"🔍 分析工作簿結構: {os.path.basename(file_path)}")
        
        wb = None
        try:
            wb = self.excel.Workbooks.Open(os.path.abspath(file_path))
            self.workbooks.append(wb)
            
            structure = {
                'total_sheets': wb.Worksheets.Count,
                'sheet_names': [],
                'main_data_sheet': None,
                'validation_sheets': [],
                'hidden_sheets': []
            }
            
            # 分析每個工作表
            for i in range(1, wb.Worksheets.Count + 1):
                ws = wb.Worksheets(i)
                sheet_info = {
                    'name': ws.Name,
                    'visible': ws.Visible,
                    'used_range': ws.UsedRange.Address if ws.UsedRange else None,
                    'row_count': ws.UsedRange.Rows.Count if ws.UsedRange else 0,
                    'has_data_validation': False
                }
                
                # 檢查是否有資料驗證
                try:
                    if ws.UsedRange:
                        # 簡單檢查是否有驗證（這個檢查可能很慢）
                        sample_range = ws.Range(ws.Cells(1, 1), ws.Cells(min(10, ws.UsedRange.Rows.Count), ws.UsedRange.Columns.Count))
                        for cell in sample_range:
                            if cell.Validation.Type != constants.xlValidateInputOnly:
                                sheet_info['has_data_validation'] = True
                                break
                except:
                    pass  # 忽略驗證檢查錯誤
                
                structure['sheet_names'].append(sheet_info)
                
                # 判斷工作表類型
                if ws.Visible == constants.xlSheetHidden:
                    structure['hidden_sheets'].append(ws.Name)
                elif sheet_info['row_count'] > 1 and not structure['main_data_sheet']:
                    structure['main_data_sheet'] = ws.Name
                elif sheet_info['has_data_validation'] and sheet_info['row_count'] > 0:
                    structure['validation_sheets'].append(ws.Name)
            
            self.log(f"📊 工作簿結構分析完成:")
            self.log(f"   - 總工作表數: {structure['total_sheets']}")
            self.log(f"   - 主資料工作表: {structure['main_data_sheet']}")
            self.log(f"   - 驗證資料工作表: {', '.join(structure['validation_sheets'])}")
            self.log(f"   - 隱藏工作表: {', '.join(structure['hidden_sheets'])}")
            
            return structure
            
        except Exception as e:
            self.log(f"❌ 分析工作簿結構失敗: {e}")
            return {}
        
        finally:
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                    if wb in self.workbooks:
                        self.workbooks.remove(wb)
                except:
                    pass
    
    def process_reviewer_complete(self, file_path: str, reviewer: str, 
                                column_name: str, output_folder: str,
                                structure_info: Dict = None) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        完整的審查者處理流程
        
        關鍵改進：
        1. 完整複製所有工作表
        2. 智慧識別主資料工作表
        3. 只對主資料工作表套用篩選
        4. 保護資料驗證工作表
        """
        
        reviewer_clean = self.sanitize_folder_name(str(reviewer).strip())
        self.log(f"\n📝 開始處理審查者: {reviewer} -> {reviewer_clean}")
        
        # 建立輸出資料夾
        reviewer_folder = os.path.join(output_folder, reviewer_clean)
        os.makedirs(reviewer_folder, exist_ok=True)
        
        # 生成檔案名稱
        base_name = os.path.basename(file_path)
        name_without_ext = os.path.splitext(base_name)[0]
        new_filename = f"{name_without_ext} - {reviewer_clean}.xlsx"
        dst_path = os.path.join(reviewer_folder, new_filename)
        
        wb_source = None
        wb_dest = None
        
        try:
            # 步驟 1: 開啟來源檔案
            self.log("  📖 開啟來源檔案...")
            wb_source = self.excel.Workbooks.Open(os.path.abspath(file_path))
            self.workbooks.append(wb_source)
            
            # 步驟 2: 使用 SaveCopyAs 複製整個工作簿
            self.log("  📋 複製整個工作簿（保留所有工作表）...")
            wb_source.SaveCopyAs(os.path.abspath(dst_path))
            
            # 關閉來源檔案
            wb_source.Close(SaveChanges=False)
            self.workbooks.remove(wb_source)
            wb_source = None
            
            # 步驟 3: 開啟複製的檔案
            self.log("  📂 開啟複製的檔案...")
            wb_dest = self.excel.Workbooks.Open(os.path.abspath(dst_path))
            self.workbooks.append(wb_dest)
            
            # 步驟 4: 識別主資料工作表
            main_sheet_name = None
            if structure_info and structure_info.get('main_data_sheet'):
                main_sheet_name = structure_info['main_data_sheet']
            else:
                # 預設使用第一個可見工作表
                for i in range(1, wb_dest.Worksheets.Count + 1):
                    ws = wb_dest.Worksheets(i)
                    if ws.Visible == constants.xlSheetVisible:
                        main_sheet_name = ws.Name
                        break
            
            if not main_sheet_name:
                raise ValueError("找不到主資料工作表")
            
            self.log(f"  🎯 識別主資料工作表: {main_sheet_name}")
            
            # 步驟 5: 尋找欄位
            main_ws = wb_dest.Worksheets(main_sheet_name)
            col_idx = self.find_column_com(main_ws, column_name)
            if not col_idx:
                raise ValueError(f"在工作表 '{main_sheet_name}' 中找不到欄位 '{column_name}'")
            
            self.log(f"  🔍 找到欄位 '{column_name}' 在第 {col_idx} 欄")
            
            # 步驟 6: 套用篩選（只對主資料工作表）
            self.log(f"  🎯 對主資料工作表套用篩選...")
            self.apply_smart_filter(main_ws, col_idx, str(reviewer))
            
            # 步驟 7: 檢查其他工作表（記錄但不修改）
            self.log(f"  📄 檢查其他工作表狀態:")
            for i in range(1, wb_dest.Worksheets.Count + 1):
                ws = wb_dest.Worksheets(i)
                if ws.Name != main_sheet_name:
                    visibility = "可見" if ws.Visible == constants.xlSheetVisible else "隱藏"
                    row_count = ws.UsedRange.Rows.Count if ws.UsedRange else 0
                    self.log(f"    - {ws.Name}: {visibility}, {row_count} 行")
            
            # 步驟 8: 儲存檔案
            self.log("  💾 儲存處理後的檔案...")
            wb_dest.Save()
            wb_dest.Close()
            self.workbooks.remove(wb_dest)
            wb_dest = None
            
            self.log(f"  ✅ 審查者 {reviewer} 處理完成")
            return True, reviewer_folder, new_filename
            
        except Exception as e:
            self.log(f"  ❌ 處理審查者 {reviewer} 失敗: {e}")
            self.log(f"  📋 錯誤詳情: {traceback.format_exc()}")
            
            # 清理工作簿
            for wb in [wb_source, wb_dest]:
                if wb:
                    try:
                        wb.Close(SaveChanges=False)
                        if wb in self.workbooks:
                            self.workbooks.remove(wb)
                    except:
                        pass
            
            return False, None, None
    
    def find_column_com(self, worksheet, column_name: str) -> Optional[int]:
        """使用 COM 在工作表中尋找欄位"""
        try:
            if not worksheet.UsedRange:
                return None
                
            first_row = worksheet.UsedRange.Rows(1)
            
            for col in range(1, first_row.Columns.Count + 1):
                cell_value = first_row.Cells(1, col).Value
                if cell_value and str(cell_value).strip() == column_name:
                    return col
            
            return None
            
        except Exception as e:
            self.log(f"    ⚠️ 尋找欄位錯誤: {e}")
            return None
    
    def apply_smart_filter(self, worksheet, col_idx: int, reviewer: str):
        """智慧套用篩選"""
        try:
            # 清除現有篩選
            worksheet.AutoFilterMode = False
            
            # 確保有資料範圍
            if not worksheet.UsedRange:
                self.log("    ⚠️ 工作表沒有資料範圍")
                return
            
            # 套用自動篩選
            used_range = worksheet.UsedRange
            used_range.AutoFilter(Field=col_idx, Criteria1=reviewer)
            
            # 計算篩選後的行數
            visible_rows = 0
            for row in range(2, used_range.Rows.Count + 1):  # 跳過標題行
                if not worksheet.Rows(row).Hidden:
                    visible_rows += 1
            
            self.log(f"    ✅ 篩選套用成功，顯示 {visible_rows} 行資料")
            
        except Exception as e:
            self.log(f"    ❌ 套用篩選失敗: {e}")
            raise
    
    def sanitize_folder_name(self, name: str) -> str:
        """清理資料夾名稱"""
        invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|', '#', '%']
        sanitized = name.strip()
        
        for char in invalid_chars:
            sanitized = sanitized.replace(char, '_')
        
        if len(sanitized) > 255:
            sanitized = sanitized[:255].rstrip()
        
        return sanitized
    
    def process_multiple_reviewers_complete(self, file_path: str, reviewers: List[str], 
                                          column_name: str, output_folder: str) -> Tuple[int, int]:
        """
        完整的多審查者處理流程
        
        特點：
        1. 預先分析工作簿結構
        2. 重複使用 Excel 實例
        3. 智慧錯誤恢復
        4. 詳細進度追蹤
        """
        
        self.log(f"\n🚀 開始處理多個審查者")
        self.log(f"📁 來源檔案: {file_path}")
        self.log(f"📊 審查者數量: {len(reviewers)}")
        self.log(f"📋 審查者欄位: {column_name}")
        self.log(f"📂 輸出資料夾: {output_folder}")
        
        # 啟動 Excel
        if not self.start_excel():
            return 0, len(reviewers)
        
        # 預先分析工作簿結構
        structure_info = self.analyze_workbook_structure(file_path)
        
        processed = 0
        failed = 0
        
        try:
            for i, reviewer in enumerate(reviewers):
                self.log(f"\n📝 進度: {i+1}/{len(reviewers)}")
                
                success, folder_path, filename = self.process_reviewer_complete(
                    file_path, reviewer, column_name, output_folder, structure_info
                )
                
                if success:
                    processed += 1
                    self.log(f"✅ 成功: {reviewer}")
                else:
                    failed += 1
                    self.log(f"❌ 失敗: {reviewer}")
                
                # 定期清理和垃圾收集
                if (i + 1) % 3 == 0:
                    self.log("🧹 執行定期清理...")
                    self.partial_cleanup()
                    gc.collect()
                    time.sleep(0.5)
        
        except Exception as e:
            self.log(f"❌ 處理過程發生嚴重錯誤: {e}")
            self.log(f"📋 錯誤詳情: {traceback.format_exc()}")
        
        finally:
            # 完全清理
            self.cleanup()
        
        # 總結
        self.log(f"\n📊 處理總結:")
        self.log(f"  ✅ 成功: {processed}/{len(reviewers)}")
        self.log(f"  ❌ 失敗: {failed}")
        self.log(f"  📁 輸出位置: {output_folder}")
        
        return processed, failed
    
    def partial_cleanup(self):
        """部分清理：關閉工作簿但保留 Excel 實例"""
        try:
            # 關閉所有追蹤的工作簿
            for wb in self.workbooks[:]:
                try:
                    wb.Close(SaveChanges=False)
                    self.workbooks.remove(wb)
                except:
                    pass
            
            # 垃圾收集
            gc.collect()
            
        except Exception as e:
            self.log(f"⚠️ 部分清理警告: {e}")
    
    def cleanup(self):
        """完全清理 Excel COM 物件"""
        self.log("🧹 開始完全清理 Excel COM 物件...")
        
        try:
            # 關閉所有工作簿
            for wb in self.workbooks[:]:
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass
            self.workbooks.clear()
            
            # 恢復 Excel 設定並關閉
            if self.excel:
                try:
                    self.excel.ScreenUpdating = True
                    self.excel.EnableEvents = True
                    self.excel.Calculation = constants.xlCalculationAutomatic
                    self.excel.DisplayAlerts = True
                    self.excel.Quit()
                except:
                    pass
                finally:
                    self.excel = None
            
            # 強制垃圾收集
            gc.collect()
            time.sleep(1)
            
            self.log("✅ Excel COM 清理完成")
            
        except Exception as e:
            self.log(f"⚠️ 清理過程警告: {e}")
        
        # 關閉日誌檔案
        if self.log_file:
            self.log_file.close()
            self.log_file = None


def demo_complete_solution():
    """完整解決方案示範"""
    print("🎯 Excel COM 完整解決方案示範")
    print("=" * 60)
    
    # 示範用參數（請根據實際情況修改）
    file_path = r"C:\path\to\your\excel\file.xlsx"  # 請修改為實際路徑
    reviewers = ["張三", "李四", "王五", "趙六"]
    column_name = "Reviewer"
    output_folder = r"C:\path\to\output"  # 請修改為實際路徑
    
    # 檢查檔案是否存在
    if not os.path.exists(file_path):
        print(f"⚠️ 示範檔案不存在: {file_path}")
        print("📝 請修改 file_path 變數為實際的 Excel 檔案路徑")
        return
    
    # 建立管理器
    manager = ExcelCOMManager(visible=False, enable_logging=True)
    
    try:
        # 執行完整處理
        processed, failed = manager.process_multiple_reviewers_complete(
            file_path, reviewers, column_name, output_folder
        )
        
        print(f"\n🎉 處理完成！")
        print(f"✅ 成功處理: {processed} 個審查者")
        print(f"❌ 處理失敗: {failed} 個審查者")
        
        if processed > 0:
            print(f"\n💡 建議檢查項目：")
            print(f"1. 開啟生成的檔案確認可正常開啟")
            print(f"2. 檢查資料驗證下拉選單是否正常")
            print(f"3. 確認所有相關工作表都已保留")
            print(f"4. 驗證篩選條件是否正確套用")
        
    except Exception as e:
        print(f"❌ 示範執行失敗: {e}")
        traceback.print_exc()
    
    finally:
        # 確保清理
        manager.cleanup()


if __name__ == "__main__":
    demo_complete_solution()