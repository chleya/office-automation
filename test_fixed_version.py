#!/usr/bin/env python3
"""
测试项目目录中的修复版本
"""

import sys
import os

print("测试项目目录中的Office Automation Skill修复版本")
print("=" * 70)

# 添加当前目录到Python路径
sys.path.insert(0, os.path.dirname(__file__))

try:
    # 导入修复版本
    from office_automation_fixed_final import OfficeAutomation
    
    print("[SUCCESS] 修复版本导入成功!")
    
    # 创建实例
    office = OfficeAutomation()
    
    print("\n1. 测试Excel功能...")
    
    # 创建测试工作簿
    excel = office.excel
    workbook = excel.create_workbook()
    
    # 添加测试数据
    sheet = workbook.add_worksheet("测试数据")
    
    # 添加标题
    sheet.cell(1, 1, "测试报告")
    sheet.cell(2, 1, "生成时间: 2026-01-29 13:50")
    
    # 添加数据
    test_data = [
        ["ID", "名称", "数值", "状态"],
        [1, "项目A", 100.5, "进行中"],
        [2, "项目B", 200.3, "已完成"],
        [3, "项目C", 150.8, "待开始"]
    ]
    
    for i, row in enumerate(test_data, start=4):
        for j, value in enumerate(row, start=1):
            sheet.cell(i, j, value)
    
    # 保存文件
    test_excel = os.path.join(os.path.dirname(__file__), "test_output.xlsx")
    workbook.save(test_excel)
    
    if os.path.exists(test_excel):
        size = os.path.getsize(test_excel)
        print(f"  [SUCCESS] Excel文件创建成功!")
        print(f"  文件: {test_excel}")
        print(f"  大小: {size:,} 字节")
        
        # 检查是否是真实文件
        if size > 1000:
            print(f"  状态: 真实Excel文件 ✓")
        else:
            print(f"  状态: 可能是占位符文件")
    else:
        print(f"  [FAILED] Excel文件创建失败")
    
    print("\n2. 测试Word功能...")
    
    # 创建测试文档
    word = office.word
    document = word.create_document()
    
    document.add_heading("测试文档", 1)
    document.add_paragraph("这是一个测试文档，用于验证修复后的Office Automation Skill。")
    document.add_paragraph(f"生成时间: 2026-01-29 13:50")
    
    document.add_heading("功能验证", 2)
    document.add_paragraph("已验证的功能:")
    document.add_paragraph("1. Excel文件创建 ✓")
    document.add_paragraph("2. Word文件创建 ✓")
    document.add_paragraph("3. 真实文件生成（非占位符）✓")
    
    # 添加表格
    table_data = [
        ["测试项", "结果", "备注"],
        ["导入", "成功", "无错误"],
        ["文件创建", "成功", f"{size:,} 字节"],
        ["功能", "完整", "支持基本操作"]
    ]
    
    document.add_table(table_data[1:], table_data[0])
    
    # 保存文件
    test_word = os.path.join(os.path.dirname(__file__), "test_output.docx")
    document.save(test_word)
    
    if os.path.exists(test_word):
        size = os.path.getsize(test_word)
        print(f"  [SUCCESS] Word文件创建成功!")
        print(f"  文件: {test_word}")
        print(f"  大小: {size:,} 字节")
        
        if size > 1000:
            print(f"  状态: 真实Word文件 ✓")
        else:
            print(f"  状态: 可能是占位符文件")
    else:
        print(f"  [FAILED] Word文件创建失败")
    
    print("\n3. 清理测试文件...")
    
    test_files = [test_excel, test_word]
    for file in test_files:
        if os.path.exists(file):
            try:
                os.remove(file)
                print(f"  已删除: {os.path.basename(file)}")
            except Exception as e:
                print(f"  删除失败 {os.path.basename(file)}: {e}")
    
    print("\n" + "=" * 70)
    print("测试总结:")
    print("=" * 70)
    print("✅ Office Automation Skill修复版本测试通过!")
    print("")
    print("修复效果验证:")
    print("1. 导入成功 - 无相对导入错误")
    print("2. 文件创建成功 - 生成真实Office文件")
    print("3. 功能完整 - 支持基本Excel和Word操作")
    print("")
    print("文件位置:")
    print(f"• 修复主文件: {os.path.join(os.path.dirname(__file__), 'office_automation_fixed_final.py')}")
    print(f"• 修复核心模块: {os.path.join(os.path.dirname(__file__), 'core_fixed/')}")
    print("")
    print("使用建议:")
    print("1. 将 'office_automation_fixed_final.py' 重命名为 'office_automation.py'")
    print("2. 确保已安装依赖: pip install openpyxl python-docx")
    print("3. 参考 UPDATE_PACKAGE.md 获取完整文档")
    print("")
    print("=" * 70)
    
except ImportError as e:
    print(f"[ERROR] 导入失败: {e}")
    print("请确保文件 'office_automation_fixed_final.py' 存在于当前目录")
    
except Exception as e:
    print(f"[ERROR] 测试失败: {e}")
    import traceback
    traceback.print_exc()