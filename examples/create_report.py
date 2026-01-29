#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Create Report Example

This example demonstrates how to create a complete business report
using the Office Automation skill. The report includes:
1. Cover page with title and metadata
2. Executive summary
3. Data analysis with tables
4. Charts and visualizations
5. Conclusions and recommendations
"""

import os
import sys
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from office_automation import OfficeAutomation


def create_business_report():
    """Create a complete business report with multiple sections."""
    print("=" * 60)
    print("Creating Business Report Example")
    print("=" * 60)
    
    # Initialize Office Automation
    print("\n1. Initializing Office Automation...")
    office = OfficeAutomation()
    print("   [OK] Office Automation initialized")
    
    # Create output directory
    output_dir = Path("output/reports")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        # ============================================
        # 1. Create Word Document Report
        # ============================================
        print("\n2. Creating Word document report...")
        
        # Create document
        doc = office.word.create_document()
        
        # Add cover page
        print("   Adding cover page...")
        doc.add_heading("季度业务报告", level=0)
        doc.add_paragraph("2024年第一季度")
        doc.add_paragraph("编制部门: 业务分析部")
        doc.add_paragraph(f"生成日期: 2024-01-28")
        doc.add_paragraph("")
        
        # Add table of contents
        print("   Adding table of contents...")
        doc.add_heading("目录", level=1)
        doc.add_paragraph("1. 执行摘要")
        doc.add_paragraph("2. 销售数据分析")
        doc.add_paragraph("3. 市场表现")
        doc.add_paragraph("4. 财务概况")
        doc.add_paragraph("5. 结论与建议")
        doc.add_paragraph("")
        
        # Add executive summary
        print("   Adding executive summary...")
        doc.add_heading("1. 执行摘要", level=1)
        doc.add_paragraph(
            "本季度公司整体表现良好，销售额同比增长15%，净利润增长12%。"
            "主要增长动力来自新产品线的推出和海外市场的拓展。"
        )
        doc.add_paragraph("")
        
        # Add sales data table
        print("   Adding sales data table...")
        doc.add_heading("2. 销售数据分析", level=1)
        
        sales_data = [
            ["产品线", "Q1销售额(万元)", "同比增长", "市场份额"],
            ["智能手机", "1250", "18%", "25%"],
            ["笔记本电脑", "890", "12%", "18%"],
            ["智能家居", "560", "25%", "15%"],
            ["配件", "320", "8%", "10%"],
            ["合计", "3020", "15%", "N/A"],
        ]
        
        doc.add_table(sales_data, style="Light Grid")
        doc.add_paragraph("")
        
        # Add market performance
        print("   Adding market performance...")
        doc.add_heading("3. 市场表现", level=1)
        doc.add_paragraph(
            "本季度公司在主要市场的表现如下："
        )
        
        market_data = [
            ["市场", "销售额(万元)", "增长率", "排名"],
            ["华北", "850", "20%", "1"],
            ["华东", "780", "15%", "2"],
            ["华南", "650", "18%", "3"],
            ["西部", "420", "22%", "4"],
            ["海外", "320", "30%", "N/A"],
        ]
        
        doc.add_table(market_data, style="Light Grid")
        doc.add_paragraph("")
        
        # Add financial overview
        print("   Adding financial overview...")
        doc.add_heading("4. 财务概况", level=1)
        
        financial_data = [
            ["指标", "Q1 2024", "Q4 2023", "变化"],
            ["营业收入", "3020", "2625", "+15%"],
            ["营业成本", "1812", "1650", "+10%"],
            ["毛利率", "40%", "37%", "+3%"],
            ["净利润", "423", "378", "+12%"],
            ["现金流", "580", "520", "+12%"],
        ]
        
        doc.add_table(financial_data, style="Light Grid")
        doc.add_paragraph("")
        
        # Add conclusions and recommendations
        print("   Adding conclusions and recommendations...")
        doc.add_heading("5. 结论与建议", level=1)
        
        conclusions = [
            "1. 销售额持续增长，但增速略有放缓",
            "2. 新产品线表现突出，贡献显著增长",
            "3. 海外市场拓展初见成效，潜力巨大",
            "4. 成本控制良好，利润率稳步提升",
        ]
        
        for conclusion in conclusions:
            doc.add_paragraph(conclusion)
        
        doc.add_paragraph("")
        doc.add_heading("建议措施", level=2)
        
        recommendations = [
            "1. 加大新产品研发投入，保持创新优势",
            "2. 深化海外市场布局，建立本地化团队",
            "3. 优化供应链管理，进一步降低成本",
            "4. 加强数字化转型，提升运营效率",
        ]
        
        for recommendation in recommendations:
            doc.add_paragraph(recommendation)
        
        # Save document
        report_path = output_dir / "business_report.docx"
        doc.save(report_path)
        print(f"   [OK] Word report saved: {report_path}")
        
        # ============================================
        # 2. Create Excel Data Analysis
        # ============================================
        print("\n3. Creating Excel data analysis...")
        
        # Create workbook
        workbook = office.excel.create_workbook()
        
        # Add sales data sheet
        print("   Adding sales data sheet...")
        sales_sheet = workbook.add_worksheet("销售数据")
        
        # Add headers
        headers = ["月份", "产品A", "产品B", "产品C", "合计", "增长率"]
        for col, header in enumerate(headers, start=1):
            sales_sheet.cell(row=1, column=col, value=header)
        
        # Add monthly data
        monthly_data = [
            ["1月", 120, 85, 65, 270, "N/A"],
            ["2月", 135, 92, 70, 297, "10%"],
            ["3月", 150, 105, 80, 335, "13%"],
            ["4月", 165, 115, 90, 370, "10%"],
            ["5月", 180, 125, 100, 405, "9%"],
            ["6月", 195, 135, 110, 440, "9%"],
        ]
        
        for row_idx, row_data in enumerate(monthly_data, start=2):
            for col_idx, cell_value in enumerate(row_data, start=1):
                sales_sheet.cell(row=row_idx, column=col_idx, value=cell_value)
        
        # Add formulas for totals
        for row in range(2, 8):  # Rows 2-7
            # Sum formula for column E (合计)
            sum_formula = f"=SUM(B{row}:D{row})"
            sales_sheet.cell(row=row, column=5, value=sum_formula)
            
            # Growth formula for column F (增长率)
            if row > 2:  # Skip first month
                growth_formula = f"=(E{row}-E{row-1})/E{row-1}"
                sales_sheet.cell(row=row, column=6, value=growth_formula)
        
        # Format headers
        for cell in sales_sheet[1]:
            cell.font = office.excel.Font(bold=True)
            cell.fill = office.excel.PatternFill(start_color="C6E0B4", 
                                                end_color="C6E0B4", 
                                                fill_type="solid")
        
        # Add market share sheet
        print("   Adding market share sheet...")
        market_sheet = workbook.add_worksheet("市场份额")
        
        market_data = [
            ["区域", "Q1 2023", "Q1 2024", "变化"],
            ["华北", "22%", "25%", "+3%"],
            ["华东", "20%", "22%", "+2%"],
            ["华南", "18%", "20%", "+2%"],
            ["西部", "15%", "18%", "+3%"],
            ["海外", "10%", "15%", "+5%"],
        ]
        
        for row_idx, row_data in enumerate(market_data, start=1):
            for col_idx, cell_value in enumerate(row_data, start=1):
                market_sheet.cell(row=row_idx, column=col_idx, value=cell_value)
        
        # Add chart
        print("   Adding sales trend chart...")
        chart = office.excel.BarChart()
        chart.title = "产品销售趋势"
        chart.x_axis.title = "月份"
        chart.y_axis.title = "销售额"
        
        # Add data to chart (simplified)
        # In production, you'd use proper data references
        
        sales_sheet.add_chart(chart, "H2")
        
        # Save workbook
        excel_path = output_dir / "sales_analysis.xlsx"
        workbook.save(excel_path)
        print(f"   [OK] Excel analysis saved: {excel_path}")
        
        # ============================================
        # 3. Create PowerPoint Presentation
        # ============================================
        print("\n4. Creating PowerPoint presentation...")
        
        # Create presentation
        presentation = office.powerpoint.create_presentation()
        
        # Add title slide
        print("   Adding title slide...")
        title_slide = presentation.add_slide(
            layout="title",
            title="季度业务报告",
            content="2024年第一季度\n业务分析部"
        )
        
        # Add agenda slide
        print("   Adding agenda slide...")
        agenda_slide = presentation.add_slide(
            layout="title_and_content",
            title="议程",
            content="• 执行摘要\n• 销售数据分析\n• 市场表现\n• 财务概况\n• 结论与建议"
        )
        
        # Add sales summary slide
        print("   Adding sales summary slide...")
        sales_slide = presentation.add_slide(
            layout="title_and_content",
            title="销售业绩概览",
            content="• 总销售额: 3020万元 (+15%)\n• 毛利率: 40% (+3%)\n• 净利润: 423万元 (+12%)\n• 现金流: 580万元 (+12%)"
        )
        
        # Add market share slide
        print("   Adding market share slide...")
        market_slide = presentation.add_slide(
            layout="title_and_content",
            title="市场份额变化",
            content="• 华北: 25% (+3%)\n• 华东: 22% (+2%)\n• 华南: 20% (+2%)\n• 西部: 18% (+3%)\n• 海外: 15% (+5%)"
        )
        
        # Add recommendations slide
        print("   Adding recommendations slide...")
        rec_slide = presentation.add_slide(
            layout="title_and_content",
            title="建议措施",
            content="1. 加大新产品研发投入\n2. 深化海外市场布局\n3. 优化供应链管理\n4. 加强数字化转型"
        )
        
        # Add thank you slide
        print("   Adding thank you slide...")
        thank_you_slide = presentation.add_slide(
            layout="title_only",
            title="谢谢！\n问题与讨论"
        )
        
        # Save presentation
        ppt_path = output_dir / "business_presentation.pptx"
        presentation.save(ppt_path)
        print(f"   [OK] PowerPoint presentation saved: {ppt_path}")
        
        # ============================================
        # 4. Convert to PDF (if possible)
        # ============================================
        print("\n5. Converting to PDF formats...")
        
        try:
            # Convert Word to PDF
            pdf_report_path = output_dir / "business_report.pdf"
            office.converter.convert(report_path, "pdf", pdf_report_path)
            print(f"   [OK] PDF report: {pdf_report_path}")
            
        except Exception as e:
            print(f"   [NOTE] PDF conversion requires additional setup: {e}")
        
        # ============================================
        # 5. WPS Integration (if available)
        # ============================================
        print("\n6. Checking WPS integration...")
        
        if office.wps.available:
            print(f"   [OK] WPS Office detected: {office.wps.version}")
            
            # Optimize for WPS
            wps_report_path = output_dir / "business_report_wps.wps"
            office.wps.optimize_for_wps(report_path, wps_report_path)
            print(f"   [OK] WPS-optimized report: {wps_report_path}")
        else:
            print("   [INFO] WPS Office not detected. Using standard Office formats.")
            print("   [NOTE] All generated files are compatible with WPS Office.")
        
        # ============================================
        # Summary
        # ============================================
        print("\n" + "=" * 60)
        print("REPORT GENERATION COMPLETE")
        print("=" * 60)
        
        print(f"\nGenerated files in '{output_dir}':")
        print(f"1. Word Report: {report_path}")
        print(f"2. Excel Analysis: {excel_path}")
        print(f"3. PowerPoint Presentation: {ppt_path}")
        
        if office.wps.available:
            print(f"4. WPS Optimized: {wps_report_path}")
        
        print("\nAll files are compatible with:")
        print("- Microsoft Office")
        print("- WPS Office")
        print("- LibreOffice")
        print("- Google Docs (with conversion)")
        
        return True
        
    except Exception as e:
        print(f"\n[ERROR] Failed to create report: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Main function to run the report creation example."""
    print("Office Automation - Business Report Example")
    print("=" * 60)
    print("\nThis example demonstrates creating a complete business report")
    print("with Word, Excel, and PowerPoint files.")
    
    success = create_business_report()
    
    if success:
        print("\nExample completed successfully!")
        print("\nNext steps:")
        print("1. Open the generated files in Office or WPS")
        print("2. Modify the example to use your own data")
        print("3. Integrate with your business systems")
    else:
        print("\n❌ Example failed. Check error messages above.")
    
    return success


if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⚠️ Interrupted by user.")
        sys.exit(1)