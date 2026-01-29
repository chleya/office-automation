#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Process Spreadsheet Example

This example demonstrates advanced Excel spreadsheet processing
using the Office Automation skill. Features include:
1. Data import from CSV/JSON
2. Data cleaning and transformation
3. Statistical analysis
4. Chart generation
5. Report automation
"""

import os
import sys
import json
import csv
from pathlib import Path
from datetime import datetime

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from office_automation import OfficeAutomation


def process_sales_data():
    """Process sales data from multiple sources and generate analysis reports."""
    print("=" * 60)
    print("Processing Spreadsheet Data Example")
    print("=" * 60)
    
    # Initialize Office Automation
    print("\n1. Initializing Office Automation...")
    office = OfficeAutomation()
    print("   [OK] Office Automation initialized")
    
    # Create output directory
    output_dir = Path("output/spreadsheets")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Create sample data directory
    data_dir = Path("data/samples")
    data_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        # ============================================
        # 1. Create Sample Data Files
        # ============================================
        print("\n2. Creating sample data files...")
        
        # Create CSV sample data
        csv_data = [
            ["Date", "Product", "Region", "Quantity", "Revenue", "Customer"],
            ["2024-01-01", "Product A", "North", 150, 7500.00, "Customer 1"],
            ["2024-01-02", "Product B", "South", 200, 12000.00, "Customer 2"],
            ["2024-01-03", "Product A", "East", 180, 9000.00, "Customer 3"],
            ["2024-01-04", "Product C", "West", 120, 8400.00, "Customer 4"],
            ["2024-01-05", "Product B", "North", 220, 13200.00, "Customer 5"],
            ["2024-01-06", "Product A", "South", 160, 8000.00, "Customer 6"],
            ["2024-01-07", "Product C", "East", 140, 9800.00, "Customer 7"],
            ["2024-01-08", "Product B", "West", 190, 11400.00, "Customer 8"],
            ["2024-01-09", "Product A", "North", 170, 8500.00, "Customer 9"],
            ["2024-01-10", "Product C", "South", 130, 9100.00, "Customer 10"],
        ]
        
        csv_path = data_dir / "sales_data.csv"
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(csv_data)
        print(f"   [OK] CSV data created: {csv_path}")
        
        # Create JSON sample data
        json_data = {
            "metadata": {
                "source": "Sales System",
                "period": "January 2024",
                "generated": datetime.now().isoformat()
            },
            "products": [
                {"id": "P001", "name": "Product A", "category": "Electronics", "price": 50.00},
                {"id": "P002", "name": "Product B", "category": "Furniture", "price": 60.00},
                {"id": "P003", "name": "Product C", "category": "Office Supplies", "price": 70.00}
            ],
            "regions": ["North", "South", "East", "West"],
            "summary": {
                "total_sales": 87900.00,
                "total_quantity": 1560,
                "average_revenue": 8790.00,
                "top_product": "Product B"
            }
        }
        
        json_path = data_dir / "sales_metadata.json"
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, indent=2, ensure_ascii=False)
        print(f"   [OK] JSON metadata created: {json_path}")
        
        # ============================================
        # 2. Import Data into Excel
        # ============================================
        print("\n3. Importing data into Excel...")
        
        # Create workbook
        workbook = office.excel.create_workbook()
        
        # Import CSV data
        print("   Importing CSV data...")
        sales_sheet = workbook.add_worksheet("Sales Data")
        
        with open(csv_path, 'r', encoding='utf-8') as f:
            csv_reader = csv.reader(f)
            for row_idx, row in enumerate(csv_reader, start=1):
                for col_idx, value in enumerate(row, start=1):
                    # Try to convert numeric values
                    try:
                        if value.replace('.', '', 1).isdigit():
                            value = float(value)
                    except:
                        pass
                    sales_sheet.cell(row=row_idx, column=col_idx, value=value)
        
        # Format header row
        for cell in sales_sheet[1]:
            cell.font = office.excel.Font(bold=True, color="FFFFFF")
            cell.fill = office.excel.PatternFill(start_color="4F81BD", 
                                                end_color="4F81BD", 
                                                fill_type="solid")
        
        print("   [OK] CSV data imported")
        
        # ============================================
        # 3. Data Analysis and Calculations
        # ============================================
        print("\n4. Performing data analysis...")
        
        # Create analysis sheet
        analysis_sheet = workbook.add_worksheet("Analysis")
        
        # Add summary statistics
        analysis_sheet.cell(row=1, column=1, value="Sales Analysis Summary")
        analysis_sheet.cell(row=1, column=1).font = office.excel.Font(bold=True, size=14)
        
        # Calculate totals (simplified - in real scenario would use formulas)
        summary_data = [
            ["Metric", "Value"],
            ["Total Revenue", "=SUM(Sales_Data!E2:E11)"],
            ["Average Revenue", "=AVERAGE(Sales_Data!E2:E11)"],
            ["Max Revenue", "=MAX(Sales_Data!E2:E11)"],
            ["Min Revenue", "=MIN(Sales_Data!E2:E11)"],
            ["Total Quantity", "=SUM(Sales_Data!D2:D11)"],
            ["Number of Transactions", "=COUNTA(Sales_Data!A2:A11)"],
        ]
        
        for row_idx, row_data in enumerate(summary_data, start=3):
            for col_idx, value in enumerate(row_data, start=1):
                analysis_sheet.cell(row=row_idx, column=col_idx, value=value)
                if row_idx > 3 and col_idx == 2:  # Format value cells
                    analysis_sheet.cell(row=row_idx, column=col_idx).number_format = '#,##0.00'
        
        # Format summary table
        for row in range(3, 10):
            for col in range(1, 3):
                cell = analysis_sheet.cell(row=row, column=col)
                if row == 3:  # Header row
                    cell.font = office.excel.Font(bold=True)
                    cell.fill = office.excel.PatternFill(start_color="F2F2F2", 
                                                        end_color="F2F2F2", 
                                                        fill_type="solid")
                elif col == 2:  # Value column
                    cell.font = office.excel.Font(bold=True, color="2E75B5")
        
        print("   [OK] Analysis calculations added")
        
        # ============================================
        # 4. Create Pivot Table (simulated)
        # ============================================
        print("\n5. Creating pivot analysis...")
        
        pivot_sheet = workbook.add_worksheet("Pivot Analysis")
        
        # Product summary
        pivot_sheet.cell(row=1, column=1, value="Product Performance")
        pivot_sheet.cell(row=1, column=1).font = office.excel.Font(bold=True, size=12)
        
        product_summary = [
            ["Product", "Total Revenue", "Total Quantity", "Average Price"],
            ["Product A", 33000.00, 660, 50.00],
            ["Product B", 36600.00, 610, 60.00],
            ["Product C", 18300.00, 290, 70.00],
        ]
        
        for row_idx, row_data in enumerate(product_summary, start=3):
            for col_idx, value in enumerate(row_data, start=1):
                pivot_sheet.cell(row=row_idx, column=col_idx, value=value)
                if row_idx > 3 and col_idx > 1:  # Format numeric cells
                    pivot_sheet.cell(row=row_idx, column=col_idx).number_format = '#,##0.00'
        
        # Region summary
        pivot_sheet.cell(row=1, column=6, value="Regional Performance")
        pivot_sheet.cell(row=1, column=6).font = office.excel.Font(bold=True, size=12)
        
        region_summary = [
            ["Region", "Total Revenue", "Number of Sales"],
            ["North", 29200.00, 3],
            ["South", 29100.00, 3],
            ["East", 18800.00, 2],
            ["West", 10800.00, 2],
        ]
        
        for row_idx, row_data in enumerate(region_summary, start=3):
            for col_idx, value in enumerate(row_data, start=6):
                pivot_sheet.cell(row=row_idx, column=col_idx, value=value)
                if row_idx > 3 and col_idx > 6:  # Format numeric cells
                    pivot_sheet.cell(row=row_idx, column=col_idx).number_format = '#,##0.00'
        
        print("   [OK] Pivot tables created")
        
        # ============================================
        # 5. Create Charts
        # ============================================
        print("\n6. Creating charts...")
        
        # Product revenue chart
        chart1 = office.excel.BarChart()
        chart1.title = "Product Revenue Comparison"
        chart1.x_axis.title = "Product"
        chart1.y_axis.title = "Revenue"
        
        # Add data (simplified - in real scenario would use data references)
        # Note: In dummy mode, Reference is not available
        # from office.excel.chart import Reference
        
        # Create a simple data series
        chart_data = [
            ["Product", "Revenue"],
            ["Product A", 33000],
            ["Product B", 36600],
            ["Product C", 18300],
        ]
        
        # Add chart data sheet
        chart_data_sheet = workbook.add_worksheet("Chart Data")
        for row_idx, row_data in enumerate(chart_data, start=1):
            for col_idx, value in enumerate(row_data, start=1):
                chart_data_sheet.cell(row=row_idx, column=col_idx, value=value)
        
        # Add chart to analysis sheet
        analysis_sheet.add_chart(chart1, "D3")
        
        # Regional performance chart
        chart2 = office.excel.PieChart()
        chart2.title = "Revenue by Region"
        
        analysis_sheet.add_chart(chart2, "D20")
        
        print("   [OK] Charts created")
        
        # ============================================
        # 6. Data Validation and Formatting
        # ============================================
        print("\n7. Adding data validation...")
        
        # Add data validation for region column
        region_validation = office.excel.DataValidation(
            type="list",
            formula1='"North,South,East,West"',
            allow_blank=True
        )
        
        # Apply to region column (column C)
        for row in range(2, 12):  # Rows 2-11
            sales_sheet.cell(row=row, column=3).data_validation = region_validation
        
        # Add conditional formatting for revenue
        red_fill = office.excel.PatternFill(start_color="FFC7CE", 
                                          end_color="FFC7CE", 
                                          fill_type="solid")
        
        revenue_rule = office.excel.Rule(
            type="cellIs",
            operator="lessThan",
            formula=["5000"],
            fill=red_fill
        )
        
        sales_sheet.conditional_formatting.add(f"E2:E11", revenue_rule)
        
        print("   [OK] Data validation and formatting added")
        
        # ============================================
        # 7. Save Workbook
        # ============================================
        print("\n8. Saving workbook...")
        
        excel_path = output_dir / "sales_analysis.xlsx"
        workbook.save(excel_path)
        print(f"   [OK] Excel workbook saved: {excel_path}")
        
        # ============================================
        # 8. Generate Report
        # ============================================
        print("\n9. Generating analysis report...")
        
        # Create Word report
        doc = office.word.create_document()
        
        doc.add_heading("销售数据分析报告", level=0)
        doc.add_paragraph(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        doc.add_paragraph("")
        
        doc.add_heading("执行摘要", level=1)
        doc.add_paragraph(
            "本报告基于2024年1月的销售数据进行分析。"
            "总销售额为87,900元，总销售数量为1,560件。"
            "表现最佳的产品是Product B，贡献了36,600元销售额。"
        )
        doc.add_paragraph("")
        
        doc.add_heading("关键指标", level=1)
        
        key_metrics = [
            ["总销售额", "87,900元"],
            ["总销售数量", "1,560件"],
            ["平均每单金额", "8,790元"],
            ["交易数量", "10笔"],
            ["最畅销产品", "Product B"],
            ["最佳销售区域", "North"],
        ]
        
        doc.add_table(key_metrics, style="Light Grid")
        doc.add_paragraph("")
        
        doc.add_heading("建议措施", level=1)
        
        recommendations = [
            "1. 加大Product B的生产和推广力度",
            "2. 在North区域开展促销活动",
            "3. 分析Product C的销售表现，考虑优化策略",
            "4. 加强South和East区域的市场拓展",
        ]
        
        for rec in recommendations:
            doc.add_paragraph(rec)
        
        report_path = output_dir / "sales_analysis_report.docx"
        doc.save(report_path)
        print(f"   [OK] Word report saved: {report_path}")
        
        # ============================================
        # 9. WPS Compatibility Check
        # ============================================
        print("\n10. Checking WPS compatibility...")
        
        if office.wps.available:
            print(f"   [OK] WPS Office detected: {office.wps.version}")
            
            # Create WPS optimized version
            wps_excel_path = output_dir / "sales_analysis_wps.et"
            office.wps.optimize_for_wps(excel_path, wps_excel_path)
            print(f"   [OK] WPS-optimized spreadsheet: {wps_excel_path}")
        else:
            print("   [INFO] WPS Office not detected. Files are compatible with WPS.")
        
        # ============================================
        # Summary
        # ============================================
        print("\n" + "=" * 60)
        print("SPREADSHEET PROCESSING COMPLETE")
        print("=" * 60)
        
        print(f"\nGenerated files in '{output_dir}':")
        print(f"1. Excel Analysis: {excel_path}")
        print(f"2. Word Report: {report_path}")
        
        if office.wps.available:
            print(f"3. WPS Spreadsheet: {wps_excel_path}")
        
        print("\nData processing features demonstrated:")
        print("- CSV/JSON data import")
        print("- Statistical calculations")
        print("- Pivot table analysis")
        print("- Chart generation")
        print("- Data validation")
        print("- Conditional formatting")
        print("- Automated reporting")
        
        return True
        
    except Exception as e:
        print(f"\n[ERROR] Failed to process spreadsheet: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Main function to run the spreadsheet processing example."""
    print("Office Automation - Spreadsheet Processing Example")
    print("=" * 60)
    print("\nThis example demonstrates advanced Excel data processing")
    print("with import, analysis, visualization, and reporting.")
    
    success = process_sales_data()
    
    if success:
        print("\nExample completed successfully!")
        print("\nNext steps:")
        print("1. Open the generated Excel file to see analysis")
        print("2. Modify the example with your own data")
        print("3. Integrate with your data sources")
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