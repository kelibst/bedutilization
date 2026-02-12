import os
import sys
from openpyxl import Workbook
from config import WorkbookConfig
import phase1_structure as p1
import test_excel_open

def debug_build():
    print("Starting Debug Build...")
    config = WorkbookConfig(year=2026)
    output_path = "Debug.xlsx"
    
    wb = Workbook()
    
    # 1. Control sheet
    print("Building Control sheet...")
    p1.build_control_sheet(wb, config)
    
    # 2. DailyData
    print("Building DailyData sheet...")
    # p1.build_daily_data_sheet(wb, config)
    
    # Comment out others to start
    # p1.build_admissions_sheet(wb, config)
    # p1.build_deaths_data_sheet(wb, config)
    # p1.build_transfers_sheet(wb, config)
    # for ward in config.WARDS:
    #     p1.build_ward_sheet(wb, config, ward)
    # p1.build_emergency_combined_sheet(wb, config)
    # p1.build_monthly_summary_sheet(wb, config)
    # p1.build_ages_summary_sheet(wb, config)
    # p1.build_deaths_report_sheet(wb, config)
    # p1.build_cod_summary_sheet(wb, config)
    # p1.build_statement_of_inpatient_sheet(wb, config)
    # p1.build_non_insured_report_sheet(wb, config)

    wb.save(output_path)
    print(f"Saved {output_path}")
    
    print("Testing if Excel can open it...")
    test_excel_open.test_excel_open(output_path)

if __name__ == "__main__":
    try:
        debug_build()
    except Exception as e:
        print(f"Build failed: {e}")
