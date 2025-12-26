import os
import time
import win32com.client as win32
from pathlib import Path

# =====================
# CONFIG
# =====================
EXCEL_FILE = r"path"
OUTPUT_DIR = r"path"
FILTER_FIELD = "****"
IMAGE_FORMAT = "png"   # png, jpg, bmp, gif

# =====================
# SETUP
# =====================
Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

excel = win32.DispatchEx("Excel.Application")
excel.Visible = False  # Set to True for debugging
excel.DisplayAlerts = False
excel.ScreenUpdating = False
excel.EnableEvents = False

# 🔴 CRITICAL: Set calculation to manual to prevent premature updates
#excel.Calculation = -4135  # xlCalculationManual

try:
    wb = excel.Workbooks.Open(EXCEL_FILE)

    # --------------------------------------------------
    # STEP 1: READ FILTER VALUES ONCE
    # --------------------------------------------------
    filter_values = None

    for sheet in wb.Worksheets:
        if sheet.PivotTables().Count > 0:
            pivot = sheet.PivotTables(1)
            try:
                field = pivot.PivotFields(FILTER_FIELD)
                field.EnableMultiplePageItems = False

                filter_values = [
                    item.Name
                    for item in field.PivotItems()
                    if item.Name and item.Name != "(blank)"
                ]
                break
            except Exception:
                continue

    if not filter_values:
        raise RuntimeError(f"Could not find field '{FILTER_FIELD}'")

    print(f"Found {len(filter_values)} values: {filter_values}")

    # --------------------------------------------------
    # STEP 2: PROCESS EACH FILTER VALUE
    # --------------------------------------------------
    for value_idx, value in enumerate(filter_values):
        print(f"\n{'='*60}")
        print(f"Processing [{value_idx+1}/{len(filter_values)}]: {value}")
        print('='*60)

        # --------------------------------------------------
        # 2A: APPLY FILTER TO ALL PIVOTS
        # --------------------------------------------------
        pivots_updated = 0
        for sheet in wb.Worksheets:
            if sheet.PivotTables().Count == 0:
                continue

            for pivot_idx in range(1, sheet.PivotTables().Count + 1):
                pivot = sheet.PivotTables(pivot_idx)
                try:
                    field = pivot.PivotFields(FILTER_FIELD)
                    field.EnableMultiplePageItems = False
                    
                    # 🔴 CRITICAL: Clear current page first
                    field.ClearAllFilters()
                    time.sleep(0.1)
                    
                    # Set new filter value
                    field.CurrentPage = value
                    pivots_updated += 1
                    
                except Exception as e:
                    print(f"  ⚠ Could not update {sheet.Name}.{pivot.Name}: {e}")
                    continue

        print(f"  ✓ Updated {pivots_updated} pivot tables")
        
        # --------------------------------------------------
        # 2B: FORCE REFRESH ALL PIVOTS
        # --------------------------------------------------
        time.sleep(0.5)  # Let filters settle
        
        for sheet in wb.Worksheets:
            if sheet.PivotTables().Count == 0:
                continue
            for pivot_idx in range(1, sheet.PivotTables().Count + 1):
                try:
                    pivot = sheet.PivotTables(pivot_idx)
                    pivot.RefreshTable()
                except Exception:
                    pass
        
        # --------------------------------------------------
        # 2C: FORCE FULL RECALCULATION
        # --------------------------------------------------
        excel.CalculateFull()
        time.sleep(1.0)  # 🔴 Increased wait time
        
        # 🔴 Additional recalc to ensure charts update
        wb.RefreshAll()
        time.sleep(0.5)

        # --------------------------------------------------
        # STEP 3: EXPORT CHARTS
        # --------------------------------------------------
        charts_exported = 0
        
        for sheet in wb.Worksheets:
            pivots = sheet.PivotTables()
            charts = sheet.ChartObjects()
            
            if pivots.Count == 0:
                continue

            for pivot_idx in range(1, pivots.Count + 1):
                pivot = pivots(pivot_idx)

                for chart_idx in range(1, charts.Count + 1):
                    chart_obj = charts(chart_idx)
                    chart = chart_obj.Chart

                    # Check if chart is linked to this pivot
                    try:
                        pivot_of_chart = chart.PivotLayout.PivotTable
                        if pivot_of_chart.Name != pivot.Name:
                            continue
                    except Exception:
                        continue

                    # 🔴 FORCE CHART REFRESH
                    try:
                        chart.Refresh()
                        time.sleep(0.3)  # Wait for chart to redraw
                    except Exception:
                        pass

                    # Generate filename with directory per filter value
                    safe_value = "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in str(value))
                    
                    # Create directory for this filter value
                    value_dir = os.path.join(OUTPUT_DIR, safe_value)
                    Path(value_dir).mkdir(parents=True, exist_ok=True)
                    
                    # Clean sheet name for filename
                    safe_sheet_name = "".join(c if c.isalnum() or c in (' ', '_', '-') else '_' for c in sheet.Name)
                    filename = f"{safe_sheet_name}_chart{chart_idx}.{IMAGE_FORMAT}"
                    output_path = os.path.join(value_dir, filename)
                    
                    # 🔴 EXPORT WITH RETRY LOGIC
                    max_retries = 3
                    for attempt in range(max_retries):
                        try:
                            # Ensure path length is not too long (Windows MAX_PATH = 260)
                            if len(output_path) > 250:
                                print(f"  ⚠ Path too long, skipping: {output_path}")
                                break
                            
                            chart.Export(output_path)
                            charts_exported += 1
                            print(f"  ✓ Exported: {safe_value}/{filename}")
                            break
                        except Exception as e:
                            if attempt < max_retries - 1:
                                print(f"  ⚠ Export attempt {attempt+1} failed, retrying...")
                                time.sleep(0.5)
                            else:
                                print(f"  ✗ Failed to export after {max_retries} attempts: {safe_value}/{filename}")
                                print(f"    Error: {str(e)}")

        print(f"  → Total charts exported for '{value}': {charts_exported}")

    print(f"\n{'='*60}")
    print(f"✅ COMPLETE! Processed {len(filter_values)} filter values")
    print('='*60)

finally:
    # 🔴 RESTORE EXCEL STATE
    #excel.Calculation = -4105  # xlCalculationAutomatic
    excel.ScreenUpdating = True
    excel.EnableEvents = True

    wb.Close(SaveChanges=False)  # Don't save changes to prevent corruption
    excel.Quit()
