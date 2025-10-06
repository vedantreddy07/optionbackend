"""
Excel Diagnostics Tool
This script will help identify the correct button name and macro to trigger
"""

import win32com.client as win32
import pythoncom
from pathlib import Path

def diagnose_excel_file():
    """Diagnose Excel file to find buttons and macros"""
    
    pythoncom.CoInitialize()
    
    try:
        print("="*70)
        print("EXCEL DIAGNOSTICS TOOL")
        print("="*70)
        
        # Connect to Excel
        print("\n1. Connecting to Excel...")
        excel = win32.GetActiveObject("Excel.Application")
        print("   ✓ Connected to Excel")
        
        # Find workbook
        print("\n2. Looking for SmartOptionChain workbook...")
        wb = None
        for workbook in excel.Workbooks:
            print(f"   Found: {workbook.Name}")
            if "SmartOptionChain" in workbook.Name:
                wb = workbook
                print(f"   ✓ Selected: {workbook.Name}")
                break
        
        if not wb:
            print("   ✗ Workbook not found!")
            return
        
        # Get Option_Chain sheet
        print("\n3. Accessing Option_Chain sheet...")
        ws = wb.Sheets("Option_Chain")
        print("   ✓ Sheet found")
        
        # List all shapes (buttons)
        print("\n4. SHAPES/BUTTONS IN SHEET:")
        print("-" * 70)
        for i, shape in enumerate(ws.Shapes, 1):
            print(f"\n   Shape {i}:")
            print(f"      Name: {shape.Name}")
            print(f"      Type: {shape.Type}")
            
            try:
                if hasattr(shape, 'OnAction'):
                    print(f"      OnAction: {shape.OnAction}")
            except:
                print(f"      OnAction: (not accessible)")
            
            try:
                if hasattr(shape, 'AlternativeText'):
                    print(f"      Alt Text: {shape.AlternativeText}")
            except:
                pass
            
            try:
                if hasattr(shape, 'TextFrame'):
                    text = shape.TextFrame.Characters().Text
                    print(f"      Text: {text}")
            except:
                pass
        
        # List all VBA modules and macros
        print("\n\n5. VBA MODULES AND MACROS:")
        print("-" * 70)
        try:
            vb_project = wb.VBProject
            
            for component in vb_project.VBComponents:
                print(f"\n   Module: {component.Name}")
                print(f"   Type: {component.Type}")
                
                try:
                    code_module = component.CodeModule
                    line_count = code_module.CountOfLines
                    
                    # Find all Sub procedures
                    subs_found = []
                    for i in range(1, min(line_count + 1, 500)):  # Limit to first 500 lines
                        try:
                            line = code_module.Lines(i, 1)
                            if line.strip().startswith("Sub ") or line.strip().startswith("Public Sub "):
                                # Extract sub name
                                sub_name = line.split("Sub ")[1].split("(")[0].strip()
                                subs_found.append(sub_name)
                        except:
                            continue
                    
                    if subs_found:
                        print(f"   Procedures found:")
                        for sub in subs_found:
                            print(f"      - {sub}")
                except Exception as e:
                    print(f"   Could not read code: {e}")
        
        except Exception as e:
            print(f"   VBA Project not accessible: {e}")
            print("   This is normal if Trust Access to VBA is not enabled")
        
        # Check specific cells
        print("\n\n6. DROPDOWN CELLS CONFIGURATION:")
        print("-" * 70)
        
        cells_to_check = {
            "B2": "Symbol",
            "B3": "Option Expiry",
            "B4": "Future Expiry",
            "B6": "Chain Length",
            "F587": "User ID",
            "F615": "Token"
        }
        
        for cell_ref, description in cells_to_check.items():
            try:
                cell = ws.Range(cell_ref)
                value = cell.Value
                
                print(f"\n   Cell {cell_ref} ({description}):")
                print(f"      Current Value: {value}")
                
                # Check if it has validation
                try:
                    if hasattr(cell, 'Validation'):
                        validation = cell.Validation
                        if validation.Type == 3:  # List validation
                            formula = validation.Formula1
                            print(f"      Validation: {formula}")
                except:
                    print(f"      Validation: None")
                    
            except Exception as e:
                print(f"   Cell {cell_ref}: Error - {e}")
        
        # Provide recommendations
        print("\n\n7. RECOMMENDATIONS:")
        print("="*70)
        print("\nBased on the diagnostics above:")
        print("\n1. Look for a button with 'Option Chain' or similar text")
        print("2. Note the button's exact Name (e.g., 'Button 2', 'btnFetch', etc.)")
        print("3. Note the OnAction macro name if shown")
        print("\n4. Update final_excel_handler.py with:")
        print("   - Correct button name in ws.Shapes('YOUR_BUTTON_NAME')")
        print("   - Correct macro name in excel.Run('YOUR_MACRO_NAME')")
        
        print("\n" + "="*70)
        
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    print("\nMake sure your SmartOptionChainExcel file is OPEN in Excel")
    input("Press Enter to run diagnostics...")
    
    diagnose_excel_file()
    
    print("\n\nDiagnostics complete!")
    input("Press Enter to exit...")