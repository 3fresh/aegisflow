"""
Extract PROGRAM list from MOSAIC_CONVERT generated Excel file
and generate SAS run script format text file

Features:
1. Read Index worksheet from Excel file
2. Count unique values in PROGRAM column and their corresponding tocnumber counts
3. Sort by tocnumber count in descending order
4. Generate %runpgm format text file
"""

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog


def analyze_programs(excel_file):
    """
    Analyze PROGRAM column in Excel file
    
    Parameters:
    -----------
    excel_file : str
        Excel file path
        
    Returns:
    --------
    program_stats : pandas.DataFrame
        Statistical data containing program, tocnumber_count, tocnumber_list
    """
    print("=" * 80)
    print("Analyzing Excel file...")
    print("=" * 80)
    
    # Read Index worksheet
    try:
        df = pd.read_excel(excel_file, sheet_name='Index')
        print(f"OK: Successfully read Index worksheet ({len(df)} rows of data)")
    except Exception as e:
        print(f"ERROR: Failed to read Excel file: {e}")
        return None
    
    # Check required columns
    if 'PROGRAM' not in df.columns:
        print("ERROR: PROGRAM column not found")
        return None
    
    if 'tocnumber' not in df.columns:
        print("ERROR: tocnumber column not found")
        return None
    
    # Count tocnumber for each PROGRAM
    program_data = []
    
    # Use dictionary to maintain first occurrence order
    seen_programs = {}
    order_index = 0
    
    for program in df['PROGRAM']:
        if pd.isna(program) or program == '':
            continue
        
        if program not in seen_programs:
            seen_programs[program] = order_index
            order_index += 1
    
    # Process each program in first occurrence order
    for program in seen_programs.keys():
        # Get all tocnumbers corresponding to this program
        toc_numbers = df[df['PROGRAM'] == program]['tocnumber'].dropna().unique()
        toc_count = len(toc_numbers)
        toc_list = ', '.join(sorted(str(t) for t in toc_numbers))
        
        program_data.append({
            'PROGRAM': program,
            'tocnumber_count': toc_count,
            'tocnumber_list': toc_list,
            'order': seen_programs[program]
        })
    
    # Create DataFrame and sort by first occurrence order
    program_stats = pd.DataFrame(program_data)
    program_stats = program_stats.sort_values('order')
    
    print(f"\nOK: Statistics complete")
    print(f"  - Total unique PROGRAM values: {len(program_stats)}")
    print(f"  - Total tables/listings/figures: {len(df)}")
    
    return program_stats


def generate_sas_script(program_stats):
    """
    Generate SAS run script format text
    
    Parameters:
    -----------
    program_stats : pandas.DataFrame
        Program statistical data
        
    Returns:
    --------
    script_lines : list
        脚本行列表
    """
    script_lines = []
    
    # Add comment header
    script_lines.append("/* Generated SAS Program Execution Script */")
    script_lines.append("/* Programs ordered by first appearance in Excel file */")
    script_lines.append("")
    script_lines.append("/* Program Statistics: */")
    
    # Add all program comment information uniformly
    for _, row in program_stats.iterrows():
        program = row['PROGRAM']
        count = row['tocnumber_count']
        script_lines.append(f"/*   {program}: {count} table(s) */")
    
    script_lines.append("")
    script_lines.append("/* ====== Program Execution Commands ====== */")
    script_lines.append("")
    
    # Add execution commands for all programs
    for _, row in program_stats.iterrows():
        program = row['PROGRAM']
        script_lines.append(f"%runpgm(pgm={program}, error_override=y);")
    
    return script_lines


def save_script_file(script_lines, output_file):
    """
    Save script to file
    
    Parameters:
    -----------
    script_lines : list
        List of script lines
    output_file : str
        Output file path
    """
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(script_lines))
        print(f"\nOK: Script saved to: {output_file}")
        return True
    except Exception as e:
        print(f"\nERROR: Failed to save file: {e}")
        return False


def print_statistics(program_stats):
    """
    Print statistical information
    
    Parameters:
    -----------
    program_stats : pandas.DataFrame
        Program statistical data
    """
    print("\n" + "=" * 80)
    print("PROGRAM Statistics (ordered by first appearance in Excel)")
    print("=" * 80)
    print(f"{'No.':<6} {'PROGRAM':<30} {'Table Count':<10}")
    print("-" * 80)
    
    for idx, row in enumerate(program_stats.iterrows(), 1):
        _, data = row
        print(f"{idx:<6} {data['PROGRAM']:<30} {data['tocnumber_count']:<10}")
    
    print("-" * 80)
    print(f"Total: {len(program_stats)} unique programs")
    print("=" * 80)


def main():
    """Main function"""
    print("=" * 80)
    print("EXTRACT_PROGRAMS - Extract SAS Program List from Excel")
    print("=" * 80)
    print()
    
    # Create GUI to select input file
    root = tk.Tk()
    root.withdraw()
    
    print("[1] Select input Excel file...")
    input_file = filedialog.askopenfilename(
        title="Select Excel file generated by MOSAIC_CONVERT",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialdir=os.path.dirname(os.path.abspath(__file__))
    )
    
    if not input_file:
        print("WARN: No file selected, exiting program")
        return
    
    print(f"Selected file: {os.path.basename(input_file)}")
    
    # Analyze PROGRAM column
    print("\n[2] Analyzing PROGRAM column...")
    program_stats = analyze_programs(input_file)
    
    if program_stats is None or len(program_stats) == 0:
        print("ERROR: Analysis failed or no valid PROGRAM data found")
        return
    
    # Print statistics
    print_statistics(program_stats)
    
    # Generate script
    print("\n[3] Generating SAS run script...")
    script_lines = generate_sas_script(program_stats)
    print(f"OK: Generated {len(script_lines)} lines of script")
    
    # Select output file
    print("\n[4] Select output file location...")
    
    # Default output filename
    default_name = "run_all_pgm_generated.txt"
    default_dir = os.path.dirname(input_file)
    
    output_file = filedialog.asksaveasfilename(
        title="Save SAS script file",
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        initialdir=default_dir,
        initialfile=default_name
    )
    
    if not output_file:
        print("WARN: No output file selected, exiting program")
        return
    
    # Save file
    if save_script_file(script_lines, output_file):
        print("\n" + "=" * 80)
        print("Complete!")
        print("=" * 80)
        print(f"Output file: {output_file}")
        print(f"Total programs: {len(program_stats)}")
        print("=" * 80)
    else:
        print("\nERROR: Save failed")
        exit(1)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n" + "=" * 80)
        print("ERROR: Program execution failed")
        print("=" * 80)
        print(f"Error message: {e}")
        print("\nDetailed error:")
        import traceback
        traceback.print_exc()
        exit(1)
