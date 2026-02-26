#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Generate batch list XML from Excel/CSV input
Generate formatted batch list XML file from Excel table

Author: Assistant
Date: 2026-02-11
"""

import pandas as pd
import re
from pathlib import Path
from datetime import datetime
from xml.dom import minidom
import xml.etree.ElementTree as ET
from collections import OrderedDict
import warnings

# Try importing tkinter
try:
    import tkinter as tk
    from tkinter import filedialog
    HAS_TKINTER = True
except ImportError:
    HAS_TKINTER = False
    print("Warning: tkinter not installed, using command line input mode")
    print("To use GUI file selection, please install tkinter")


class XMLGenerator:
    """Class for generating batch list XML"""
    
    def __init__(self):
        self.df = None
        self.header_text = ""
        self.file_location = ""
        
    def load_excel(self, file_path):
        """
        Read Excel or CSV file
        
        Args:
            file_path: Input file path
        """
        file_path = Path(file_path)
        
        if not file_path.exists():
            raise FileNotFoundError(f"File does not exist: {file_path}")
        
        try:
            if file_path.suffix.lower() == '.csv':
                # Try multiple encodings to read CSV
                encodings = ['utf-8', 'gbk', 'gb2312', 'latin1']
                for encoding in encodings:
                    try:
                        self.df = pd.read_csv(file_path, encoding=encoding)
                        print(f"Successfully read CSV file with {encoding} encoding")
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    raise ValueError("Unable to read CSV file with any supported encoding")
            else:
                # Read Excel file
                self.df = pd.read_excel(file_path)
                print("Successfully read Excel file")
                
            print(f"Read {len(self.df)} rows of data")
            self._validate_columns()
            
        except Exception as e:
            raise Exception(f"Failed to read file: {str(e)}")
    
    def _validate_columns(self):
        """Validate that required columns exist"""
        required_columns = ['sect_num', 'sect_ttl', 'OUTFILE', 
                          'Output Type (Table, Listing, Figure)', 
                          'tocnumber', 'Title']
        
        missing_columns = [col for col in required_columns if col not in self.df.columns]
        
        if missing_columns:
            print("\nWarning: Missing the following columns:")
            for col in missing_columns:
                print(f"  - {col}")
            print("\nAvailable columns:")
            for col in self.df.columns:
                print(f"  - {col}")
            raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
        
        print("✓ All required columns present")
    
    def _natural_sort_key(self, section_num):
        """
        Generate natural sort key to make 14.2.10 come after 14.2.2
        
        Args:
            section_num: section number string, e.g. "14.2.10"
            
        Returns:
            Tuple for sorting
        """
        if pd.isna(section_num):
            return (0,)
        
        # Split string by numbers
        parts = re.split(r'(\d+)', str(section_num))
        result = []
        
        for part in parts:
            if part.isdigit():
                result.append(int(part))
            elif part:
                result.append(part)
        
        return tuple(result)
    
    def _check_latin1_compatibility(self, text):
        """
        Check if text contains only latin1 characters
        
        Args:
            text: Text to check
            
        Returns:
            (is_compatible, problematic_chars): Compatibility status and list of problematic characters
        """
        if pd.isna(text):
            return True, []
        
        text = str(text)
        problematic_chars = []
        
        for char in text:
            try:
                char.encode('latin1')
            except UnicodeEncodeError:
                if char not in problematic_chars:
                    problematic_chars.append(char)
        
        return len(problematic_chars) == 0, problematic_chars
    
    def _validate_latin1(self):
        """Validate if all relevant fields contain only latin1 characters"""
        print("\nChecking character encoding compatibility...")
        
        issues = []
        columns_to_check = ['sect_ttl', 'Title']
        
        for idx, row in self.df.iterrows():
            for col in columns_to_check:
                if col in self.df.columns:
                    is_compatible, problematic_chars = self._check_latin1_compatibility(row[col])
                    if not is_compatible:
                        issues.append({
                            'row': idx + 2,  # Excel row number (starting from 1, plus header)
                            'column': col,
                            'value': str(row[col])[:50],  # Show only first 50 characters
                            'problematic_chars': problematic_chars
                        })
        
        if issues:
            print("\n⚠ Warning: Non-latin1 characters detected!")
            print("\nThe following content contains characters that cannot be encoded in latin1:")
            for issue in issues[:10]:  # Show only first 10 issues
                print(f"\n  Row {issue['row']}, Column '{issue['column']}':")
                print(f"    Content: {issue['value']}...")
                print(f"    Problematic characters: {', '.join(issue['problematic_chars'])}")
            
            if len(issues) > 10:
                print(f"\n  ... {len(issues) - 10} more issues not shown")
            
            print("\nSuggestions:")
            print("  1. Check for Chinese characters and special symbols in Excel file")
            print("  2. Replace non-ASCII characters with ASCII equivalents")
            print("  3. If XML needs to support these characters, consider using UTF-8 encoding")
            
            response = input("\nContinue generating XML? (y/n): ").strip().lower()
            if response != 'y':
                raise ValueError("Operation cancelled by user")
        else:
            print("✓ All characters are latin1 compatible")
    
    def generate_xml(self, header_text, file_location, output_path, output_filename, start_number=2):
        """
        Generate XML file
        
        Args:
            header_text: header text attribute value
            file_location: file location prefix
            output_path: output XML file path
            output_filename: output PDF filename (without .pdf extension)
            start_number: header startNumber attribute value
        """
        self.header_text = header_text
        self.file_location = file_location
        
        # Validate latin1 compatibility
        self._validate_latin1()
        
        print("\nStarting XML generation...")
        
        # Group by section
        sections = self._group_by_section()
        
        # Create XML structure - maintain exact consistency with reference XML
        root = ET.Element('pdf-builder-metadata')
        
        # 添加注释（需要在prettify时处理）
        comment_text = '<!-- input files total to less than 100MB -->'
        
        # Add ruleset
        ruleset = ET.SubElement(root, 'ruleset')
        
        # Add headers
        headers = ET.SubElement(ruleset, 'headers')
        header = ET.SubElement(headers, 'header')
        header.set('text', header_text)
        header.set('startNumber', str(start_number))
        
        # Add page configuration (fixed content)
        page = ET.SubElement(ruleset, 'page')
        page.set('orientation', 'landscape')
        page.set('size', 'letter')
        page.set('measurementUnit', 'in')
        page.set('marginTop', '           0')
        page.set('marginLeft', '           0')
        page.set('marginRight', '           0')
        page.set('marginBottom', '           0')
        
        # Add font configuration (fixed content)
        font = ET.SubElement(ruleset, 'font')
        font.set('fontName', 'CourierNew')
        font.set('style', 'normal')
        font.set('size', '9')
        
        # 添加注释（需要在prettify时处理）
        # <!-- <character-encoding type="ascii" /> -->
        
        # Add document-heading (fixed content, but text uses user-defined header_text)
        doc_heading = ET.SubElement(ruleset, 'document-heading')
        doc_heading.set('text', header_text)
        doc_heading.set('fontName', 'Times New Roman')
        
        # Add sectionset
        sectionset = ET.SubElement(root, 'sectionset')
        
        # Add sections in natural sort order
        sorted_sections = sorted(sections.items(), key=lambda x: self._natural_sort_key(x[0]))
        
        for section_key, files in sorted_sections:
            section_num, section_title = section_key
            
            section = ET.SubElement(sectionset, 'section')
            section.set('name', f"{section_num} {section_title}")
            
            # Add source-files
            for file_info in files:
                source_file = ET.SubElement(section, 'source-file')
                source_file.set('filename', file_info['filename'])
                source_file.set('fileLocation', file_info['fileLocation'])
                source_file.set('number', file_info['number'])
                source_file.set('title', file_info['title'])
        
        # Extract base path from file_location for pdf-import and audit-import
        # For example: "root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/" -> "root/cdar/d980/d9802c00001/ar/dr2"
        base_path = file_location
        if '/tlf/' in base_path:
            base_path = base_path.split('/tlf/')[0]
        base_path = base_path.rstrip('/')
        doc_path = f"{base_path}/tlf/doc/"
        
        # Add output-pdf (using custom output filename)
        output_pdf = ET.SubElement(root, 'output-pdf')
        output_pdf.set('filename', f'{output_filename}.pdf')
        pdf_import = ET.SubElement(output_pdf, 'pdf-import')
        pdf_import.set('path', doc_path)
        
        # Add output-audit (using custom output filename)
        output_audit = ET.SubElement(root, 'output-audit')
        output_audit.set('filename', f'{output_filename}_audit.pdf')
        audit_import = ET.SubElement(output_audit, 'audit-import')
        audit_import.set('path', doc_path)
        
        # Format XML
        xml_string = self._prettify_xml(root)
        
        # Save to file
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(xml_string)
        
        print(f"\n✓ XML file generated: {output_path}")
        print(f"  - Contains {len(sorted_sections)} sections")
        print(f"  - Contains {len(self.df)} source-files")
        
        return output_path
    
    def _group_by_section(self):
        """Group and sort by section"""
        sections = OrderedDict()
        
        # Clean data
        df_clean = self.df.copy()
        
        # 填充空值
        for col in df_clean.columns:
            if col in ['sect_num', 'sect_ttl', 'OUTFILE', 'Title']:
                df_clean[col] = df_clean[col].fillna('')
            if col == 'tocnumber':
                df_clean[col] = df_clean[col].fillna('')
        
        # Process each row in original order
        for idx, row in df_clean.iterrows():
            sect_num = str(row['sect_num']).strip()
            sect_ttl = str(row['sect_ttl']).strip()
            
            if not sect_num:
                print(f"Warning: Row {idx + 2} missing sect_num, skipping")
                continue
            
            section_key = (sect_num, sect_ttl)
            
            # Build source-file information
            outfile = str(row['OUTFILE']).strip()
            filename = f"{outfile}.rtf" if outfile else ""
            
            output_type = str(row['Output Type (Table, Listing, Figure)']).strip()
            tocnumber = str(row['tocnumber']).strip()
            number = f"{output_type} {tocnumber}" if output_type and tocnumber else ""
            
            title = str(row['Title']).strip()
            
            file_info = {
                'filename': filename,
                'fileLocation': self.file_location,
                'number': number,
                'title': title,
                'original_order': idx
            }
            
            # Add to section
            if section_key not in sections:
                sections[section_key] = []
            sections[section_key].append(file_info)
        
        return sections
    
    def _prettify_xml(self, elem):
        """Format XML for better readability, maintain exact consistency with reference XML"""
        # Use UTF-8 encoding
        rough_string = ET.tostring(elem, encoding='utf-8')
        reparsed = minidom.parseString(rough_string)
        
        # Get formatted XML with 4-space indentation
        pretty = reparsed.toprettyxml(indent="    ", encoding='UTF-8').decode('utf-8')
        
        # Remove extra blank lines
        lines = []
        for line in pretty.split('\n'):
            if line.strip():
                lines.append(line)
        
        result = '\n'.join(lines)
        
        # Add comment after <pdf-builder-metadata>
        result = result.replace(
            '<pdf-builder-metadata>',
            '<pdf-builder-metadata>\n<!-- input files total to less than 100MB -->'
        )
        
        # Add comment after font element
        result = result.replace(
            '<font fontName="CourierNew" style="normal" size="9"/>',
            '<font fontName="CourierNew" style="normal" size="9"/>\n        <!-- <character-encoding type="ascii" /> -->'
        )
        
        # Fix self-closing tag format to match reference XML
        # header tag
        result = result.replace(
            '<header text=',
            '<header text='
        ).replace('></header>', ' />')
        
        # page tag - maintain multi-line format
        result = result.replace(
            '<page orientation=',
            '<page\n            orientation='
        ).replace(
            '" measurementUnit=',
            '"\n            measurementUnit='
        ).replace(
            '" marginTop=',
            '"\n            marginTop='
        ).replace(
            '" marginLeft=',
            '"\n            marginLeft='
        ).replace(
            '" marginRight=',
            '"\n            marginRight='
        ).replace(
            '" marginBottom=',
            '"\n            marginBottom='
        )
        
        # Ensure source-file is single line
        result = result.replace('</source-file>', ' />')
        result = result.replace('<source-file ', '<source-file ')
        
        # Ensure other self-closing tags are formatted correctly
        for tag in ['font', 'document-heading', 'pdf-import', 'audit-import']:
            result = result.replace(f'</{tag}>', ' />')
        
        return result


def interactive_mode():
    """Interactive mode"""
    print("=" * 70)
    print("  根据Excel生成Batch List XML工具")
    print("  Generate Batch List XML from Excel")
    print("=" * 70)
    print()
    
    generator = XMLGenerator()
    
    # 1. Select input file
    print("Step 1/6: Select Input File")
    print("-" * 70)
    
    if HAS_TKINTER:
        # Use graphical window for selection
        try:
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            
            print("Please select Excel or CSV file in the popup window...")
            
            # Check if 02_output directory exists
            default_dir = Path.cwd() / "02_output"
            if not default_dir.exists():
                default_dir = Path.cwd()
            
            input_file = filedialog.askopenfilename(
                title="Select Excel/CSV File",
                filetypes=[
                    ("Excel and CSV Files", "*.xlsx *.xls *.csv"),
                    ("Excel Files", "*.xlsx *.xls"),
                    ("CSV Files", "*.csv"),
                    ("All Files", "*.*")
                ],
                initialdir=str(default_dir)
            )
            
            root.destroy()
            
            if not input_file:
                print("❌ No file selected, operation cancelled")
                return
            
            print(f"✓ Selected: {Path(input_file).name}")
            
        except Exception as e:
            print(f"⚠ Graphical window error: {e}")
            print("Switching to command line input mode...")
            HAS_TKINTER_TEMP = False
    else:
        HAS_TKINTER_TEMP = False
    
    # If no tkinter or graphical window failed, use command line input
    if not HAS_TKINTER or 'HAS_TKINTER_TEMP' in locals() and not HAS_TKINTER_TEMP:
        input_file = input("Enter Excel/CSV file path: ").strip().strip('"').strip("'")
        if not input_file:
            print("❌ File path cannot be empty, operation cancelled")
            return
    
    try:
        generator.load_excel(input_file)
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return
    
    # 2. Get header text
    print("\nStep 2/6: Set Header Text")
    print("-" * 70)
    print("This will be used for: <header text=\"...\"> and <document-heading text=\"...\">")
    default_header = "AZD0901 CSR DR2 Batch 1 Listings"
    header_text = input(f"Enter header text [Default: {default_header}]: ").strip()
    if not header_text:
        header_text = default_header
    print(f"✓ Using header: {header_text}")
    
    # 3. Get output filename (for PDF/audit filenames)
    print("\nStep 3/6: Set Output PDF Filename")
    print("-" * 70)
    print("IMPORTANT: Filename cannot contain spaces (use '_' instead)")
    print("This will be used for: <output-pdf> and <output-audit>")
    default_output_name = header_text.replace(' ', '_')
    while True:
        output_filename = input(f"Enter output filename [Default: {default_output_name}]: ").strip()
        if not output_filename:
            output_filename = default_output_name
        
        # Validate: no spaces allowed
        if ' ' in output_filename:
            print("❌ Error: Filename cannot contain spaces. Please use '_' instead.")
            print(f"   Example: {output_filename.replace(' ', '_')}")
            continue
        
        print(f"✓ Using output filename: {output_filename}")
        print(f"  - PDF: {output_filename}.pdf")
        print(f"  - Audit: {output_filename}_audit.pdf")
        break
    
    # 4. Get file location
    print("\nStep 4/6: Set File Location")
    print("-" * 70)
    default_location = "root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/"
    file_location = input(f"Enter file location [Default: {default_location}]: ").strip()
    if not file_location:
        file_location = default_location
    print(f"✓ Using file location: {file_location}")
    
    # 5. Select output path
    print("\nStep 5/6: Select Output Location")
    print("-" * 70)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    default_filename = f"_batch_list_{timestamp}.xml"
    default_output = f"03_xml/{default_filename}"
    
    if HAS_TKINTER:
        try:
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            
            print("Please select save location and filename in the popup window...")
            
            # Check if 03_xml directory exists
            xml_dir = Path.cwd() / "03_xml"
            if not xml_dir.exists():
                xml_dir = Path.cwd()
            
            output_path = filedialog.asksaveasfilename(
                title="Save XML File",
                defaultextension=".xml",
                filetypes=[
                    ("XML Files", "*.xml"),
                    ("All Files", "*.*")
                ],
                initialdir=str(xml_dir),
                initialfile=default_filename
            )
            
            root.destroy()
            
            if not output_path:
                print("❌ No save location selected, operation cancelled")
                return
            
            print(f"✓ Save path: {output_path}")
            
        except Exception as e:
            print(f"⚠ Graphical window error: {e}")
            print("Switching to command line input mode...")
            output_path = input(f"Enter output XML file path [Default: {default_output}]: ").strip()
            if not output_path:
                output_path = default_output
            print(f"✓ Save path: {output_path}")
    else:
        output_path = input(f"Enter output XML file path [Default: {default_output}]: ").strip()
        if not output_path:
            output_path = default_output
        print(f"✓ Save path: {output_path}")
    
    # 6. Get start number
    print("\nStep 6/6: Set Starting Page Number")
    print("-" * 70)
    start_number = input("Enter starting page number [Default: 2]: ").strip()
    if not start_number:
        start_number = 2
    else:
        try:
            start_number = int(start_number)
        except ValueError:
            print("⚠ Warning: Invalid page number, using default value 2")
            start_number = 2
    
    # 7. Generate XML
    print("\n" + "=" * 70)
    try:
        output_file = generator.generate_xml(
            header_text=header_text,
            file_location=file_location,
            output_path=output_path,
            output_filename=output_filename,
            start_number=start_number
        )
        print("=" * 70)
        print("✓ Operation completed successfully!")
        print(f"\nGenerated file: {output_file.absolute()}")
        
    except Exception as e:
        print(f"\n✗ Generation failed: {e}")
        import traceback
        traceback.print_exc()


def main():
    """Main function"""
    import sys
    
    try:
        if len(sys.argv) > 1:
            # Command line mode
            print("Command line mode not yet implemented, please use interactive mode")
            print("Run directly: python generate_batch_xml.py")
        else:
            # Interactive mode
            interactive_mode()
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user")
    except Exception as e:
        print(f"\n\n❌ Unexpected error occurred: {e}")
        import traceback
        traceback.print_exc()
        print("\nPlease check:")
        print("1. Are required packages installed: pip install pandas openpyxl")
        print("2. Is Python version >=3.6")
        print("3. Is input file format correct")


if __name__ == "__main__":
    main()
