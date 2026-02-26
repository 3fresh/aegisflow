#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Cleanup Script - Remove unused files and folders
"""
import os
import shutil

def main():
    work_dir = r"c:\Users\kplp794\OneDrive - AZCollaboration\Desktop\roooooot\00-工具开发\credit_latest\03_mastertoc"
    os.chdir(work_dir)
    
    print("=" * 50)
    print("Cleaning up unused files and folders...")
    print("=" * 50)
    print()
    
    folders_to_delete = ["__pycache__", "archive"]
    
    for folder in folders_to_delete:
        if os.path.exists(folder):
            try:
                shutil.rmtree(folder)
                print(f"[✓] Deleted {folder}/")
            except Exception as e:
                print(f"[✗] Failed to delete {folder}/: {e}")
        else:
            print(f"[!] {folder}/ not found (already deleted)")
    
    print()
    print("=" * 50)
    print("Cleanup completed!")
    print("=" * 50)
    print()
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
