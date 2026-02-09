"""
CSV export functionality for validated OCR results
"""
import csv
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import List
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))

from models.form_schema import DailyWardEntry


def export_to_csv(
    entries: List[DailyWardEntry],
    output_path: str,
    include_metadata: bool = True
) -> int:
    """
    Export validated entries to CSV file
    
    Args:
        entries: List of DailyWardEntry objects
        output_path: Path to output CSV file
        include_metadata: Include confidence/notes columns
    
    Returns:
        Number of entries exported
    """
    if not entries:
        raise ValueError("No entries to export")
    
    # Convert entries to CSV rows
    rows = []
    for entry in entries:
        row = entry.to_csv_row()
        rows.append(row)
    
    # Create DataFrame
    df = pd.DataFrame(rows)
    
    # Reorder columns for Excel import compatibility
    column_order = [
        'Ward', 'Date', 'Admissions', 'Discharges',
        'Deaths', 'DeathsU24', 'TransfersIn', 'TransfersOut'
    ]
    
    if include_metadata:
        column_order.extend(['Confidence', 'ReviewedBy', 'Notes'])
    
    # Ensure all columns exist
    for col in column_order:
        if col not in df.columns:
            df[col] = ''
    
    df = df[column_order]
    
    # Write to CSV
    df.to_csv(output_path, index=False, encoding='utf-8-sig')
    
    return len(entries)


def export_with_audit_log(
    entries: List[DailyWardEntry],
    output_path: str,
    audit_log_path: str = None
) -> dict:
    """
    Export entries with detailed audit log
    
    Args:
        entries: List of DailyWardEntry objects
        output_path: Path to output CSV file
        audit_log_path: Path to audit log (default: {output}_audit.txt)
    
    Returns:
        Export statistics dict
    """
    if audit_log_path is None:
        audit_log_path = str(Path(output_path).with_suffix('')) + '_audit.txt'
    
    # Export main CSV
    count = export_to_csv(entries, output_path, include_metadata=True)
    
    # Generate statistics
    stats = {
        'total_entries': count,
        'export_time': datetime.now().isoformat(),
        'high_confidence': 0,
        'medium_confidence': 0,
        'low_confidence': 0,
        'with_notes': 0,
        'reviewed': 0
    }
    
    for entry in entries:
        avg_conf = entry.get_average_confidence()
        if avg_conf >= 0.85:
            stats['high_confidence'] += 1
        elif avg_conf >= 0.70:
            stats['medium_confidence'] += 1
        else:
            stats['low_confidence'] += 1
        
        if entry.notes:
            stats['with_notes'] += 1
        if entry.reviewed_by:
            stats['reviewed'] += 1
    
    # Write audit log
    with open(audit_log_path, 'w', encoding='utf-8') as f:
        f.write("=" * 60 + "\n")
        f.write("OCR Export Audit Log\n")
        f.write("=" * 60 + "\n\n")
        
        f.write(f"Export Time: {stats['export_time']}\n")
        f.write(f"Output File: {output_path}\n\n")
        
        f.write(f"Total Entries: {stats['total_entries']}\n")
        f.write(f"Reviewed: {stats['reviewed']}\n")
        f.write(f"With Manual Notes: {stats['with_notes']}\n\n")
        
        f.write("Confidence Distribution:\n")
        f.write(f"  High (≥85%): {stats['high_confidence']}\n")
        f.write(f"  Medium (70-85%): {stats['medium_confidence']}\n")
        f.write(f"  Low (<70%): {stats['low_confidence']}\n\n")
        
        f.write("Entry Details:\n")
        f.write("-" * 60 + "\n")
        
        for i, entry in enumerate(entries, 1):
            f.write(f"\n{i}. {entry.ward_code} - {entry.entry_date}\n")
            f.write(f"   Avg Confidence: {entry.get_average_confidence():.1%}\n")
            
            low_fields = entry.get_low_confidence_fields()
            if low_fields:
                f.write(f"   Low Confidence Fields: {', '.join(low_fields)}\n")
            
            if entry.notes:
                f.write(f"   Notes: {entry.notes}\n")
    
    return stats


def create_import_template(output_path: str):
    """
    Create an empty CSV template for manual entry
    
    Args:
        output_path: Path to output template CSV
    """
    headers = [
        'Ward', 'Date', 'Admissions', 'Discharges',
        'Deaths', 'DeathsU24', 'TransfersIn', 'TransfersOut',
        'Confidence', 'ReviewedBy', 'Notes'
    ]
    
    # Create empty DataFrame with headers
    df = pd.DataFrame(columns=headers)
    
    # Add example row
    example = {
        'Ward': 'MW',
        'Date': '2026-01-15',
        'Admissions': 5,
        'Discharges': 3,
        'Deaths': 1,
        'DeathsU24': 0,
        'TransfersIn': 2,
        'TransfersOut': 1,
        'Confidence': 0.85,
        'ReviewedBy': 'StaffName',
        'Notes': 'Any corrections made'
    }
    
    df = pd.concat([df, pd.DataFrame([example])], ignore_index=True)
    
    # Write to CSV
    df.to_csv(output_path, index=False, encoding='utf-8-sig')


def merge_csv_files(input_files: List[str], output_path: str) -> int:
    """
    Merge multiple CSV export files into one
    
    Args:
        input_files: List of CSV file paths
        output_path: Path to merged output CSV
    
    Returns:
        Total number of entries in merged file
    """
    dfs = []
    for file_path in input_files:
        df = pd.read_csv(file_path, encoding='utf-8-sig')
        dfs.append(df)
    
    merged = pd.concat(dfs, ignore_index=True)
    
    # Sort by Ward and Date
    merged = merged.sort_values(['Ward', 'Date'])
    
    # Write merged CSV
    merged.to_csv(output_path, index=False, encoding='utf-8-sig')
    
    return len(merged)
