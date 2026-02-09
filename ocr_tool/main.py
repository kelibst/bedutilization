"""
OCR Tool for Bed Utilization Forms - Main Entry Point

Usage:
    python -m ocr_tool.main [image_files...]
    
Example:
    python -m ocr_tool.main sample_data/ward_forms/form1.jpg form2.jpg
"""
import sys
import argparse
from pathlib import Path


def main():
    """Main entry point for OCR tool"""
    parser = argparse.ArgumentParser(
        description='Extract data from handwritten Daily Ward State forms',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process single image
  python -m ocr_tool.main ward_form.jpg
  
  # Process multiple images
  python -m ocr_tool.main form1.jpg form2.jpg form3.jpg
  
  # Process all images in directory
  python -m ocr_tool.main sample_data/ward_forms/*.jpg
        """
    )
    
    parser.add_argument('images', nargs='+', help='Image file(s) to process')
    parser.add_argument('--output', '-o', default='ocr_output.csv', 
                       help='Output CSV file (default: ocr_output.csv)')
    parser.add_argument('--debug', action='store_true',
                       help='Save debug images showing preprocessing steps')
    parser.add_argument('--no-review', action='store_true',
                       help='Skip review interface (auto-export)')
    
    args = parser.parse_args()
    
    print("=" * 60)
    print("  OCR Tool for Bed Utilization Forms")
    print("  Ghana Health Service - Hohoe Municipal Hospital")
    print("=" * 60)
    print()
    
    # Validate image files
    image_paths = []
    for img_path in args.images:
        path = Path(img_path)
        if not path.exists():
            print(f"ERROR: File not found: {img_path}")
            sys.exit(1)
        if not path.suffix.lower() in ['.jpg', '.jpeg', '.png', '.tiff', '.bmp']:
            print(f"WARNING: {img_path} may not be a valid image file")
        image_paths.append(path)
    
    print(f"Found {len(image_paths)} image(s) to process")
    print()
    
    # TODO: Implement OCR processing pipeline
    print("TODO: OCR processing not yet implemented")
    print()
    print("Next steps:")
    print("  1. Implement TrOCR extraction engine (extraction/trocr_engine.py)")
    print("  2. Implement review UI (ui/review_window.py)")
    print("  3. Implement CSV export (export/csv_export.py)")
    print()
    print("For now, you can test image preprocessing:")
    print(f"  python -c \"from ocr_tool.preprocessing.enhance import preprocess_image; preprocess_image('{image_paths[0]}', debug=True)\"")


if __name__ == '__main__':
    main()
