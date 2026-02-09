"""
Simple test script to demonstrate OCR pipeline
Tests preprocessing + TrOCR on actual form images
"""
import sys
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from ocr_tool.preprocessing.enhance import preprocess_image
from ocr_tool.extraction.trocr_engine import get_engine
import cv2
import numpy as np


def test_ocr_on_image(image_path: str):
    """Test OCR pipeline on a single image"""
    print(f"\n{'='*60}")
    print(f"Testing OCR on: {image_path}")
    print('='*60)
    
    # Step 1: Preprocess image
    print("\n[1/3] Preprocessing image...")
    try:
        preprocessed = preprocess_image(image_path, debug=False)
        print(f"  ✓ Preprocessed: {preprocessed.shape}")
    except Exception as e:
        print(f"  ✗ Preprocessing failed: {e}")
        return
    
    # Step 2: Extract text from a sample region (top 20% of image - header)
    print("\n[2/3] Extracting text from header region...")
    try:
        h, w = preprocessed.shape
        header_region = preprocessed[0:int(h*0.2), :]
        
        engine = get_engine()
        text, confidence = engine.extract_text(header_region)
        
        print(f"  ✓ Extracted text: '{text}'")
        print(f"  ✓ Confidence: {confidence:.1%}")
    except Exception as e:
        print(f"  ✗ Text extraction failed: {e}")
        return
    
    # Step 3: Try extracting from multiple regions
    print("\n[3/3] Extracting from multiple regions...")
    try:
        # Divide image into 5 horizontal strips
        strip_height = h // 5
        regions = []
        for i in range(5):
            start = i * strip_height
            end = start + strip_height if i < 4 else h
            regions.append(preprocessed[start:end, :])
        
        results = engine.extract_text_batch(regions)
        
        for i, (text, conf) in enumerate(results, 1):
            print(f"  Region {i}: '{text[:50]}...' (conf: {conf:.1%})")
    
    except Exception as e:
        print(f"  ✗ Batch extraction failed: {e}")
        return
    
    print(f"\n{'='*60}")
    print("Test complete!")
    print('='*60)


if __name__ == '__main__':
    # Test on available images
    images_dir = Path(__file__).parent.parent / 'images'
    
    print("\nOCR Pipeline Test Script")
    print("This will test preprocessing + TrOCR extraction on your form images")
    print("\nNOTE: First run will download TrOCR model (~500MB)")
    print("Subsequent runs will be much faster!")
    
    # Find image files
    image_files = list(images_dir.glob('*.jpg')) + list(images_dir.glob('*.png'))
    
    if not image_files:
        print(f"\nNo images found in {images_dir}")
        print("Please add some form images and try again.")
        sys.exit(1)
    
    print(f"\nFound {len(image_files)} image(s)")
    
    # Test on first image
    test_image = image_files[0]
    print(f"\nTesting on: {test_image.name}")
    
    test_ocr_on_image(str(test_image))
