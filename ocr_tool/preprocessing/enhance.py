"""
Image preprocessing pipeline for OCR
Enhances image quality for better text recognition
"""
import cv2
import numpy as np
from typing import Tuple, Optional


def preprocess_image(image_path: str, debug: bool = False) -> np.ndarray:
    """
    Main preprocessing pipeline: applies all enhancement steps
    
    Args:
        image_path: Path to input image
        debug: If True, saves intermediate steps
    
    Returns:
        Enhanced grayscale image ready for OCR
    """
    # Load image
    img = cv2.imread(image_path)
    if img is None:
        raise ValueError(f"Cannot read image: {image_path}")
    
    # Convert to grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Apply preprocessing steps
    deskewed = deskew_image(gray)
    denoised = denoise_image(deskewed)
    enhanced = enhance_contrast(denoised)
    binary = binarize_image(enhanced)
    
    if debug:
        # Save intermediate steps for debugging
        import os
        base_name = os.path.splitext(os.path.basename(image_path))[0]
        debug_dir = os.path.join(os.path.dirname(image_path), 'debug')
        os.makedirs(debug_dir, exist_ok=True)
        
        cv2.imwrite(os.path.join(debug_dir, f"{base_name}_1_gray.png"), gray)
        cv2.imwrite(os.path.join(debug_dir, f"{base_name}_2_deskewed.png"), deskewed)
        cv2.imwrite(os.path.join(debug_dir, f"{base_name}_3_denoised.png"), denoised)
        cv2.imwrite(os.path.join(debug_dir, f"{base_name}_4_enhanced.png"), enhanced)
        cv2.imwrite(os.path.join(debug_dir, f"{base_name}_5_binary.png"), binary)
    
    return binary


def deskew_image(image: np.ndarray) -> np.ndarray:
    """
    Detect and correct image rotation/skew
    
    Args:
        image: Grayscale input image
    
    Returns:
        Rotated image
    """
    # Detect skew angle
    angle = detect_skew_angle(image)
    
    if abs(angle) < 0.5:
        # Angle too small, skip rotation
        return image
    
    # Rotate image
    (h, w) = image.shape[:2]
    center = (w // 2, h // 2)
    M = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(image, M, (w, h),
                             flags=cv2.INTER_CUBIC,
                             borderMode=cv2.BORDER_REPLICATE)
    
    return rotated


def detect_skew_angle(image: np.ndarray) -> float:
    """
    Detect skew angle using Hough transform
    
    Args:
        image: Grayscale image
    
    Returns:
        Skew angle in degrees
    """
    # Edge detection
    edges = cv2.Canny(image, 50, 150, apertureSize=3)
    
    # Hough line transform
    lines = cv2.HoughLines(edges, 1, np.pi / 180, 200)
    
    if lines is None:
        return 0.0
    
    # Calculate angles
    angles = []
    for rho, theta in lines[:, 0]:
        angle = (theta * 180 / np.pi) - 90
        angles.append(angle)
    
    if not angles:
        return 0.0
    
    # Return median angle
    return float(np.median(angles))


def denoise_image(image: np.ndarray) -> np.ndarray:
    """
    Remove noise from image
    
    Args:
        image: Grayscale input image
    
    Returns:
        Denoised image
    """
    # Non-local means denoising (good for handwriting)
    denoised = cv2.fastNlMeansDenoising(image, None, 
                                        h=10,  # Filter strength
                                        templateWindowSize=7,
                                        searchWindowSize=21)
    return denoised


def enhance_contrast(image: np.ndarray) -> np.ndarray:
    """
    Enhance image contrast using CLAHE
    (Contrast Limited Adaptive Histogram Equalization)
    
    Args:
        image: Grayscale input image
    
    Returns:
        Contrast-enhanced image
    """
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(image)
    return enhanced


def binarize_image(image: np.ndarray) -> np.ndarray:
    """
    Convert to binary (black/white) using adaptive thresholding
    
    Args:
        image: Grayscale input image
    
    Returns:
        Binary image
    """
    # Adaptive threshold works better for varying lighting
    binary = cv2.adaptiveThreshold(image, 255,
                                   cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY,
                                   11,  # Block size
                                   2)   # C constant
    return binary


def is_blurry(image: np.ndarray, threshold: float = 100.0) -> Tuple[bool, float]:
    """
    Detect if image is too blurry using Laplacian variance
    
    Args:
        image: Grayscale input image
        threshold: Variance threshold (lower = more blurry)
    
    Returns:
        (is_blurry, variance_score)
    """
    laplacian = cv2.Laplacian(image, cv2.CV_64F)
    variance = laplacian.var()
    
    return (variance < threshold, float(variance))


def resize_if_needed(image: np.ndarray, max_dimension: int = 3000) -> np.ndarray:
    """
    Resize image if it's too large (for faster processing)
    
    Args:
        image: Input image
        max_dimension: Maximum width or height
    
    Returns:
        Resized image (or original if small enough)
    """
    h, w = image.shape[:2]
    
    if max(h, w) <= max_dimension:
        return image
    
    # Calculate scaling factor
    if h > w:
        scale = max_dimension / h
    else:
        scale = max_dimension / w
    
    new_w = int(w * scale)
    new_h = int(h * scale)
    
    resized = cv2.resize(image, (new_w, new_h), interpolation=cv2.INTER_AREA)
    return resized
