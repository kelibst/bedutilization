"""
TrOCR-based text extraction engine
Uses Microsoft's TrOCR transformer model for handwriting recognition
"""
import torch
from transformers import TrOCRProcessor, VisionEncoderDecoderModel
from PIL import Image
import numpy as np
from typing import Tuple, Optional
import logging

logger = logging.getLogger(__name__)


class TrOCREngine:
    """
    Wrapper for TrOCR model with lazy loading and caching
    """
    
    def __init__(self, model_name: str = "microsoft/trocr-large-handwritten"):
        """
        Initialize TrOCR engine
        
        Args:
            model_name: HuggingFace model identifier
        """
        self.model_name = model_name
        self._processor = None
        self._model = None
        self._device = None
        
    def _load_model(self):
        """Lazy load model on first use"""
        if self._model is not None:
            return
        
        logger.info(f"Loading TrOCR model: {self.model_name}")
        logger.info("This may take a few minutes on first run (downloading ~500MB)...")
        
        # Load processor and model
        self._processor = TrOCRProcessor.from_pretrained(self.model_name)
        self._model = VisionEncoderDecoderModel.from_pretrained(self.model_name)
        
        # Use GPU if available
        self._device = "cuda" if torch.cuda.is_available() else "cpu"
        self._model = self._model.to(self._device)
        
        logger.info(f"Model loaded successfully (device: {self._device})")
    
    def extract_text(self, image: np.ndarray, return_confidence: bool = True) -> Tuple[str, float]:
        """
        Extract text from image region using TrOCR
        
        Args:
            image: Grayscale or color image as numpy array
            return_confidence: Whether to calculate confidence score
        
        Returns:
            (extracted_text, confidence_score)
        """
        # Ensure model is loaded
        self._load_model()
        
        # Convert numpy array to PIL Image
        if len(image.shape) == 2:
            # Grayscale
            pil_image = Image.fromarray(image).convert('RGB')
        else:
            pil_image = Image.fromarray(image)
        
        # Preprocess image
        pixel_values = self._processor(pil_image, return_tensors="pt").pixel_values
        pixel_values = pixel_values.to(self._device)
        
        # Generate text
        with torch.no_grad():
            generated_ids = self._model.generate(
                pixel_values,
                max_length=64,
                num_beams=4,
                early_stopping=True,
                return_dict_in_generate=True,
                output_scores=True
            )
        
        # Decode text
        text = self._processor.batch_decode(generated_ids.sequences, skip_special_tokens=True)[0]
        
        # Calculate confidence if requested
        confidence = 1.0
        if return_confidence and hasattr(generated_ids, 'sequences_scores'):
            # Use sequence score as confidence (convert from log probability)
            if len(generated_ids.sequences_scores) > 0:
                log_prob = generated_ids.sequences_scores[0].item()
                confidence = float(np.exp(log_prob))
                # Clamp to [0, 1]
                confidence = max(0.0, min(1.0, confidence))
        
        return text.strip(), confidence
    
    def extract_text_batch(self, images: list[np.ndarray]) -> list[Tuple[str, float]]:
        """
        Extract text from multiple images in batch (faster)
        
        Args:
            images: List of images as numpy arrays
        
        Returns:
            List of (text, confidence) tuples
        """
        if not images:
            return []
        
        # Ensure model is loaded
        self._load_model()
        
        # Convert to PIL images
        pil_images = []
        for img in images:
            if len(img.shape) == 2:
                pil_images.append(Image.fromarray(img).convert('RGB'))
            else:
                pil_images.append(Image.fromarray(img))
        
        # Preprocess all images
        pixel_values = self._processor(pil_images, return_tensors="pt").pixel_values
        pixel_values = pixel_values.to(self._device)
        
        # Generate text for all images
        with torch.no_grad():
            generated_ids = self._model.generate(
                pixel_values,
                max_length=64,
                num_beams=4,
                early_stopping=True
            )
        
        # Decode all texts
        texts = self._processor.batch_decode(generated_ids, skip_special_tokens=True)
        
        # Return with default confidence (batch processing doesn't return scores easily)
        return [(text.strip(), 0.85) for text in texts]
    
    def unload_model(self):
        """Free GPU/CPU memory by unloading model"""
        if self._model is not None:
            del self._model
            del self._processor
            self._model = None
            self._processor = None
            
            if torch.cuda.is_available():
                torch.cuda.empty_cache()
            
            logger.info("Model unloaded from memory")


# Global engine instance (singleton pattern)
_global_engine: Optional[TrOCREngine] = None


def get_engine() -> TrOCREngine:
    """Get or create global TrOCR engine instance"""
    global _global_engine
    if _global_engine is None:
        _global_engine = TrOCREngine()
    return _global_engine


def extract_text_from_region(image: np.ndarray) -> Tuple[str, float]:
    """
    Convenience function to extract text from image region
    
    Args:
        image: Image region as numpy array
    
    Returns:
        (text, confidence)
    """
    engine = get_engine()
    return engine.extract_text(image)
