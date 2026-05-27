import os
import joblib
import logging
import numpy as np
from sentence_transformers import SentenceTransformer

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class TextEmbedder:
    """
    Handles text embedding using paraphrase-multilingual-mpnet-base-v2 and caching.
    """
    def __init__(self, model_name="paraphrase-multilingual-mpnet-base-v2", cache_path="models/embeddings_cache.joblib"):
        self.model_name = model_name
        self.cache_path = cache_path
        self._model = None
        
    @property
    def model(self):
        if self._model is None:
            logging.info(f"Loading embedding model: {self.model_name}")
            self._model = SentenceTransformer(self.model_name)
        return self._model
    
    def get_embeddings(self, texts, use_cache=True):
        """
        Embeds a list of texts and caches the results to disk.
        """
        from src.feature_extractor import ArabicAccidentFeatureExtractor
        if use_cache and os.path.exists(self.cache_path):
            logging.info(f"Loading cached embeddings from {self.cache_path}")
            return joblib.load(self.cache_path)
        
        logging.info(f"Generating embeddings for {len(texts)} records...")
        dense_embeddings = self.model.encode(texts, show_progress_bar=True)
        
        logging.info("Extracting structured NLP rule-based features...")
        feat_extractor = ArabicAccidentFeatureExtractor()
        structured_features = feat_extractor.transform(texts)
        
        logging.info("Combining dense embeddings and structured features...")
        embeddings = np.hstack((dense_embeddings, structured_features))
        
        if use_cache:
            logging.info(f"Caching embeddings to {self.cache_path}")
            joblib.dump(embeddings, self.cache_path)
            
        logging.info(f"Final features shape: {embeddings.shape}")
        return embeddings

def get_embeddings(texts: list, use_cache=True) -> np.ndarray:
    """
    Main entry point for generating embeddings.
    """
    embedder = TextEmbedder()
    return embedder.get_embeddings(texts, use_cache=use_cache)

if __name__ == "__main__":
    test_texts = ["صدمة في الباب الأمامي", "هذا حادث بسيط"]
    embeddings = get_embeddings(test_texts)
    print(f"Sample embeddings shape: {embeddings.shape}")
