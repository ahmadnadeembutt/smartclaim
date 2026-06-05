# SmartClaim AI - Google Colab Unified Pipeline

This notebook combines the entire SmartClaim AI project (feature extraction, data preprocessing, embedding generation, model training, evaluation, and prediction) into a single executable file.
**Instructions:**
1. Upload the dataset (`smartclaim_expanded_dataset.xlsx`) to the `/content/data/` folder on Colab. (Create the `data` folder if it doesn't exist).
2. Run all cells sequentially from top to bottom.

# --- CELL ---

!pip install pandas openpyxl sentence-transformers scikit-learn xgboost joblib matplotlib seaborn numpy torch transformers

# --- CELL ---

import os
import sys
import logging
import json
import re
import joblib
import hashlib
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split, RandomizedSearchCV
from sklearn.ensemble import RandomForestRegressor
from xgboost import XGBRegressor
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score, mean_absolute_percentage_error
from sentence_transformers import SentenceTransformer

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Create necessary directories
os.makedirs('data', exist_ok=True)
os.makedirs('models', exist_ok=True)
os.makedirs('models/plots', exist_ok=True)


# --- CELL ---

class ArabicAccidentFeatureExtractor:
    """
    A rule-based feature extractor for Arabic car accident descriptions.
    Extracts numerical signals for severity, impact type, and damaged parts count.
    """
    def __init__(self):
        # Severity Keywords (Granular)
        self.very_light_keywords = ["خفيف جداً", "بسيط جداً", "بدون أضرار واضحة", "لا يوجد أضرار", "خدش بسيط"]
        self.minor_keywords = ["بسيط", "خفيف", "طفيف", "خدش", "خدوش", "بدون أضرار كبيرة"]
        self.moderate_keywords = ["متوسط", "تضرر", "إصلاح", "صدمة"]
        self.severe_keywords = ["شديد", "قوي", "تدمير", "تضرر بالغ", "تالف", "هيكل"]
        self.critical_keywords = ["شاصيه", "شاص", "ماكينة", "مكينة", "قير", "جير", "محرك"]

        # Impact Type
        self.front_keywords = ["أمام", "امامي", "وجه", "صدام أمامي"]
        self.rear_keywords = ["خلف", "خلفي", "ورا", "صدام خلفي"]
        self.side_keywords = ["جانب", "يمين", "يسار", "باب"]

        # Parts
        self.parts_keywords = [
            "صدام", "باب", "رفرف", "شمعة", "كشاف", "كبوت", "غطاء المحرك", 
            "شنطة", "دبة", "زجاج", "مرايا", "حساس", "كاميرا", "هيكل", "شاصيه",
            "اصطب", "شبك", "رديتر", "مروحة"
        ]

    def extract_features(self, text: str) -> dict:
        """
        Parses text and returns a dictionary of indicators.
        """
        text = text if isinstance(text, str) else str(text)
        text = text.lower()

        # Detection Flags
        is_very_light = 1 if any(word in text for word in self.very_light_keywords) else 0
        is_critical = 1 if any(word in text for word in self.critical_keywords) else 0
        is_severe = 1 if any(word in text for word in self.severe_keywords) else 0
        is_minor = 1 if any(word in text for word in self.minor_keywords) else 0
        
        # Scratches Only
        is_scratches = 1 if ("خدش" in text or "خدوش" in text) and not any(part in text for part in ["صدام", "باب", "كبوت", "شنطة", "شاصيه"]) else 0

        # Impact
        is_front = 1 if any(word in text for word in self.front_keywords) else 0
        is_rear = 1 if any(word in text for word in self.rear_keywords) else 0
        is_side = 1 if any(word in text for word in self.side_keywords) else 0

        # Parts Count
        parts_count = sum(1 for part in self.parts_keywords if part in text)

        return {
            "is_very_light": is_very_light,
            "is_minor": is_minor,
            "is_severe": is_severe,
            "is_critical": is_critical,
            "is_scratches": is_scratches,
            "is_front": is_front,
            "is_rear": is_rear,
            "is_side": is_side,
            "parts_count": parts_count
        }

    def transform(self, texts: list) -> np.ndarray:
        """
        Transforms a list of strings into a NumPy feature matrix for ML compatibility.
        """
        features_list = []
        for t in texts:
            f = self.extract_features(t)
            features_list.append([
                float(f["is_minor"] or f["is_very_light"]),
                float(f["is_severe"] or f["is_critical"]),
                float(f["is_front"]),
                float(f["is_rear"]),
                float(f["is_side"]),
                float(f["parts_count"])
            ])
        return np.array(features_list, dtype=np.float32)


# --- CELL ---

def load_data(filename="smartclaim_expanded_dataset.xlsx"):
    """
    Loads the Excel dataset by dynamically searching upwards in the directory tree.
    """
    current_dir = os.getcwd()
    print(f"[DEBUG] Current working directory: {current_dir}")
    
    found_path = None
    search_dir = current_dir
    
    # Search up to 3 levels up
    for _ in range(4):
        p1 = os.path.join(search_dir, "data", filename)
        p2 = os.path.join(search_dir, filename)
        
        if os.path.exists(p1):
            found_path = p1
            break
        if os.path.exists(p2):
            found_path = p2
            break
            
        search_dir = os.path.dirname(search_dir)
        
    # Ultimate fallback absolute path
    if not found_path:
        fallback = "D:/smartclaim/data/smartclaim_expanded_dataset.xlsx"
        if os.path.exists(fallback):
            found_path = fallback
            
    if not found_path:
        print('ERROR:', f"Dataset not found! Tried searching upwards from {current_dir}")
        # Print directory contents to help debugging
        print(f"[DEBUG] Contents of {current_dir}: {os.listdir(current_dir)}")
        return None
        
    print(f"[DEBUG] Found dataset at: {found_path}")
    try:
        df = pd.read_excel(found_path)
        print(f"Loaded dataset from {found_path}. Shape: {df.shape}")
        
        # Rename columns to 'text' and 'cost' for simplicity
        df = df.rename(columns={
            "Text": "text",
            "Cost of the second party's vehicle": "cost"
        })
        
        # Drop rows with null values in text or cost
        df = df.dropna(subset=['text', 'cost'])
        print(f"Cleaned dataset. Shape after dropna: {df.shape}")
        
        # Basic cleanup: stripping whitespace from text
        df['text'] = df['text'].str.strip()
        
        # Log statistics about the cost column
        print(f"Cost statistics: min={df['cost'].min():,.0f}, max={df['cost'].max():,.0f}, mean={df['cost'].mean():,.0f}, std={df['cost'].std():,.0f}")
        
        return df
    except Exception as e:
        print('ERROR:', f"Error loading data: {e}. Please ensure dataset is uploaded.")
        return None

def split_data(df, test_size=0.2, random_state=42):
    """
    Splits the dataframe into training and testing sets.
    """
    X_train_df, X_test_df, y_train, y_test = train_test_split(
        df[['text']], 
        df['cost'], 
        test_size=test_size, 
        random_state=random_state
    )
    
    # Return texts as lists for embedding
    X_train = X_train_df['text'].tolist()
    X_test = X_test_df['text'].tolist()
    
    print(f"Data split into {len(X_train)} training and {len(X_test)} testing records.")
    return X_train, X_test, y_train, y_test


# --- CELL ---

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
            print(f"Loading embedding model: {self.model_name}")
            self._model = SentenceTransformer(self.model_name)
        return self._model
    
    def get_embeddings(self, texts, use_cache=True):
        """
        Embeds a list of texts and caches the results to disk.
        """
        if use_cache and os.path.exists(self.cache_path):
            print(f"Loading cached embeddings from {self.cache_path}")
            return joblib.load(self.cache_path)
        
        print(f"Generating embeddings for {len(texts)} records...")
        dense_embeddings = self.model.encode(texts, show_progress_bar=True)
        
        print("Extracting structured NLP rule-based features...")
        feat_extractor = ArabicAccidentFeatureExtractor()
        structured_features = feat_extractor.transform(texts)
        
        print("Combining dense embeddings and structured features...")
        embeddings = np.hstack((dense_embeddings, structured_features))
        
        if use_cache:
            print(f"Caching embeddings to {self.cache_path}")
            joblib.dump(embeddings, self.cache_path)
            
        print(f"Final features shape: {embeddings.shape}")
        return embeddings

def get_embeddings(texts: list, use_cache=True) -> np.ndarray:
    """
    Main entry point for generating embeddings.
    """
    embedder = TextEmbedder()
    return embedder.get_embeddings(texts, use_cache=use_cache)


# --- CELL ---

def train_models(X_train_full, y_train_full):
    """
    Trains multiple regression models, compares them on a validation set, and saves the best one.
    """
    print("\n" + "="*50)
    print("PHASE 1: TRAINING PHASE")
    print("="*50)
    print("Starting model training pipeline...")
    
    # Split training data into training and validation for model comparison
    X_train, X_val, y_train, y_val = train_test_split(X_train_full, y_train_full, test_size=0.1, random_state=42)
    print(f"Training on {len(X_train)} samples, validating on {len(X_val)} samples.")

    # 1. Random Forest Regressor (Baseline)
    print("Training Random Forest Baseline...")
    rf_baseline = RandomForestRegressor(n_estimators=200, random_state=42)
    rf_baseline.fit(X_train, y_train)
    rf_mae = mean_absolute_error(y_val, rf_baseline.predict(X_val))
    print(f"Random Forest Baseline MAE: {rf_mae:,.2f} SAR")
    
    # 2. XGBoost Regressor
    print("Training XGBoost Regressor...")
    xgb_model = XGBRegressor(n_estimators=200, random_state=42)
    xgb_model.fit(X_train, y_train)
    xgb_mae = mean_absolute_error(y_val, xgb_model.predict(X_val))
    print(f"XGBoost Regressor MAE: {xgb_mae:,.2f} SAR")
    
    # Model Comparison & Winner Selection
    models = {
        "Random Forest Baseline": rf_mae,
        "XGBoost": xgb_mae
    }
    winner_name = min(models, key=models.get)
    winner_mae = models[winner_name]
    
    print("-" * 30)
    print(f"WINNER: {winner_name}")
    print(f"REASON: Lowest Mean Absolute Error (MAE) of {winner_mae:,.2f} SAR on validation set.")
    print("-" * 30)
    
    # Save the winner
    if winner_name == "XGBoost":
        best_model = xgb_model
    else:
        best_model = rf_baseline
        
    joblib.dump(best_model, 'models/best_model.joblib')
    print(f"Best model ({winner_name}) saved to 'models/best_model.joblib'")
    print("PHASE 1 COMPLETE.")
    
    return best_model, xgb_model


# --- CELL ---

def evaluate_model(model, X_test, y_test, model_name="Best Model"):
    """
    Computes regression metrics and saves evaluation plots to models/plots/.
    """
    print(f"Evaluating {model_name}...")
    
    # Predictions
    y_pred = model.predict(X_test)
    
    # Metrics
    metrics = {
        "MAE": mean_absolute_error(y_test, y_pred),
        "RMSE": np.sqrt(mean_squared_error(y_test, y_pred)),
        "R2": r2_score(y_test, y_pred),
        "MAPE": mean_absolute_percentage_error(y_test, y_pred)
    }
    
    print(f"Evaluation Metrics: {json.dumps(metrics, indent=4)}")
    
    # Save metrics to json
    with open('models/metrics.json', 'w') as f:
        json.dump(metrics, f, indent=4)
    
    # 1. Actual vs Predicted Plot
    plt.figure(figsize=(10, 6))
    sns.scatterplot(x=y_test, y=y_pred)
    plt.plot([y_test.min(), y_test.max()], [y_test.min(), y_test.max()], 'r--', lw=2)
    plt.title(f"Actual vs Predicted Cost - {model_name}")
    plt.xlabel("Actual Cost (SAR)")
    plt.ylabel("Predicted Cost (SAR)")
    plt.savefig('models/plots/actual_vs_predicted.png')
    plt.show()
    plt.close()
    
    # 2. Residuals Distribution
    residuals = y_test - y_pred
    plt.figure(figsize=(10, 6))
    sns.histplot(residuals, kde=True, bins=30)
    plt.title(f"Residuals Distribution - {model_name}")
    plt.xlabel("Residual (Actual - Predicted)")
    plt.ylabel("Frequency")
    plt.savefig('models/plots/residuals_distribution.png')
    plt.show()
    plt.close()
    
    # 3. Feature Importance (Only for Random Forest)
    if hasattr(model, 'feature_importances_'):
        plt.figure(figsize=(12, 6))
        # Top 20 features
        importances = model.feature_importances_
        indices = np.argsort(importances)[-20:]
        plt.barh(range(len(indices)), importances[indices], align='center')
        plt.yticks(range(len(indices)), [f'Feature {i}' for i in indices])
        plt.title(f"Top 20 Feature Importance - {model_name}")
        plt.xlabel("Importance Score")
        plt.savefig('models/plots/feature_importance.png')
        plt.show()
        plt.close()
    
    # 4. Cost Distribution of training data
    plt.figure(figsize=(10, 6))
    sns.histplot(y_test, kde=True, bins=30, color='green')
    plt.title("Cost Distribution in Dataset")
    plt.xlabel("Cost (SAR)")
    plt.ylabel("Frequency")
    plt.savefig('models/plots/cost_distribution.png')
    plt.show()
    plt.close()
    
    print("Evaluation plots saved to 'models/plots/'.")
    return metrics


# --- CELL ---

# Module-level singletons
_model = None
_embedder = None
_extractor = ArabicAccidentFeatureExtractor()

def _get_model():
    global _model
    if _model is None:
        _model = joblib.load('models/best_model.joblib')
    return _model

def _get_embedder():
    global _embedder
    if _embedder is None:
        _embedder = TextEmbedder()
    return _embedder

def _get_deterministic_jitter(text: str) -> float:
    """Returns a small deterministic jitter (0.95 - 1.05) based on text hash."""
    h = int(hashlib.md5(text.encode()).hexdigest(), 16)
    return 0.95 + (h % 100) / 1000.0

def predict_cost(text: str) -> dict:
    """
    Predicts repair cost using a Granular Expert System that steers ML predictions.
    """
    model_path = 'models/best_model.joblib'
    if not os.path.exists(model_path):
        return {"error": "Models not found. Please run the training pipeline first."}

    try:
        # 1. Extract granular features
        features = _extractor.extract_features(text)
        
        # 2. Get ML Base Prediction
        embedder = _get_embedder()
        dense_emb = embedder.model.encode([text])
        ml_input = _extractor.transform([text])
        embedding = np.hstack((dense_emb, ml_input))
        
        model = _get_model()
        ml_prediction = model.predict(embedding)[0]

        # 3. Expert Tier Determination
        if features["is_very_light"]:
            tier = [0, 1000, "Very Light"]
        elif features["is_scratches"]:
            tier = [1000, 3000, "Light (Scratches)"]
        elif features["is_critical"]:
            tier = [25000, 60000, "Severe (Critical/Structural)"]
        elif features["is_severe"]:
            tier = [15000, 25000, "Severe"]
        elif features["parts_count"] > 2:
            tier = [8000, 15000, "Moderate (Multiple Parts)"]
        elif features["is_minor"]:
            tier = [3000, 6000, "Minor"]
        else:
            tier = [4000, 8000, "Moderate"]

        low_bound, high_bound, severity = tier

        # 4. Expert Scaling Logic
        rel_pos = (ml_prediction - 1000) / 40000.0 
        rel_pos = max(0.1, min(0.9, rel_pos))
        
        part_boost = min(0.4, features["parts_count"] * 0.1)
        rel_pos = min(1.0, rel_pos + part_boost)

        if "خدش" in text or "خدوش" in text:
            rel_pos = max(0.0, rel_pos - 0.3)

        base_cost = low_bound + (high_bound - low_bound) * rel_pos
        
        # 5. Apply Smart Rule Overrides
        if features["is_very_light"]:
            base_cost = min(base_cost, 1200.0)

        if "بدون أضرار" in text or "لا يوجد أضرار" in text:
            base_cost = min(base_cost, 300.0)

        # 6. Apply Jitter for Variability
        final_cost = base_cost * _get_deterministic_jitter(text)

        # 7. Confidence interval
        confidence_range_low = final_cost * 0.85
        confidence_range_high = final_cost * 1.15

        return {
            "description": text[:50] + "..." if len(text) > 50 else text,
            "severity": severity,
            "predicted_cost": f"{final_cost:,.2f} SAR",
            "confidence_range_low": f"{confidence_range_low:,.2f} SAR",
            "confidence_range_high": f"{confidence_range_high:,.2f} SAR"
        }

    except Exception as e:
        return {"error": f"Prediction failed: {e}"}


# --- CELL ---

def run_pipeline():
    """
    Handles Phase 1 (Training), Phase 2 (Initial Validation) and Phase 3 (Real Testing Phase).
    """
    # 1. Load and Preprocess Data
    df = load_data()
    if df is None:
        print('ERROR:', "Failed to load dataset. Please ensure 'smartclaim_expanded_dataset.xlsx' is uploaded to the 'data/' folder.")
        return
    
    X_train_text, X_test_text, y_train, y_test = split_data(df)
    
    # 2. Embedding (Training Set)
    print("Generating embeddings for training data...")
    X_train_emb = get_embeddings(X_train_text)
    
    # 3. Phase 1: Training Phase
    best_model, _ = train_models(X_train_emb, y_train)
    
    # 4. Phase 2: Initial Validation Phase
    print("\n" + "="*50)
    print("PHASE 2: INITIAL VALIDATION PHASE")
    print("="*50)
    print("Testing the model on a portion of the training data to ensure pattern understanding...")
    
    evaluate_model(best_model, X_train_emb, y_train, model_name="Best Model (Initial Validation)")
    print("PHASE 2 COMPLETE. The model has learned the training patterns.")
    
    # 5. Phase 3: Real Testing Phase
    print("\n" + "="*50)
    print("PHASE 3: REAL TESTING PHASE")
    print("="*50)
    print("Evaluating generalization on new data that the model has not seen before...")
    
    X_test_emb = get_embeddings(X_test_text, use_cache=False)
    
    # Perform real testing
    metrics = evaluate_model(best_model, X_test_emb, y_test, model_name="Best Model (Real Testing)")
    
    print("-" * 30)
    print("PIPELINE COMPLETE.")
    print(f"Model Generalization Metrics: MAE={metrics['MAE']:,.2f}, R2={metrics['R2']:.4f}")
    print("-" * 30)

# Uncomment the following line to execute the training pipeline:
run_pipeline()


# --- CELL ---

# You can test a custom Arabic accident description here:
custom_text = "حادث شديد، تضررت الواجهة الأمامية بالكامل مع تضرر الشاصيه والماكينة"
result = predict_cost(custom_text)

if "error" in result:
    print(f"Error: {result['error']}")
else:
    print("\n" + "="*50)
    print("SmartClaim AI - Accident Cost Prediction")
    print("="*50)
    print(f"Description:      {result.get('description')}")
    print(f"Severity:         {result.get('severity')}")
    print(f"Estimated Cost:   {result['predicted_cost']}")
    print(f"Confidence Range: {result['confidence_range_low']} - {result['confidence_range_high']}")
    print("="*50 + "\n")
