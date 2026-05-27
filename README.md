# SmartClaim AI — Saudi Car Accident Cost Prediction

SmartClaim AI is an end-to-end NLP regression pipeline designed for the Saudi Arabian insurance industry. It takes Arabic car accident reports as input and predicts the vehicle repair/damage cost in Saudi Riyals (SAR).

## 🚀 Key Features
- **Arabic NLP**: Optimized for Saudi legal and everyday Arabic accident reports.
- **Granular Expert Estimation**: A 7-tier rule-based system that steers ML predictions to ensure realistic Saudi market pricing (0 - 60,000+ SAR).
- **Anti-Overestimation Logic**: Specifically calibrated to prevent "bucketing" and ensure light impacts are priced accurately.
- **Deep Learning Embeddings**: Powered by `paraphrase-multilingual-mpnet-base-v2`.
- **Hybrid Modeling**: Uses Random Forest and XGBoost for robust predictions.
- **Confidence Estimation**: Provides a ±15% confidence interval for every prediction.

## 🏗️ Architecture

```text
[ Input Arabic Text ]
       |
       v
[ Granular Feature Extractor ] (Arabic keywords & parts count)
       |
       v
[ sentence-transformers ] (MPNet Embedding)
       |
       v
[ 7-Tier Expert Logic ] (Mapping to sub-ranges: Very Light -> Critical)
       |
       v
[ Final Estimated Cost (SAR) ]
```

## 📁 Project Structure

```text
smartclaim/
├── data/
│   └── smartclaim_expanded_dataset.xlsx  # Saudi claims dataset
├── models/
│   ├── plots/                            # Performance charts
│   └── best_model.joblib                 # Final production model
├── src/
│   ├── feature_extractor.py              # NEW: Expert keyword/part detection
│   ├── preprocess.py                     # Data cleaning & loading
│   ├── predict.py                        # REFACTORED: Expert 7-tier estimation
│   ├── train.py                          # ML training & tuning
│   └── evaluate.py                       # Evaluation & metrics
├── main.py                               # CLI Entry point
└── verify_tests.py                       # Batch verification test suite
```

## 🛠️ Installation

1. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## 📖 Usage

The project follows a **Three-Phase AI Workflow**:

### 1. Training & Initial Validation
Build the pipeline and check if the model learned the training patterns:
```bash
python main.py train
```

### 2. Real Generalization Testing (Phase 3)
Test the model on 200 held-out records it has **never seen before**:
```bash
python main.py evaluate
```

### 3. Batch Verification (Expert Rules)
Run a specific suite of cases to verify the **Granular Expert System** (e.g., checking if chassis damage is expensive and scratches are cheap):
```bash
python verify_tests.py
```

### 4. Single Prediction
Predict the cost for a custom report:
```bash
python main.py predict --text "حادث بسيط صدم من الخلف عند الإشارة"
```

## 📊 Model Performance (Phase 3 Final Results)

Based on the latest evaluation on the 20% hold-out set:

| Metric | Best Model (XGBoost/RF) | Status |
|---|---|---|
| **R² (R-Squared)** | **0.856** | ✅ Excellent Correlation |
| **MAE (Mean Absolute Error)** | **4,661 SAR** | ✅ Realistic for market |
| **MAPE (Percentage Error)** | **64.2%** | ⚠️ Skewed by low-value claims |

> [!TIP]
> The **Granular Expert System** implemented in `src/predict.py` directly compensates for high percentage errors in minor claims by enforcing mandatory pricing floors and ceilings based on detected keywords.

