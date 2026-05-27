import logging
import joblib
import numpy as np
from sklearn.ensemble import RandomForestRegressor
from xgboost import XGBRegressor
from sklearn.model_selection import RandomizedSearchCV, train_test_split
from sklearn.metrics import mean_absolute_error

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def train_models(X_train_full, y_train_full):
    """
    Trains multiple regression models, compares them on a validation set, and saves the best one.
    """
    logging.info("\n" + "="*50)
    logging.info("PHASE 1: TRAINING PHASE")
    logging.info("="*50)
    logging.info("Starting model training pipeline...")
    
    # Split training data into training and validation for model comparison
    X_train, X_val, y_train, y_val = train_test_split(X_train_full, y_train_full, test_size=0.1, random_state=42)
    logging.info(f"Training on {len(X_train)} samples, validating on {len(X_val)} samples.")

    # 1. Random Forest Regressor (Baseline)
    logging.info("Training Random Forest Baseline...")
    rf_baseline = RandomForestRegressor(n_estimators=200, random_state=42)
    rf_baseline.fit(X_train, y_train)
    rf_mae = mean_absolute_error(y_val, rf_baseline.predict(X_val))
    logging.info(f"Random Forest Baseline MAE: {rf_mae:,.2f} SAR")
    
    # 2. XGBoost Regressor
    logging.info("Training XGBoost Regressor...")
    xgb_model = XGBRegressor(n_estimators=200, random_state=42)
    xgb_model.fit(X_train, y_train)
    xgb_mae = mean_absolute_error(y_val, xgb_model.predict(X_val))
    logging.info(f"XGBoost Regressor MAE: {xgb_mae:,.2f} SAR")
    
    # 3. Hyperparameter Tuning for Random Forest
    logging.info("Starting hyperparameter tuning for Random Forest...")
    param_dist = {
        'n_estimators': [100, 200, 300, 500],
        'max_depth': [None, 10, 20, 30],
        'min_samples_split': [2, 5, 10],
        'min_samples_leaf': [1, 2, 4]
    }
    
    rf_tuned_search = RandomizedSearchCV(
        RandomForestRegressor(random_state=42), 
        param_distributions=param_dist, 
        n_iter=10, # Reduced for speed, can be 20
        cv=3, 
        scoring='neg_mean_absolute_error', 
        random_state=42, 
        n_jobs=-1
    )
    rf_tuned_search.fit(X_train, y_train)
    best_rf = rf_tuned_search.best_estimator_
    tuned_rf_mae = mean_absolute_error(y_val, best_rf.predict(X_val))
    logging.info(f"Tuned Random Forest MAE: {tuned_rf_mae:,.2f} SAR")
    
    # Model Comparison & Winner Selection
    models = {
        "Random Forest Baseline": rf_mae,
        "XGBoost": xgb_mae,
        "Tuned Random Forest": tuned_rf_mae
    }
    winner_name = min(models, key=models.get)
    winner_mae = models[winner_name]
    
    logging.info("-" * 30)
    logging.info(f"WINNER: {winner_name}")
    logging.info(f"REASON: Lowest Mean Absolute Error (MAE) of {winner_mae:,.2f} SAR on validation set.")
    logging.info("-" * 30)
    
    # Save the winner
    if winner_name == "XGBoost":
        best_model = xgb_model
    elif winner_name == "Random Forest Baseline":
        best_model = rf_baseline
    else:
        best_model = best_rf
        
    joblib.dump(best_model, 'models/best_model.joblib')
    logging.info(f"Best model ({winner_name}) saved to 'models/best_model.joblib'")
    logging.info("PHASE 1 COMPLETE.")
    
    return best_model, xgb_model

if __name__ == "__main__":
    X = np.random.rand(100, 768)
    y = np.random.randint(1000, 50000, 100)
    train_models(X, y)
