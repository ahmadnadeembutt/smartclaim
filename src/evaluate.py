import logging
import json
import os
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score, mean_absolute_percentage_error

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def evaluate_model(model, X_test, y_test, model_name="Best Model"):
    """
    Computes regression metrics and saves evaluation plots to models/plots/.
    """
    logging.info(f"Evaluating {model_name}...")
    
    # Predictions
    y_pred = model.predict(X_test)
    
    # Metrics
    metrics = {
        "MAE": mean_absolute_error(y_test, y_pred),
        "RMSE": np.sqrt(mean_squared_error(y_test, y_pred)),
        "R2": r2_score(y_test, y_pred),
        "MAPE": mean_absolute_percentage_error(y_test, y_pred)
    }
    
    logging.info(f"Evaluation Metrics: {json.dumps(metrics, indent=4)}")
    
    # Save metrics to json
    with open('models/metrics.json', 'w') as f:
        json.dump(metrics, f, indent=4)
    
    # Ensure plots directory exists
    os.makedirs('models/plots', exist_ok=True)
    
    # 1. Actual vs Predicted Plot
    plt.figure(figsize=(10, 6))
    sns.scatterplot(x=y_test, y=y_pred)
    plt.plot([y_test.min(), y_test.max()], [y_test.min(), y_test.max()], 'r--', lw=2)
    plt.title(f"Actual vs Predicted Cost - {model_name}")
    plt.xlabel("Actual Cost (SAR)")
    plt.ylabel("Predicted Cost (SAR)")
    plt.savefig('models/plots/actual_vs_predicted.png')
    plt.close()
    
    # 2. Residuals Distribution
    residuals = y_test - y_pred
    plt.figure(figsize=(10, 6))
    sns.histplot(residuals, kde=True, bins=30)
    plt.title(f"Residuals Distribution - {model_name}")
    plt.xlabel("Residual (Actual - Predicted)")
    plt.ylabel("Frequency")
    plt.savefig('models/plots/residuals_distribution.png')
    plt.close()
    
    # 3. Feature Importance (Only for Random Forest)
    if hasattr(model, 'feature_importances_'):
        plt.figure(figsize=(12, 6))
        # Top 20 features (assuming 768 or similar length embeddings)
        importances = model.feature_importances_
        indices = np.argsort(importances)[-20:]
        plt.barh(range(len(indices)), importances[indices], align='center')
        plt.yticks(range(len(indices)), [f'Feature {i}' for i in indices])
        plt.title(f"Top 20 Feature Importance - {model_name}")
        plt.xlabel("Importance Score")
        plt.savefig('models/plots/feature_importance.png')
        plt.close()
    
    # 4. Cost Distribution of training data (using y_test as proxy for distribution)
    plt.figure(figsize=(10, 6))
    sns.histplot(y_test, kde=True, bins=30, color='green')
    plt.title("Cost Distribution in Dataset")
    plt.xlabel("Cost (SAR)")
    plt.ylabel("Frequency")
    plt.savefig('models/plots/cost_distribution.png')
    plt.close()
    
    logging.info("Evaluation plots saved to 'models/plots/'.")
    return metrics

if __name__ == "__main__":
    # Mock data for testing
    from sklearn.ensemble import RandomForestRegressor
    X = np.random.rand(100, 768)
    y = np.random.randint(1000, 50000, 100)
    model = RandomForestRegressor().fit(X, y)
    evaluate_model(model, X, y)
