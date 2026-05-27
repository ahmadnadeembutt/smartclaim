import argparse
import sys
import logging
import joblib
from src.preprocess import load_data, split_data
from src.embed import get_embeddings
from src.train import train_models
from src.evaluate import evaluate_model
from src.predict import predict_cost

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def run_train():
    """
    Handles Phase 1 (Training) and Phase 2 (Initial Validation).
    """
    # 1. Load and Preprocess Data
    df = load_data()
    if df is None:
        logging.error("Failed to load dataset.")
        return
    
    X_train_text, X_test_text, y_train, y_test = split_data(df)
    
    # 2. Embedding (Training Set)
    logging.info("Generating embeddings for training data...")
    X_train_emb = get_embeddings(X_train_text)
    
    # 3. Phase 1: Training Phase (Called within train_models)
    best_model, _ = train_models(X_train_emb, y_train)
    
    # 4. Phase 2: Initial Validation Phase
    logging.info("\n" + "="*50)
    logging.info("PHASE 2: INITIAL VALIDATION PHASE")
    logging.info("="*50)
    logging.info("Testing the model on a portion of the training data to ensure pattern understanding...")
    
    # Evaluate best model on the training embeddings (Self-Evaluation/Initial Validation)
    evaluate_model(best_model, X_train_emb, y_train, model_name="Best Model (Initial Validation)")
    
    logging.info("PHASE 2 COMPLETE. The model has learned the training patterns.")
    logging.info("Training pipeline finished successfully.")

def run_evaluate():
    """
    Handles Phase 3 (Real Testing Phase).
    """
    logging.info("\n" + "="*50)
    logging.info("PHASE 3: REAL TESTING PHASE")
    logging.info("="*50)
    logging.info("Evaluating generalization on new data that the model has not seen before...")
    
    # 1. Load Data
    df = load_data()
    if df is None:
        logging.error("Failed to load dataset.")
        return
        
    _, X_test_text, _, y_test = split_data(df)
    
    # 2. Embedding (Test Set)
    # We don't use cache for the test set here to ensure we embed the 200 test records,
    # as the current cache only contains the 800 training records.
    X_test_emb = get_embeddings(X_test_text, use_cache=False)
    
    # 3. Load Best Model
    try:
        model = joblib.load('models/best_model.joblib')
        
        # Perform real testing
        metrics = evaluate_model(model, X_test_emb, y_test, model_name="Best Model (Real Testing)")
        
        logging.info("-" * 30)
        logging.info("PHASE 3 COMPLETE.")
        logging.info(f"Model Generalization Metrics: MAE={metrics['MAE']:,.2f}, R2={metrics['R2']:.4f}")
        logging.info("-" * 30)
        
    except FileNotFoundError:
        logging.error("Best model not found. Please run 'python main.py train' first.")

def run_predict(text):
    """
    Runs a single prediction on input text.
    """
    result = predict_cost(text)
    if "error" in result:
        print(f"Error: {result['error']}")
    else:
        print("\n" + "="*50)
        print("SmartClaim AI - Accident Cost Prediction")
        print("="*50)
        print(f"Description:      {result.get('description', text)}")
        print(f"Severity:         {result.get('severity', 'N/A')}")
        print(f"Estimated Cost:   {result['predicted_cost']}")
        print(f"Confidence Range: {result['confidence_range_low']} - {result['confidence_range_high']}")
        print("="*50 + "\n")

def main():
    parser = argparse.ArgumentParser(description="SmartClaim AI - Saudi Car Accident Cost Prediction")
    subparsers = parser.add_subparsers(dest="command", help="Available commands")
    
    # Train command
    subparsers.add_parser("train", help="Run Phase 1 (Training) and Phase 2 (Initial Validation)")
    
    # Evaluate command
    subparsers.add_parser("evaluate", help="Run Phase 3 (Real Testing)")
    
    # Predict command
    predict_parser = subparsers.add_parser("predict", help="Predict cost for Arabic text")
    predict_parser.add_argument("--text", type=str, required=True, help="Arabic accident report text")
    
    args = parser.parse_args()
    
    if args.command == "train":
        run_train()
    elif args.command == "evaluate":
        run_evaluate()
    elif args.command == "predict":
        run_predict(args.text)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
