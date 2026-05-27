import pandas as pd
import logging
from sklearn.model_selection import train_test_split

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def load_data(file_path="data/smartclaim_expanded_dataset.xlsx"):
    """
    Loads the Excel dataset and performs basic cleaning and renaming.
    """
    try:
        df = pd.read_excel(file_path)
        logging.info(f"Loaded dataset from {file_path}. Shape: {df.shape}")
        
        # Rename columns to 'text' and 'cost' for simplicity
        df = df.rename(columns={
            "Text": "text",
            "Cost of the second party's vehicle": "cost"
        })
        
        # Drop rows with null values in text or cost
        df = df.dropna(subset=['text', 'cost'])
        logging.info(f"Cleaned dataset. Shape after dropna: {df.shape}")
        
        # Basic cleanup: stripping whitespace from text
        df['text'] = df['text'].str.strip()
        
        # Log statistics about the cost column
        logging.info(f"Cost statistics: min={df['cost'].min():,.0f}, max={df['cost'].max():,.0f}, mean={df['cost'].mean():,.0f}, std={df['cost'].std():,.0f}")
        
        return df
    except Exception as e:
        logging.error(f"Error loading data: {e}")
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
    
    logging.info(f"Data split into {len(X_train)} training and {len(X_test)} testing records.")
    return X_train, X_test, y_train, y_test

if __name__ == "__main__":
    data_df = load_data()
    if data_df is not None:
        X_tr, X_te, y_tr, y_te = split_data(data_df)
