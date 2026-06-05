import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Preformatted
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER

def create_pdf():
    pdf_path = "D:/smartclaim/data/SmartClaim_Pipeline_Explanation.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=letter, rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50)
    
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Justify', parent=styles['Normal'], alignment=TA_JUSTIFY, spaceAfter=12, fontSize=11, leading=15))
    styles.add(ParagraphStyle(name='CodeText', parent=styles['Code'], fontSize=9, leading=12, backColor='#f0f0f0', borderPadding=5))
    styles.add(ParagraphStyle(name='Heading1_Center', parent=styles['Heading1'], alignment=TA_CENTER, spaceAfter=20))
    
    Story = []
    
    # Title
    Story.append(Paragraph("SmartClaim AI: Pipeline Code Explanation", styles['Heading1_Center']))
    Story.append(Paragraph("A Detailed Guide for University Presentation", styles['Heading3']))
    Story.append(Spacer(1, 20))
    
    # 1. Overview
    Story.append(Paragraph("1. Overview of the Pipeline", styles['Heading2']))
    Story.append(Paragraph("The notebook <b>SmartClaim_Colab_Pipeline.ipynb</b> represents an end-to-end Machine Learning pipeline. It processes Arabic traffic accident descriptions and predicts realistic vehicle repair costs. The pipeline follows a 'Hybrid Expert-Steered ML' approach, meaning it combines statistical Machine Learning (like XGBoost or Random Forest) with human-encoded rules (Expert System) to prevent unrealistic cost predictions.", styles['Justify']))
    
    # 2. Arabic Feature Extraction
    Story.append(Paragraph("2. Arabic Feature Extraction Engine", styles['Heading2']))
    Story.append(Paragraph("The code defines a class called <code>ArabicAccidentFeatureExtractor</code>. Its purpose is to parse free-text Arabic accident descriptions and extract structured, numerical features.", styles['Justify']))
    Story.append(Paragraph("<b>How it works:</b>", styles['Justify']))
    Story.append(Paragraph("• It defines dictionaries of Arabic keywords categorized by Severity (e.g., Very Light, Minor, Critical), Impact Direction (Front, Rear, Side), and specific car parts (e.g., Door, Chassis, Bumper).", styles['Justify']))
    Story.append(Paragraph("• The <code>extract_features()</code> method scans the text and flags the presence of these keywords as 1 or 0 (binary flags), and counts the number of distinct parts damaged.", styles['Justify']))
    Story.append(Paragraph("• The <code>transform()</code> method converts these flags into a dense NumPy array that Machine Learning models can understand.", styles['Justify']))
    
    # 3. Embedding Generation
    Story.append(Paragraph("3. Semantic Text Embedding", styles['Heading2']))
    Story.append(Paragraph("Machine learning models cannot read text directly; they require numbers. The <code>TextEmbedder</code> class uses the HuggingFace <code>sentence-transformers</code> library, specifically the <b>paraphrase-multilingual-mpnet-base-v2</b> model.", styles['Justify']))
    Story.append(Paragraph("• <b>Dense Embeddings:</b> It transforms each Arabic sentence into a 768-dimensional dense vector representing its semantic meaning.", styles['Justify']))
    Story.append(Paragraph("• <b>Feature Concatenation:</b> It concatenates this 768-dimensional vector with the 6 structured features from our Arabic Feature Extractor, resulting in a rich 774-dimensional hybrid feature vector.", styles['Justify']))
    Story.append(Paragraph("• <b>Caching:</b> To save time in Colab, it caches embeddings to disk using <code>joblib</code>.", styles['Justify']))
    
    # 4. Data Loading and Preparation
    Story.append(Paragraph("4. Data Preparation", styles['Heading2']))
    Story.append(Paragraph("The <code>load_data()</code> function dynamically searches for the dataset <code>smartclaim_expanded_dataset.xlsx</code>. It uses the Pandas library to read the Excel file, renames columns to 'text' and 'cost', drops missing values, and strips whitespace. Then, <code>split_data()</code> divides the data into Training (80%) and Testing (20%) sets to ensure the model can be evaluated on unseen data.", styles['Justify']))
    
    # 5. Model Training
    Story.append(Paragraph("5. Model Training and Selection", styles['Heading2']))
    Story.append(Paragraph("The <code>train_models()</code> function implements a competitive training phase:", styles['Justify']))
    Story.append(Paragraph("• It takes the training set and splits it further (90% training, 10% internal validation).", styles['Justify']))
    Story.append(Paragraph("• It trains two robust models: <b>Random Forest Regressor</b> (an ensemble of decision trees) and <b>XGBoost Regressor</b> (a gradient boosting algorithm).", styles['Justify']))
    Story.append(Paragraph("• It compares both models using Mean Absolute Error (MAE) on the validation set. The model with the lower error is selected as the 'Winner' and saved to disk using <code>joblib.dump()</code>.", styles['Justify']))
    
    # 6. Model Evaluation
    Story.append(Paragraph("6. Model Evaluation", styles['Heading2']))
    Story.append(Paragraph("The <code>evaluate_model()</code> function tests the winning model on the 20% hold-out test set. It calculates metrics like R-Squared (R2), MAE, and RMSE. Crucially, it uses <code>matplotlib</code> and <code>seaborn</code> to generate diagnostic plots:", styles['Justify']))
    Story.append(Paragraph("• <b>Actual vs Predicted Plot:</b> Shows how closely predictions follow the ideal diagonal line.", styles['Justify']))
    Story.append(Paragraph("• <b>Residuals Distribution:</b> Checks if errors are normally distributed.", styles['Justify']))
    Story.append(Paragraph("• <b>Feature Importance:</b> Displays which variables had the most impact on the prediction.", styles['Justify']))
    
    # 7. The Expert System
    Story.append(Paragraph("7. The 7-Tier Expert Cost Estimation Engine", styles['Heading2']))
    Story.append(Paragraph("This is the most innovative part of the code. The <code>predict_cost()</code> function prevents the ML model from producing generic averages (a flaw known as 'bucketing bias').", styles['Justify']))
    Story.append(Paragraph("<b>The Logic:</b>", styles['Justify']))
    Story.append(Paragraph("1. <b>Tier Assignment:</b> Based on the rule-based extracted keywords, the accident is strictly assigned to one of 7 cost tiers (e.g., 'Very Light' -> 0-1,000 SAR, 'Critical' -> 25,000-60,000 SAR).", styles['Justify']))
    Story.append(Paragraph("2. <b>Relative Positioning:</b> The raw ML prediction is normalized into a ratio (0.1 to 0.9). This ratio decides where the final cost sits *within* the assigned tier bounds.", styles['Justify']))
    Story.append(Paragraph("3. <b>Penalties and Boosts:</b> If scratches are detected, the cost shifts lower in the tier. If many parts are damaged, the cost shifts higher.", styles['Justify']))
    Story.append(Paragraph("4. <b>Smart Overrides:</b> If the text explicitly says 'no damage', a hard cap is enforced.", styles['Justify']))
    
    # 8. Pipeline Execution
    Story.append(Paragraph("8. Full Pipeline Execution Flow", styles['Heading2']))
    Story.append(Paragraph("The <code>run_pipeline()</code> function ties everything together in three phases:", styles['Justify']))
    Story.append(Paragraph("• <b>Phase 1:</b> Train candidates and select the winner.", styles['Justify']))
    Story.append(Paragraph("• <b>Phase 2:</b> Validate the winner on training data to verify learning.", styles['Justify']))
    Story.append(Paragraph("• <b>Phase 3:</b> Evaluate strictly on the hold-out test data to prove real-world generalization capability.", styles['Justify']))
    
    Story.append(Spacer(1, 20))
    Story.append(Paragraph("<i>Generated by SmartClaim AI Assistant</i>", styles['Normal']))
    
    doc.build(Story)
    print(f"PDF successfully created at: {pdf_path}")

if __name__ == "__main__":
    create_pdf()
