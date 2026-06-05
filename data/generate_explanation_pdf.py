import os
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.lib import colors

def create_explanation_pdf():
    pdf_path = "D:/smartclaim/data/SmartClaim_Code_Explanation.pdf"
    doc = SimpleDocTemplate(pdf_path, pagesize=letter, rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50)
    
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Justify', parent=styles['Normal'], alignment=TA_JUSTIFY, spaceAfter=12, fontSize=11, leading=16))
    styles.add(ParagraphStyle(name='CodeText', parent=styles['Code'], fontSize=10, leading=14, backColor='#f0f0f0', borderPadding=8, spaceAfter=12))
    styles.add(ParagraphStyle(name='Heading1_Center', parent=styles['Heading1'], alignment=TA_CENTER, spaceAfter=20, textColor=colors.darkblue))
    styles.add(ParagraphStyle(name='Heading2_Step', parent=styles['Heading2'], spaceAfter=10, textColor=colors.darkred))
    styles.add(ParagraphStyle(name='BulletItem', parent=styles['Normal'], fontSize=11, leading=16, leftIndent=20, spaceAfter=8))
    
    Story = []
    
    # Title
    Story.append(Paragraph("SmartClaim AI: Code and Commands Explanation", styles['Heading1_Center']))
    Story.append(Spacer(1, 10))
    
    # Section 1: Colab Commands
    Story.append(Paragraph("1. Explanation of Google Colab Commands", styles['Heading2_Step']))
    Story.append(Paragraph("The commands entered in the Colab Notebook are designed to download the project, install its dependencies, and make the local web server available on the public internet so you can view it from anywhere.", styles['Justify']))
    
    Story.append(Paragraph("<b>Command 1: Downloading the Project</b>", styles['Normal']))
    Story.append(Paragraph("!git clone https://github.com/ahmadnadeembutt/smartclaim.git", styles['CodeText']))
    Story.append(Paragraph("This command downloads (clones) the entire SmartClaim repository from GitHub into the Google Colab environment. It brings in all the python scripts, models, and data needed.", styles['Justify']))
    
    Story.append(Paragraph("<b>Command 2: Installing Dependencies</b>", styles['Normal']))
    Story.append(Paragraph("%cd smartclaim<br/>!pip install -r requirements.txt", styles['CodeText']))
    Story.append(Paragraph("The `%cd` command changes the current directory to the newly downloaded folder. The `pip install` command reads the `requirements.txt` file and installs all the necessary Python libraries required to run the AI model and the frontend (like Streamlit, Pandas, Scikit-learn, Sentence-Transformers).", styles['Justify']))
    
    Story.append(Paragraph("<b>Command 3: Running the App & Creating a Public Link</b>", styles['Normal']))
    Story.append(Paragraph("!streamlit run app.py &>/content/logs.txt &<br/>!npx localtunnel --port 8501", styles['CodeText']))
    Story.append(Paragraph("<b>`!streamlit run app.py`</b>: This starts the Streamlit web server. The `&` at the end tells Colab to run it in the background so the cell doesn't get stuck. The output is saved to `logs.txt`.", styles['BulletItem']))
    Story.append(Paragraph("<b>`!npx localtunnel`</b>: Streamlit runs locally inside Colab on port 8501. Because Colab is a closed cloud environment, you cannot access port 8501 directly. Localtunnel creates a temporary public URL that securely forwards traffic from the internet to port 8501 inside Colab.", styles['BulletItem']))
    
    Story.append(Spacer(1, 20))
    
    # Section 2: Frontend Code
    Story.append(Paragraph("2. Explanation of the Frontend Code (app.py)", styles['Heading2_Step']))
    Story.append(Paragraph("The `app.py` file uses Streamlit to create the user interface. It acts as the bridge between the user and the Machine Learning models.", styles['Justify']))
    
    Story.append(Paragraph("<b>A. Model Caching and Preloading</b>", styles['Normal']))
    Story.append(Paragraph("@st.cache_resource()<br/>def load_ml_pipeline(): ...", styles['CodeText']))
    Story.append(Paragraph("Loading deep learning models takes a lot of time. By using `@st.cache_resource()`, Streamlit ensures that the heavy AI models are loaded into the server's memory only once when the app starts. When a user clicks 'Predict', the model is already in memory, making predictions instantaneous.", styles['Justify']))
    
    Story.append(Paragraph("<b>B. Page Navigation System</b>", styles['Normal']))
    Story.append(Paragraph("Streamlit uses a sidebar for navigation (`st.sidebar.radio`). Based on the user's selection, different functions are called to render different pages (e.g., `page_predict()`, `page_methodology()`, `page_about()`). This keeps the app organized into a multi-page dashboard without needing complex HTML/CSS routing.", styles['Justify']))
    
    Story.append(Paragraph("<b>C. The Prediction Page Logic</b>", styles['Normal']))
    Story.append(Paragraph("The main page contains a Text Area where users enter their Arabic accident description. Below it are 'Quick Example' buttons. When an example button is clicked, it uses a <b>Callback Function</b> (`on_click=set_example_text`) to instantly update the text area's internal state variable without breaking Streamlit's rendering cycle.", styles['Justify']))
    
    Story.append(Paragraph("<b>D. Integration with the AI Model</b>", styles['Normal']))
    Story.append(Paragraph("from src.predict import predict_cost<br/>...<br/>result = predict_cost(user_text)", styles['CodeText']))
    Story.append(Paragraph("When the user clicks 'Analyze & Predict', the `app.py` file takes the text from the text box and passes it to the `predict_cost()` function imported from the backend (`src/predict.py`). It then takes the resulting predicted cost category (like 'Minor', 'Moderate', or 'Severe') and displays it beautifully using Streamlit metrics and success banners.", styles['Justify']))
    
    Story.append(Paragraph("<b>E. UI Aesthetics and Styling</b>", styles['Normal']))
    Story.append(Paragraph("The app uses custom CSS injected via `st.markdown(..., unsafe_allow_html=True)` to enforce a premium Dark Theme with custom fonts, glassmorphism effects on cards, and glowing hover states on buttons.", styles['Justify']))
    
    Story.append(Spacer(1, 30))
    Story.append(Paragraph("<i>Generated by SmartClaim AI Assistant for Presentation Purposes</i>", styles['Normal']))
    
    doc.build(Story)
    print(f"PDF successfully created at: {pdf_path}")

if __name__ == "__main__":
    create_explanation_pdf()
