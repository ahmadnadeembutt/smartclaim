import pandas as pd
import numpy as np
import random
import os

# Set random seed for reproducibility
random.seed(42)
np.random.seed(42)

def generate_synthetic_data(output_path="data/smartclaim_expanded_dataset.xlsx", total_records=1066):
    """
    Generates a synthetic dataset of Arabic car accident reports and their repair costs.
    """
    
    # Templates for Arabic accident reports
    templates = [
        "صدمت سيارة من الخلف عند الإشارة المرورية مما أدى لتضرر الصدام الخلفي.",
        "انحراف مفاجئ للمركبة أدى إلى الاصطدام بالرصيف وتضرر المساعدات والجنت.",
        "خدوش عميقة في الباب الأيمن والرفرف الأمامي بسبب احتكاك في مواقف السيارات.",
        "حادث تصادم وجهاً لوجه أدى إلى تلف الواجهة الأمامية بالكامل والرديتر.",
        "سقوط جسم صلب على الزجاج الأمامي مما أدى إلى تهشمه بالكامل.",
        "انقلاب المركبة بسبب السرعة العالية مما أدى إلى تلف الهيكل بشكل بليغ.",
        "صدمة خفيفة في المرآة الجانبية لجهة السائق.",
        "تعرضت السيارة لحادث صدم وهروب أثناء التوقف.",
        "تضرر الجزء السفلي من السيارة نتيجة المرور فوق عائق صلب.",
        "احتراق جزئي في ضفيرة المحرك نتيجة التماس كهربائي بعد حادث.",
        "كسر في المصباح الأمامي الأيسر وتطعج في الكبوت.",
        "صدمة قوية في الباب الخلفي الأيسر أدت لعدم فتحه.",
        "تضرر الشكمان والصدام الخلفي نتيجة صدمة من سيارة دفع رباعي.",
        "انفجار الإطار الأمامي أدى لفقدان السيطرة والارتطام بالحاجز الخرساني.",
        "حجارة متطايرة أدت لخدوش في الطلاء الخارجي للكبوت والزجاج.",
    ]
    
    # Vocabulary to diversify reports
    locations = ["الإشارة", "الدوار", "المواقف", "طريق سريع", "شارع فرعي", "البيت"]
    severities = ["خفيفة", "متوسطة", "قوية", "بليغة", "شاملة"]
    
    data = []
    
    for i in range(total_records):
        # Pick a base template or construct one
        if random.random() > 0.3:
            report = random.choice(templates)
        else:
            report = f"صدمة {random.choice(severities)} في {random.choice(locations)} أدت لتضرر أجزاء من السيارة."
        
        # Calculate a realistic cost based on keywords in the report
        # Basic cost logic
        if "بليغ" in report or "انقلاب" in report or "كامل" in report:
            cost = random.randint(15000, 65000)
        elif "خفيف" in report or "خدوش" in report or "مرآة" in report:
            cost = random.randint(500, 3000)
        elif "متوسط" in report or "باب" in report or "كبوت" in report:
            cost = random.randint(3000, 12000)
        else:
            cost = random.randint(1000, 25000)
            
        data.append({"Text": report, "Cost of the second party's vehicle": float(cost)})

    df = pd.DataFrame(data)
    
    # Ensure directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # Save to Excel
    df.to_excel(output_path, index=False)
    print(f"Dataset generated successfully at {output_path} with {len(df)} records.")

if __name__ == "__main__":
    generate_synthetic_data()
