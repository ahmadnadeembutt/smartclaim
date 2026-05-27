import sys
import os
from src.predict import predict_cost

def run_verification():
    test_cases = [
        "حادث بسيط صدم من الخلف عند الإشارة", # Expected: ~4000-6000 (Minor)
        "صدمة في الباب الجانبي وخدوش في الرفرف", # Expected: ~3000-5000 (Manual improvement: ~3000)
        "تدمير قوي في الواجهة الأمامية وتضرر الشاصيه والماكينة", # Expected: 25,000 - 60,000 (Critical)
        "تضرر الباب الأيمن والرفرف والصدام الخلفي يحتاج إصلاح", # Expected: 8,000 - 15,000 (Multiple Parts)
        "صدمة خفيفة جداً بدون أضرار واضحة" # Expected: 0 - 1,500 (Very Light)
    ]

    print("\n" + "="*60)
    print("SMARTCLAIM AI - BATCH VERIFICATION TEST")
    print("="*60)

    for i, text in enumerate(test_cases, 1):
        print(f"\nTest Case {i}: {text}")
        result = predict_cost(text)
        
        if "error" in result:
            print(f"Error: {result['error']}")
        else:
            print(f"Severity:       {result['severity']}")
            print(f"Estimated Cost: {result['predicted_cost']}")
            print(f"Confidence:     {result['confidence_range_low']} - {result['confidence_range_high']}")
        print("-" * 40)

    print("\n" + "="*60)
    print("VERIFICATION COMPLETE")
    print("="*60)

if __name__ == "__main__":
    run_verification()
