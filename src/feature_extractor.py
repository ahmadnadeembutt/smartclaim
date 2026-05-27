import re
import numpy as np

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
            # Maintain backward compatibility with the ML model's expected 6-feature input
            # [is_minor, is_severe, is_front, is_rear, is_side, parts_count]
            # Note: is_minor in the model now includes is_very_light effectively
            features_list.append([
                float(f["is_minor"] or f["is_very_light"]),
                float(f["is_severe"] or f["is_critical"]),
                float(f["is_front"]),
                float(f["is_rear"]),
                float(f["is_side"]),
                float(f["parts_count"])
            ])
        return np.array(features_list, dtype=np.float32)

if __name__ == "__main__":
    extractor = ArabicAccidentFeatureExtractor()
    test_texts = [
        "حادث بسيط صدم من الخلف تضرر الصدام فقط",
        "تدمير قوي في الواجهة الأمامية الشاصيه والصدام والكبوت متضرر",
        "صدمة في الباب الجانبي الأيمن متوسطة"
    ]
    features = extractor.transform(test_texts)
    print("Features matrix shape:", features.shape)
    for t, f in zip(test_texts, features):
        print(f"Text: {t}\nFeatures [minor, severe, front, rear, side, parts]: {f}\n")
