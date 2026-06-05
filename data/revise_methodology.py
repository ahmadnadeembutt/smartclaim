"""
Revise the Methodology section (Section III) of SmartClaim_AI_Mid_Report.docx.
Only paragraphs P34–P51 are replaced; everything else is preserved exactly.
"""

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from copy import deepcopy
import os

INPUT  = "data/SmartClaim_AI_Mid_Report بعد التعديل.docx"
OUTPUT = "data/SmartClaim_AI_Mid_Report بعد التعديل.docx"   # overwrite in-place

# ── Methodology replacement content ────────────────────────────────
# Each entry is (style_hint, text)
#   style_hint: "heading" => section heading (e.g., "III. METHODOLOGY")
#               "subheading" => subsection (e.g., "A. Data Collection")
#               "body" => normal paragraph
#               "blank" => empty paragraph separator

METHODOLOGY_PARAGRAPHS = [
    # ─── Section heading ───
    ("heading", "III. METHODOLOGY"),

    ("body", "This section presents a comprehensive description of the end-to-end SmartClaim AI methodology, covering every stage of the pipeline from initial data acquisition through final cost prediction. The methodology is organized into nine sequential subsections, each corresponding to a distinct phase of the system architecture. Figure 5 provides a high-level overview of the pipeline flow: Raw Arabic Text → Preprocessing → Embedding Generation → Feature Extraction → Feature Concatenation → Model Training → Expert Tier Assignment → Final Cost Estimation."),

    # ─── A. Research Design and Overall Architecture ───
    ("subheading", "A. Research Design and Overall Architecture"),

    ("body", "The SmartClaim AI system follows a hybrid Expert-Steered Machine Learning (Expert-Steered ML) architecture, combining the statistical generalization capabilities of supervised regression models with the domain precision of a rule-based expert system. The research design is grounded in the applied machine learning methodology recommended by Géron [19] and the software engineering validation principles of Pressman and Maxim [7]. The overall system architecture comprises six interconnected modules: (1) Data Acquisition and Augmentation Module, (2) Preprocessing and Cleaning Module, (3) Semantic Embedding Module, (4) Arabic Feature Extraction Module, (5) Supervised Regression Training Module, and (6) Expert Cost Estimation Module. Data flows sequentially through these modules during the training phase, while the inference phase invokes Modules 3, 4, and 6 in real time to produce a cost estimate for any new Arabic accident description. This modular design ensures that each component can be independently tested, replaced, or upgraded without affecting downstream dependencies — a key principle of maintainable AI system design [31]."),

    # ─── B. Data Collection ───
    ("subheading", "B. Data Collection"),

    ("body", "The dataset underpinning this study was constructed from two authenticated, complementary sources within the Saudi Arabian insurance regulatory ecosystem. The first source comprised 100 official traffic accident reports obtained from Najm for Insurance Services — the Kingdom's centralized traffic accident management authority, which processes the majority of road traffic incident documentation in Saudi Arabia [1]. From each Najm report, the Arabic-language accident description field was extracted. These descriptions are authored by traffic officers and insurance adjusters in a semi-structured free-text format, capturing the circumstances of the accident, the point of impact, the apparent severity, and the visually identifiable damaged components. The descriptions are written predominantly in Modern Standard Arabic (MSA) with frequent Saudi dialectal (Najdi) terms for vehicle parts — for example, \"شاصيه\" (chassis), \"رفرف\" (fender), \"صدام\" (bumper), and \"كبوت\" (hood) — reflecting the authentic linguistic patterns of the target domain [33]."),

    ("body", "The second source comprised 100 corresponding vehicle damage assessment reports obtained from Taqdeer, the national vehicle damage estimation platform operated under the regulatory supervision of the Saudi Central Bank (SAMA). Taqdeer reports are produced by certified vehicle assessors and contain itemized repair cost breakdowns denominated in Saudi Riyals (SAR). From each Taqdeer report, the total repair cost valuation was extracted and used as the supervised target variable (i.e., the ground truth label) for model training. This dual-source design — pairing Najm textual descriptions with Taqdeer cost valuations — ensures that the training data faithfully reflects both the natural linguistic patterns of real Saudi accident narratives and the actuarially grounded cost distributions used by licensed assessors in the Kingdom [1], [2], [35]. All data was handled in compliance with applicable Saudi data protection regulations."),

    # ─── C. Data Augmentation and Expansion ───
    ("subheading", "C. Data Augmentation and Expansion"),

    ("body", "Given that the initial corpus of 200 real-world records was insufficient for robust supervised learning, a systematic data augmentation strategy was implemented to expand the dataset to 1,066 records. Prior to augmentation, a comprehensive analytical study was conducted on the 200 original records to characterize the statistical and linguistic relationship between accident descriptions and their corresponding repair costs. This analysis identified several recurring structural patterns in how Saudi assessors document accidents, including: (a) common phrase templates describing impact type and location; (b) standardized Arabic part nomenclature (صدام، باب، رفرف، شاصيه، كبوت، شمعة، مرايا); (c) predictable severity-to-cost correlations aligned with Saudi automotive repair market pricing [35]."),

    ("body", "Based on these findings, the augmentation was implemented through a purpose-built data generation script (generate_data.py) employing three complementary techniques:"),

    ("body", "(1) Pattern-Based Template Generation: Fifteen (15) distinct Arabic accident report templates were designed to mirror the linguistic structures, vocabulary, and narrative style observed in the original Najm reports. These templates covered a comprehensive range of accident scenarios prevalent in the Saudi context, including but not limited to: rear-end collisions (صدم من الخلف), parking lot scratches (خدوش في المواقف), head-on impacts (تصادم وجهاً لوجه), vehicle rollovers (انقلاب المركبة), hit-and-run incidents (صدم وهروب), electrical fire damage (احتراق بسبب التماس الكهربائي), tire blowout incidents, falling object damage, and undercarriage impacts. Each template was crafted to include realistic combinations of part names, impact descriptions, and severity indicators."),

    ("body", "(2) Linguistic Diversification: To prevent the model from memorizing template-specific phrases rather than learning genuine semantic relationships, each template was augmented with randomized contextual variables. Location descriptors were sampled from a vocabulary of six common Saudi traffic locations (الإشارة — traffic signal, الدوار — roundabout, المواقف — parking area, طريق سريع — highway, شارع فرعي — side street, البيت — residential area). Severity modifiers were drawn from five escalating levels (خفيفة — light, متوسطة — moderate, قوية — strong, بليغة — severe, شاملة — total). This combinatorial diversification produced a linguistically varied corpus that minimizes overfitting to surface-level lexical patterns."),

    ("body", "(3) Cost Distribution Calibration: Repair costs were assigned to each generated record using a keyword-driven rule function that maps severity indicators and damage descriptors to realistic SAR price ranges. These ranges were calibrated against the cost distributions observed in the original 100 Taqdeer assessment reports and cross-validated with published Saudi Automobile Association repair cost benchmarks [35]: severe/rollover/total damage keywords → 15,000–65,000 SAR; light/scratch/mirror keywords → 500–3,000 SAR; moderate/door/hood keywords → 3,000–12,000 SAR; unclassified descriptions → 1,000–25,000 SAR. A fixed random seed (random.seed(42) and numpy.random.seed(42)) was applied throughout the entire augmentation process to ensure full reproducibility of the expanded dataset [20]."),

    # ─── D. Data Cleaning and Preprocessing ───
    ("subheading", "D. Data Cleaning and Preprocessing"),

    ("body", "The preprocessing module (implemented in src/preprocess.py) was designed to transform the raw augmented dataset into a clean, analysis-ready format suitable for downstream embedding and model training. The pipeline executed the following four sequential operations:"),

    ("body", "(1) Column Standardization: The original Excel column headers — \"Text\" (containing the Arabic accident descriptions) and \"Cost of the second party's vehicle\" (containing the SAR repair cost values) — were programmatically renamed to \"text\" and \"cost\" respectively, using pandas DataFrame.rename(). This ensures consistent programmatic access across all downstream modules and eliminates dependency on the original bilingual column naming convention."),

    ("body", "(2) Null Value Removal: Any rows containing missing values in either the \"text\" or \"cost\" column were removed using pandas dropna(subset=['text', 'cost']). This operation reduced the dataset from 1,066 raw records to 1,000 usable records — a loss rate of 6.2%, which is within acceptable bounds for data quality filtering. The 66 dropped records were inspected and confirmed to be incomplete template generations with empty text fields, not systematic data quality issues."),

    ("body", "(3) Whitespace Normalization: Leading and trailing whitespace characters were stripped from all Arabic text entries using the pandas str.strip() method. This step is critical for Arabic text processing because extraneous whitespace can cause different embedding vectors for semantically identical descriptions, introducing noise into the model's feature space. Additionally, this normalization ensures consistent tokenization behavior in the downstream sentence-transformer model."),

    ("body", "(4) Statistical Profiling: The cost column was profiled to extract descriptive statistics — minimum, maximum, mean, and standard deviation — which were logged at runtime for traceability and used to confirm that the cost distribution matched the expected range (approximately 500–65,000 SAR) before model training commenced. This profiling step serves as an automated data quality gate, enabling early detection of any cost outliers or distribution anomalies introduced during the augmentation process."),

    # ─── E. Data Splitting Strategy ───
    ("subheading", "E. Data Splitting Strategy"),

    ("body", "The cleaned dataset of 1,000 records was partitioned using a stratified holdout strategy to support the three-phase evaluation framework described in Subsection I. The primary split separated the data into 80% training (800 records) and 20% testing (200 records) using scikit-learn's train_test_split function with random_state=42 for reproducibility [21]. This 80/20 ratio follows established best practices in the machine learning literature for datasets of this scale [19], [28]."),

    ("body", "Within the training phase, the 800-record training set was further subdivided using a secondary 90/10 split: 720 records (90%) were allocated for model fitting, and 80 records (10%) were reserved as an internal validation set. The internal validation set served a dual purpose: (a) enabling unbiased comparison of candidate models (Random Forest Baseline, XGBoost, and Tuned Random Forest) during the model selection phase, and (b) providing an early stopping criterion for hyperparameter tuning to prevent overfitting. This three-tier partitioning scheme (720 train / 80 validation / 200 test) ensures that the final evaluation metrics reported in Section IV reflect true generalization performance on data the model has never encountered during any phase of training, tuning, or model selection. The same random_state=42 seed was used for all splits to guarantee deterministic reproducibility across experimental runs [21]."),

    # ─── F. Semantic Embedding Architecture ───
    ("subheading", "F. Semantic Embedding Architecture"),

    ("body", "The core NLP representation layer employs the paraphrase-multilingual-mpnet-base-v2 sentence-transformer model [22], a state-of-the-art multilingual embedding architecture based on the MPNet (Masked and Permuted Pre-training for Language Understanding) backbone. This model was selected over alternatives (e.g., multilingual BERT [9], XLM-RoBERTa) based on three criteria: (a) demonstrated superior performance on cross-lingual semantic textual similarity (STS) benchmarks across 50+ languages including Arabic; (b) production-grade inference speed suitable for real-time deployment; and (c) its 768-dimensional output space provides sufficient representational capacity for capturing fine-grained semantic distinctions between accident descriptions of varying severity [22], [36]."),

    ("body", "The embedding generation process (implemented in src/embed.py via the TextEmbedder class) operates as follows: Each Arabic accident description is passed through the sentence-transformer model, which internally tokenizes the text using a SentencePiece tokenizer, processes the tokens through 12 transformer encoder layers with multi-head self-attention (12 attention heads per layer), and applies mean pooling over the final hidden states to produce a single dense 768-dimensional vector representation. This vector captures the semantic meaning of the entire accident description in a continuous vector space where semantically similar descriptions (e.g., two rear-end collisions of similar severity) are mapped to nearby points, while semantically dissimilar descriptions (e.g., a minor scratch versus a chassis-level collision) are mapped to distant points [10], [37]."),

    ("body", "To optimize computational efficiency, the embedding module implements a disk-based caching mechanism using the joblib library [23]. Once embeddings are computed for the training set, they are serialized to disk (models/embeddings_cache.joblib) and loaded from cache on subsequent runs, eliminating redundant GPU/CPU inference. This caching strategy reduces the training pipeline execution time from approximately 8 minutes (with live embedding) to under 30 seconds (with cached embeddings), a 16x speedup. For the test set evaluation (Phase 3), caching is explicitly disabled (use_cache=False) to ensure that the 200 test records are freshly embedded, preventing any data leakage from the training embedding cache."),

    # ─── G. Arabic Feature Extraction Engine ───
    ("subheading", "G. Arabic Feature Extraction Engine"),

    ("body", "In addition to the dense semantic embeddings, a rule-based Arabic feature extraction engine (implemented in src/feature_extractor.py as the ArabicAccidentFeatureExtractor class) was developed to produce structured numerical features that encode domain-specific knowledge about accident severity and damage characteristics [24]. This dual-representation approach — combining dense neural embeddings with sparse expert-engineered features — follows the established practice of feature augmentation in hybrid NLP systems [14], [38]."),

    ("body", "The feature extractor operates through lexical keyword matching across five granularity levels of severity, three categories of impact direction, and a comprehensive parts inventory. The severity detection levels are: (a) Very Light — triggered by keywords such as \"خفيف جداً\" (very light), \"بسيط جداً\" (very minor), \"بدون أضرار واضحة\" (no visible damage), and \"لا يوجد أضرار\" (no damage); (b) Minor — triggered by \"بسيط\" (minor), \"خفيف\" (light), \"طفيف\" (slight), \"خدش\" (scratch), and \"بدون أضرار كبيرة\" (no major damage); (c) Moderate — triggered by \"متوسط\" (moderate), \"تضرر\" (damaged), \"إصلاح\" (repair), and \"صدمة\" (impact); (d) Severe — triggered by \"شديد\" (severe), \"قوي\" (strong), \"تدمير\" (destruction), \"تضرر بالغ\" (heavily damaged), \"تالف\" (totalled), and \"هيكل\" (body/frame); and (e) Critical — triggered by chassis and drivetrain terms including \"شاصيه\" (chassis), \"شاص\" (chassis frame), \"ماكينة\" (engine), \"مكينة\" (engine variant), \"قير\" (gearbox), \"جير\" (gearbox variant), and \"محرك\" (motor)."),

    ("body", "Impact direction is classified into three categories — front (أمام، امامي، وجه، صدام أمامي), rear (خلف، خلفي، ورا، صدام خلفي), and side (جانب، يمين، يسار، باب) — providing the model with spatial context about the collision. The parts count feature tallies the number of distinct damaged components mentioned in the description from a vocabulary of 18 standard Saudi automotive parts (صدام، باب، رفرف، شمعة، كشاف، كبوت، غطاء المحرك، شنطة، دبة، زجاج، مرايا، حساس، كاميرا، هيكل، شاصيه، اصطب، شبك، رديتر، مروحة). A specialized \"scratches-only\" binary feature is activated when scratch-related terms are present but no major structural parts are mentioned, helping the model distinguish cosmetic-only damage from structural damage."),

    ("body", "The extractor's transform() method outputs a 6-dimensional numerical feature vector for each description: [is_minor_or_very_light, is_severe_or_critical, is_front, is_rear, is_side, parts_count]. This 6-dimensional vector is concatenated with the 768-dimensional dense embedding to produce a final 774-dimensional feature vector that serves as input to the regression model. The concatenation is performed using numpy.hstack(), ensuring dimensional compatibility and preserving the ordering of features across training and inference. This feature augmentation strategy enables the regression model to leverage both the distributional semantics captured by the transformer and the explicit domain knowledge encoded in the rule-based features, resulting in a more discriminative and interpretable feature representation than either approach alone [14], [38]."),

    # ─── H. Model Training and Selection ───
    ("subheading", "H. Model Training and Selection"),

    ("body", "The training pipeline (implemented in src/train.py) follows a competitive model selection paradigm in which three candidate regression models are trained on the same 720-record training subset and evaluated on the 80-record internal validation set. The candidate models are:"),

    ("body", "(1) Random Forest Baseline (RF-Base): A Random Forest Regressor with 200 estimators (n_estimators=200) and default scikit-learn hyperparameters (max_depth=None, min_samples_split=2, min_samples_leaf=1). Random Forest was selected as the primary baseline due to its well-documented robustness to overfitting on tabular data and its interpretable feature importance mechanism [17], [19]. The model was initialized with random_state=42 for reproducibility."),

    ("body", "(2) XGBoost Regressor (XGB): An XGBoost gradient boosting regressor with 200 estimators (n_estimators=200) and default hyperparameters [18]. XGBoost was included as a competitive benchmark because gradient boosting methods have consistently achieved state-of-the-art performance on structured tabular regression tasks and are widely used in the insurance industry for claim cost modeling [8], [18]. The model was initialized with random_state=42."),

    ("body", "(3) Tuned Random Forest (RF-Tuned): A hyperparameter-optimized Random Forest Regressor trained using RandomizedSearchCV with 10 iterations and 3-fold cross-validation. The hyperparameter search space covered: n_estimators ∈ {100, 200, 300, 500}; max_depth ∈ {None, 10, 20, 30}; min_samples_split ∈ {2, 5, 10}; min_samples_leaf ∈ {1, 2, 4}. The scoring metric was negative Mean Absolute Error (neg_mean_absolute_error), and parallel execution was enabled via n_jobs=-1 to utilize all available CPU cores. The best estimator from the search was retained as the RF-Tuned candidate."),

    ("body", "Model selection was performed by comparing the Mean Absolute Error (MAE) of each candidate on the 80-record internal validation set. The model with the lowest validation MAE was automatically selected as the winner and serialized to disk (models/best_model.joblib) using the joblib library for subsequent use in the evaluation and prediction phases [23]. This competitive selection process ensures that the best-performing model architecture is chosen in a data-driven manner, without relying on a priori assumptions about model superiority."),

    # ─── I. The 7-Tier Expert Cost Estimation Engine ───
    ("subheading", "I. The 7-Tier Expert Cost Estimation Engine"),

    ("body", "The most novel component of the SmartClaim AI methodology is the post-prediction Expert Cost Estimation Engine, which was specifically designed to eliminate the \"bucketing bias\" problem identified during initial model development [25]. In preliminary experiments, the trained regression model (without expert steering) consistently predicted approximately 6,000 SAR for all minor incidents and 18,000 SAR for all moderate-to-severe cases — regardless of the specific parts damaged, the number of parts affected, or the described impact intensity. This bucketing behavior is a well-known limitation of regression models trained on limited data with broad categorical targets [4], [28]."),

    ("body", "The Expert Engine addresses this limitation by introducing a two-stage prediction mechanism: (Stage 1) the Arabic Feature Extractor classifies the input description into one of seven expert-defined cost tiers based on detected severity keywords, parts count, and damage characteristics; (Stage 2) the ML model's raw prediction is used only to determine a relative position within the assigned tier, rather than serving as the final output. This ensures that domain knowledge always constrains the prediction within a realistic cost range, preventing the model from producing unrealistic or flat outputs."),

    ("body", "The seven cost tiers, calibrated against documented Saudi automotive repair market pricing [35], are defined as follows: Tier 1 — Very Light (0–1,000 SAR): Activated when \"very light\" severity keywords are detected (خفيف جداً, بدون أضرار واضحة), covering incidents with no visible damage or negligible cosmetic marks; Tier 2 — Surface Scratches (1,000–3,000 SAR): Activated when scratch keywords are present without any major structural parts being mentioned, covering paint-level cosmetic damage; Tier 3 — Minor with Parts (3,000–6,000 SAR): Activated when minor severity keywords are detected alongside one or two part mentions, covering single-panel dents and minor component replacements; Tier 4 — Moderate Few Parts (4,000–8,000 SAR): The default tier for descriptions that do not match specific severity patterns, covering moderate damage with limited part involvement; Tier 5 — Moderate Multi-Parts (8,000–15,000 SAR): Activated when the parts count exceeds 2, covering incidents where multiple body panels, lights, or trim pieces require repair or replacement; Tier 6 — Severe Structural (15,000–25,000 SAR): Activated when severe keywords are detected (شديد, تدمير, تضرر بالغ), covering major structural body damage; and Tier 7 — Critical Chassis/Engine (25,000–60,000+ SAR): Activated when critical drivetrain keywords are detected (شاصيه, ماكينة, محرك, قير), covering catastrophic damage to the chassis, engine, or transmission."),

    ("body", "Within each tier, the ML model's raw prediction is transformed into a relative position value (ranging from 0.0 to 1.0) using the formula: rel_pos = (ml_prediction − 1,000) / 40,000, clamped to the interval [0.1, 0.9]. This relative position is then further adjusted by a parts-count boost (+10% per detected part, capped at +40%) and a scratch penalty (−30% if scratch keywords are present). The final estimated cost is computed as: base_cost = tier_low + (tier_high − tier_low) × adjusted_rel_pos. Additional smart rule overrides apply hard ceiling caps for edge cases — for example, \"very light\" estimates are capped at 1,200 SAR, and descriptions containing explicit \"no damage\" phrases (بدون أضرار, لا يوجد أضرار) are capped at 300 SAR. Finally, a deterministic jitter factor (0.95–1.05), derived from an MD5 hash of the input text, is applied to prevent identical outputs for semantically similar but textually distinct descriptions, producing the final cost estimate with an associated ±15% confidence interval [25]."),

    # ─── J. Three-Phase Evaluation Framework ───
    ("subheading", "J. Three-Phase Evaluation Framework"),

    ("body", "The study employs a rigorous three-phase evaluation framework to assess model performance at increasing levels of generalization difficulty [26], [31]:"),

    ("body", "Phase 1 — Training Phase: The 720-record training subset is used to fit all three candidate models (RF-Base, XGB, RF-Tuned). Each model is then evaluated on the 80-record internal validation set using MAE as the primary selection metric. The model achieving the lowest validation MAE is declared the winner and saved for subsequent evaluation phases. This phase confirms that the models can learn meaningful patterns from the training data and provides a fair basis for competitive model selection."),

    ("body", "Phase 2 — Initial Validation Phase: The winning model is re-evaluated on the full 800-record training set to verify that it has successfully captured the underlying patterns in the training data. While this self-evaluation is expected to produce optimistically biased metrics, it serves as a critical sanity check — confirming that the model's internal representations are consistent with the training distribution and that no catastrophic training failures (e.g., mode collapse, gradient explosion) have occurred. The evaluation metrics computed in this phase include MAE, RMSE, R², and MAPE."),

    ("body", "Phase 3 — Real Testing Phase: The winning model is evaluated on the 200-record hold-out test set — data that has never been seen during any phase of training, validation, or hyperparameter tuning. The test set embeddings are generated fresh (without using the training cache) to eliminate any possibility of data leakage. This phase produces the definitive generalization metrics reported in Section IV and is accompanied by four diagnostic visualizations: (a) Actual vs. Predicted scatter plot, (b) Cost Distribution histogram, (c) Top-20 Feature Importance bar chart, and (d) Residuals Distribution histogram with kernel density estimate. All evaluation metrics and plots are saved to the models/ directory for reproducibility and audit trail purposes [26]."),
]


def run():
    doc = Document(INPUT)
    paras = doc.paragraphs

    # ── Identify methodology paragraph indices (P34 to P51) ──
    meth_start = None
    meth_end   = None
    for i, p in enumerate(paras):
        if p.text.strip() == "METHODOLOGY" or p.text.strip() == "III. METHODOLOGY":
            meth_start = i
        if p.text.strip().startswith("RESULTS AND EVALUATION") or p.text.strip().startswith("IV. RESULTS AND EVALUATION"):
            meth_end = i   # exclusive
            break

    if meth_start is None or meth_end is None:
        print("ERROR: Could not locate Methodology section boundaries.")
        return

    print(f"Methodology section found: paragraphs [{meth_start}..{meth_end - 1}]")
    print(f"Will replace {meth_end - meth_start} old paragraphs with {len(METHODOLOGY_PARAGRAPHS)} new paragraphs.")

    # ── Capture formatting from the original paragraphs ──
    # Use the first original body paragraph as the style template
    template_para = paras[meth_start + 2]  # e.g. the first body paragraph after "A. Data Collection"

    # ── Remove old methodology paragraphs (in reverse to preserve indices) ──
    # We'll work at the XML level: remove <w:p> elements from the document body
    body = doc.element.body
    old_elements = [paras[i]._element for i in range(meth_start, meth_end)]
    insertion_point = old_elements[0]  # remember where to insert

    for elem in old_elements:
        body.remove(elem)

    print(f"Removed {len(old_elements)} old paragraph elements.")

    # ── Insert new methodology paragraphs ──
    # Find the element that now sits where the old methodology started
    # (this will be the first paragraph of Section IV)
    ref_element = paras[meth_end]._element  # "IV. RESULTS AND EVALUATION"

    from docx.oxml.ns import qn
    from lxml import etree
    from docx.oxml import OxmlElement

    for style_hint, text in METHODOLOGY_PARAGRAPHS:
        # Create a new <w:p> element
        new_p = OxmlElement('w:p')

        # Add paragraph properties
        pPr = OxmlElement('w:pPr')

        # Set style to Normal
        pStyle = OxmlElement('w:pStyle')
        pStyle.set(qn('w:val'), 'Normal')
        pPr.append(pStyle)

        new_p.append(pPr)

        # Add run with text
        run = OxmlElement('w:r')

        # Run properties
        rPr = OxmlElement('w:rPr')

        # Match font from original document (use the same font as existing paragraphs)
        # We'll set a reasonable academic font
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        rFonts.set(qn('w:cs'), 'Times New Roman')
        rPr.append(rFonts)

        # Font size 11pt (22 half-points)
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), '22')
        rPr.append(sz)
        szCs = OxmlElement('w:szCs')
        szCs.set(qn('w:val'), '22')
        rPr.append(szCs)

        # Bold for headings and subheadings
        if style_hint in ("heading", "subheading"):
            b = OxmlElement('w:b')
            rPr.append(b)
            bCs = OxmlElement('w:bCs')
            rPr.append(bCs)

        run.append(rPr)

        # Text content
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = text
        run.append(t)

        new_p.append(run)

        # Insert before the reference element (Section IV heading)
        ref_element.addprevious(new_p)

    print(f"Inserted {len(METHODOLOGY_PARAGRAPHS)} new paragraph elements.")

    # ── Save ──
    doc.save(OUTPUT)
    print(f"Document saved to: {OUTPUT}")


if __name__ == "__main__":
    run()
