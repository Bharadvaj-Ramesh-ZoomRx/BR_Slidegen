"""Create dummy Excel data for workflow testing.

Generates two wave files with RYB/TAG/Analysis sheets matching
the column layout expected by the pipeline extractors.
"""

import pandas as pd
import numpy as np
import os

BASE = os.path.dirname(__file__)


def create_dummy_excel(out_path: str, seed: int = 42):
    """Create a source_data.xlsx with RYB, TAG, and Additonal Analysis sheets."""
    rng = np.random.default_rng(seed)

    # ── RYB Sheet ──
    # col 0=code, col 1=desc, col 7=Q3 Total, col 17=Q4 Total
    ryb_rows = []
    # Pad with 18 cols (0..17)
    def _row(code, desc, q3, q4):
        r = [None] * 18
        r[0] = code
        r[1] = desc
        r[7] = q3
        r[17] = q4
        return r

    # Q2_10Z — Message Recall
    ryb_rows.append(_row("Q2_10Z", "Message Recall", None, None))
    messages = [
        "Efficacy data shows superior outcomes",
        "Safety profile is well-tolerated",
        "NCCN Category 1 recommendation",
        "Convenient dosing schedule",
        "First-line treatment option",
    ]
    for msg in messages:
        q3 = round(rng.uniform(0.3, 0.8), 4)
        q4 = round(q3 + rng.uniform(-0.1, 0.1), 4)
        ryb_rows.append(_row(None, msg, q3, q4))

    # C1_81Z — Compelling reason
    ryb_rows.append(_row("C1_81Z", "Compelling reason to prescribe", None, None))
    ryb_rows.append(_row(None, "Top Box", round(rng.uniform(0.4, 0.7), 4), round(rng.uniform(0.4, 0.7), 4)))

    # C1_82Z — Changed opinion
    ryb_rows.append(_row("C1_82Z", "Changed opinion", None, None))
    ryb_rows.append(_row(None, "Yes", round(rng.uniform(0.2, 0.5), 4), round(rng.uniform(0.2, 0.5), 4)))

    # C1_83D1Z — Branded closing
    ryb_rows.append(_row("C1_83D1Z", "Asked to prescribe (Branded closing)", None, None))
    ryb_rows.append(_row(None, "Top Box", round(rng.uniform(0.3, 0.6), 4), round(rng.uniform(0.3, 0.6), 4)))

    # Q1_84bZ — Educational ask
    ryb_rows.append(_row("Q1_84bZ", "Educational ask", None, None))
    ryb_rows.append(_row(None, "Yes", round(rng.uniform(0.3, 0.5), 4), round(rng.uniform(0.3, 0.5), 4)))

    # C1_84FZ — Follow-up reps
    ryb_rows.append(_row("C1_84FZ", "Follow-up with J&J reps", None, None))
    followups = ["Medical Science Liaison (MSL)", "Key Account Manager (KAM)",
                 "Field Reimbursement Manager (FRM)", "No, I did not set follow-up"]
    for fu in followups:
        ryb_rows.append(_row(None, fu, round(rng.uniform(0.1, 0.5), 4), round(rng.uniform(0.1, 0.5), 4)))

    # C1_85DZ — Rx Intent
    ryb_rows.append(_row("C1_85DZ", "Prescription Intent", None, None))
    patient_types = ["1L EGFR+ mNSCLC (no CNS)", "1L EGFR+ mNSCLC (with CNS)",
                     "2L+ EGFR+ mNSCLC", "Exon 20 insertion"]
    for pt in patient_types:
        ryb_rows.append(_row(None, pt, round(rng.uniform(0.3, 0.7), 4), round(rng.uniform(0.3, 0.7), 4)))

    # Q1_87Z — Quality metrics (also used for HII)
    ryb_rows.append(_row("Q1_87Z", "Quality Metrics", None, None))
    # Need cols 19 and 20 for HII
    quality_metrics = [
        "Overall quality of sales call",
        "How knowledgeable about competitor products",
        "How knowledgeable about EGFR exon 20",
        "Organized presentation of information",
        "Compelling reason to prescribe RYB",
    ]
    for qm in quality_metrics:
        r = [None] * 21  # extend to col 20
        r[0] = None
        r[1] = qm
        r[7] = round(rng.uniform(0.5, 0.9), 4)
        r[17] = round(rng.uniform(0.5, 0.9), 4)
        r[19] = round(rng.uniform(0.4, 0.8), 4)  # Others Q4
        r[20] = round(rng.uniform(0.6, 0.95), 4)  # HI Q4
        ryb_rows.append(r)

    ryb_df = pd.DataFrame(ryb_rows)

    # ── TAG Sheet ──
    # col 0=code, col 1=desc, col 7=Q3 Total, col 13=Q4 Total
    tag_rows = []
    def _tag_row(code, desc, q3, q4):
        r = [None] * 14
        r[0] = code
        r[1] = desc
        r[7] = q3
        r[13] = q4
        return r

    # Q2_10Z — TAG Message Recall
    tag_rows.append(_tag_row("Q2_10Z", "Message Recall", None, None))
    tag_msgs = [
        "mPFS 29.4 months with chemo",
        "89% of patients stay on treatment",
        "Grade 1/2 adverse reactions most common",
        "NCCN Category 1 recommendation (mono+chemo)",
        "Longest mPFS and mOS data",
        "Once-daily oral dosing convenience",
    ]
    for msg in tag_msgs:
        q3 = round(rng.uniform(0.2, 0.7), 4)
        q4 = round(q3 + rng.uniform(-0.15, 0.1), 4)
        tag_rows.append(_tag_row(None, msg, q3, q4))

    # TAG CTA codes
    for code in ["C1_81_TAG_mNSCLC_Z", "C1_82_TAG_mNSCLC_Z",
                 "C1_83D_TAG_mNSCLC_Z", "C1_84b_TAG_mNSCLC_Z"]:
        tag_rows.append(_tag_row(code, code, None, None))
        tag_rows.append(_tag_row(None, "Top Box", round(rng.uniform(0.3, 0.6), 4), round(rng.uniform(0.3, 0.6), 4)))

    # TAG Rx Intent (C1_85D_TAG2_mNSCLCZ)
    tag_rows.append(_tag_row("C1_85D_TAG2_mNSCLCZ", "Prescription Intent TAG", None, None))
    for pt in ["1L EGFR+ mNSCLC + chemo", "2L+ EGFR+ mNSCLC + chemo"]:
        tag_rows.append(_tag_row(None, pt, round(rng.uniform(0.3, 0.6), 4), round(rng.uniform(0.3, 0.6), 4)))

    # TAG Recall order (rows 250-303 ish) — need ordinal data
    # Pad to row 250
    while len(tag_rows) < 250:
        tag_rows.append(_tag_row(None, None, None, None))

    # Nested ordinal data starting at row 250
    ordinals = ["1st", "2nd", "3rd", "4th"]
    recall_msgs = ["mPFS data", "NCCN recommendation", "Safety profile", "Daily dosing"]
    for msg in recall_msgs:
        for ordinal in ordinals:
            r = [None] * 14
            r[0] = f"TAG_{msg[:4]}"
            r[1] = msg
            r[2] = ordinal
            r[7] = round(rng.uniform(0.05, 0.3), 4)
            r[13] = round(rng.uniform(0.05, 0.3), 4)
            tag_rows.append(r)

    tag_df = pd.DataFrame(tag_rows)

    # ── Additonal Analysis Sheet ──
    # col 1=metric, col 2=RYB_Q3, col 3=RYB_Q4, col 4=TAG_Q3, col 5=TAG_Q4
    # rows 4-25: Rep performance
    # rows 30-39: Message MR/ME/Believable (cols 3-8)
    # rows 136-151: Mariposa rep performance
    # rows 197-205: Share of time
    aa_rows = []

    # Pad to have enough rows
    for _ in range(210):
        aa_rows.append([None] * 10)

    # Rep performance (rows 4-24)
    rep_metrics = [
        "Overall quality of sales call",
        "How knowledgeable about competitor products",
        "How knowledgeable about EGFR exon 20",
        "How knowledgeable about product (RYB/TAG)",
        "How knowledgeable about EGFR mNSCLC landscape",
        "Organized presentation of information",
        "Compelling reason to prescribe",
        "Compelling reason to use broader molecular testing",
        "Credible support for product claims",
        "Prepared for conversation with me",
        "Engaging when discussing product",
        "Meaningful questions about my practice",
        "Valuable use of my time",
        "Address my questions about prescribing",
        "Delivering a clear message about product",
    ]
    for i, metric in enumerate(rep_metrics):
        row_idx = 4 + i
        aa_rows[row_idx][1] = metric
        aa_rows[row_idx][2] = round(rng.uniform(0.5, 0.9), 4)  # RYB Q3
        aa_rows[row_idx][3] = round(rng.uniform(0.5, 0.9), 4)  # RYB Q4
        aa_rows[row_idx][4] = round(rng.uniform(0.4, 0.8), 4)  # TAG Q3
        aa_rows[row_idx][5] = round(rng.uniform(0.4, 0.8), 4)  # TAG Q4

    # Message MR/ME/Believable (rows 30-39)
    msg_labels = [
        "Superior efficacy outcomes",
        "Well-tolerated safety profile",
        "NCCN Category 1 recommendation",
        "Convenient dosing schedule",
        "First-line treatment option",
    ]
    for i, msg in enumerate(msg_labels):
        row_idx = 30 + i
        aa_rows[row_idx][1] = msg
        aa_rows[row_idx][2] = msg[:20]  # short label
        aa_rows[row_idx][3] = round(rng.uniform(0.3, 0.8), 4)  # MR Q3
        aa_rows[row_idx][4] = round(rng.uniform(0.3, 0.8), 4)  # MR Q4
        aa_rows[row_idx][5] = round(rng.uniform(0.4, 0.8), 4)  # Believ Q3
        aa_rows[row_idx][6] = round(rng.uniform(0.4, 0.8), 4)  # Believ Q4
        aa_rows[row_idx][7] = round(rng.uniform(0.3, 0.7), 4)  # ME Q3
        aa_rows[row_idx][8] = round(rng.uniform(0.3, 0.7), 4)  # ME Q4

    # Mariposa rep performance (rows 136-150)
    mariposa_metrics = [
        "Overall quality of sales call",
        "Compelling reason to prescribe",
        "Knowledge: product",
        "Knowledge: competitors",
        "Organized presentation",
        "Clear product message",
        "Engaging discussion",
        "Valuable use of time",
    ]
    for i, metric in enumerate(mariposa_metrics):
        row_idx = 136 + i
        aa_rows[row_idx][2] = metric
        aa_rows[row_idx][3] = round(rng.uniform(0.5, 0.9), 4)  # overall
        aa_rows[row_idx][4] = round(rng.uniform(0.55, 0.95), 4)  # segment_a (Mariposa 1st)
        aa_rows[row_idx][5] = round(rng.uniform(0.4, 0.85), 4)  # segment_b (Not 1st)

    # Share of time (rows 197-204) — straight percentages
    time_topics = [
        "Efficacy data", "Safety profile", "Dosing information",
        "Guidelines/NCCN", "Patient support", "Access/reimbursement",
        "Competitive data", "Clinical trials",
    ]
    for i, topic in enumerate(time_topics):
        row_idx = 197 + i
        aa_rows[row_idx][3] = topic
        aa_rows[row_idx][4] = round(rng.uniform(5, 25), 1)   # Q3 (already %)
        aa_rows[row_idx][7] = round(rng.uniform(5, 25), 1)   # Q4 (already %)

    aa_df = pd.DataFrame(aa_rows)

    # Write
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        ryb_df.to_excel(writer, sheet_name="RYB", index=False, header=False)
        tag_df.to_excel(writer, sheet_name="TAG", index=False, header=False)
        aa_df.to_excel(writer, sheet_name="Additonal Analysis", index=False, header=False)

    print(f"Created: {out_path}")
    print(f"  RYB: {len(ryb_df)} rows x {len(ryb_df.columns)} cols")
    print(f"  TAG: {len(tag_df)} rows x {len(tag_df.columns)} cols")
    print(f"  AA:  {len(aa_df)} rows x {len(aa_df.columns)} cols")


if __name__ == "__main__":
    # Wave 1
    create_dummy_excel(
        os.path.join(BASE, "test_project/data/WAVE_Q1Q2_2026/source_data.xlsx"),
        seed=42,
    )
    # Wave 2 (slightly different data)
    create_dummy_excel(
        os.path.join(BASE, "test_project/data/WAVE_Q3Q4_2026/source_data.xlsx"),
        seed=99,
    )
    print("\nDone — both waves created.")
