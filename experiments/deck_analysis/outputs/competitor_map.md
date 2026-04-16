# Competitor Detection — Review Doc

Auto-generated candidate competitors for each PET brand.
Review and cherry-pick valid matches before merging into `BRAND{}`.

**Score interpretation:**
- `>= 1.0` — exact color match + same therapy area + frequent minor color
- `0.5 - 1.0` — strong signal (near match or secondary rank)
- `< 0.5` — weak / coincidental, usually ignore

---

## ABILIFY_MAINTENA
- **Client:** OTSUKA
- **Therapy area:** Psychiatry
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **ALL_PET_KPI** | UNKNOWN | Cross-brand | 0.28 | `#2A4C85`, `#2D3250`, `#574964` | 1 |
| 2 | **OTEZLA** | AMGEN | Dermatology/Psoriasis | 0.27 | `#2A4C85`, `#2D3250`, `#3B5249` | 1 |
| 3 | **CCA_PP_TRACKER** | CCA | Cross-brand | 0.24 | `#6290D4`, `#7AB2B2`, `#8998B1` | 1 |
| 4 | **TEPEZZA** | JJ | Endocrinology/TED | 0.24 | `#6290D4`, `#7AB2B2`, `#8998B1` | 1 |
| 5 | **DARZALEX** | JJ | Oncology/MM | 0.20 | `#6290D4`, `#7AB2B2`, `#8998B1` | 1 |

---

## ABRYSVO
- **Client:** PFIZER
- **Therapy area:** Vaccines/Maternal
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **ABILIFY_MAINTENA** | OTSUKA | Psychiatry | 0.21 | `#4060AF`, `#F3CE0A`, `#FFC000` | 1 |
| 2 | **LYNPARZA** | AZN | Oncology | 0.15 | `#92D050`, `#F3CE0A`, `#FFC000` | 1 |
| 3 | **TEZSPIRE** | AZN | Respiratory | 0.08 | `#024A34`, `#99D5FF` | 1 |
| 4 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.06 | `#D95776`, `#FFC000` | 1 |
| 5 | **ALL_PET_KPI** | UNKNOWN | Cross-brand | 0.06 | `#4060AF` | 1 |

---

## ADBRY
- **Client:** LEO
- **Therapy area:** Dermatology/AD
- **Decks analyzed:** 2

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **OTEZLA** | AMGEN | Dermatology/Psoriasis | 1.08 | `#00609C`, `#4E7BA9`, `#68A17C` | 2 |
| 2 | **ABILIFY_MAINTENA** | OTSUKA | Psychiatry | 0.29 | `#0047BB`, `#00609C`, `#4E7BA9` | 2 |
| 3 | **CCA_PP_TRACKER** | CCA | Cross-brand | 0.27 | `#9A7CB8`, `#EAAA00`, `#FF8C19` | 2 |
| 4 | **BONE_HCP_TRACKER** | UNKNOWN | Bone | 0.23 | `#EAAA00`, `#FF8C19` | 2 |
| 5 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.23 | `#EAAA00`, `#FF8C19` | 2 |

**Top match reasoning (OTEZLA):**
- #4E7BA9 → #6F8DB4 (color=loose (39), candidate_rank=secondary, ta_overlap=0.5, minor_freq=2% (2/110))
- #929292 → #6F8DB4 (color=loose (49), candidate_rank=secondary, ta_overlap=0.5, minor_freq=2% (2/110))
- #9A7CB8 → #6F8DB4 (color=loose (46), candidate_rank=secondary, ta_overlap=0.5, minor_freq=2% (2/110))

---

## BAVENCIO
- **Client:** EMD
- **Therapy area:** Oncology
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **LYNPARZA** | AZN | Oncology | 0.39 | `#B6CA20`, `#FDCE07` | 1 |
| 2 | **JEMPERLI-ZEJULA** | GSK | Oncology/Endometrial | 0.33 | `#00857C`, `#009E96`, `#00A296` | 1 |
| 3 | **LOKELMA** | AZN | Nephrology | 0.24 | `#00857C`, `#009E96`, `#00A296` | 1 |
| 4 | **RHAPSIDO** | NOVARTIS | Immunology | 0.22 | `#00857C`, `#009E96`, `#00A296` | 1 |
| 5 | **ULTOMIRIS** | ALEXION | Hematology/PNH | 0.20 | `#00857C`, `#009E96`, `#00A296` | 1 |

---

## BLENREP
- **Client:** GSK
- **Therapy area:** Oncology/MM
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.56 | `#E53473`, `#F06A5D`, `#F36633` | 1 |
| 2 | **DATROWAY** | DSI_AZN | Oncology/mBC | 0.49 | `#183978`, `#244EA2`, `#F06A5D` | 1 |
| 3 | **JEMPERLI-ZEJULA** | GSK | Oncology/Endometrial | 0.49 | `#D0400C`, `#D51900` | 1 |
| 4 | **RYBREVANT** | JJ | Oncology/NSCLC | 0.47 | `#D0400C`, `#F36633`, `#FAC2AD` | 1 |
| 5 | **CALQUENCE** | AZN | Oncology/CLL | 0.35 | `#0070C0`, `#0090FF`, `#183978` | 1 |

**Top match reasoning (ENHERTU):**
- #E53473 → #E15759 (color=loose (44), candidate_rank=secondary, ta_overlap=0.5, minor_freq=2% (5/256))
- #F36633 → #FF8813 (color=loose (48), candidate_rank=primary, ta_overlap=0.5, minor_freq=3% (8/256))
- #F36633 → #E15759 (color=loose (45), candidate_rank=secondary, ta_overlap=0.5, minor_freq=3% (8/256))

---

## BONE_HCP_TRACKER
- **Client:** UNKNOWN
- **Therapy area:** Bone
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **CCA_PP_TRACKER** | CCA | Cross-brand | 0.18 | `#93CDDD`, `#AAA8AB`, `#F48323` | 1 |
| 2 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.11 | `#F48323`, `#F79646` | 1 |
| 3 | **DATROWAY** | DSI_AZN | Oncology/mBC | 0.10 | `#2C53A2`, `#F48323`, `#F79646` | 1 |
| 4 | **OTEZLA** | AMGEN | Dermatology/Psoriasis | 0.08 | `#2C53A2`, `#5A82D2`, `#8F77AD` | 1 |
| 5 | **BLENREP** | GSK | Oncology/MM | 0.08 | `#2C53A2`, `#57AC33`, `#8F77AD` | 1 |

---

## CALQUENCE
- **Client:** AZN
- **Therapy area:** Oncology/CLL
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.36 | `#C93535`, `#FABE00`, `#FFC000` | 1 |
| 2 | **RYBREVANT** | JJ | Oncology/NSCLC | 0.27 | `#EDA182`, `#F4C7B4`, `#FFDD80` | 1 |
| 3 | **LYNPARZA** | AZN | Oncology | 0.25 | `#FABE00`, `#FFC000` | 1 |
| 4 | **ABILIFY_MAINTENA** | OTSUKA | Psychiatry | 0.17 | `#FABE00`, `#FFC000` | 1 |
| 5 | **DARZALEX** | JJ | Oncology/MM | 0.14 | `#8EB1B5` | 1 |

---

## CCA_PP_TRACKER
- **Client:** CCA
- **Therapy area:** Cross-brand
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **DUPIXENT_EoE** | REGENERON_SANOFI | Immunology/EoE | 0.08 | `#F28E41`, `#F8696B`, `#FA9473` | 1 |
| 2 | **BONE_HCP_TRACKER** | UNKNOWN | Bone | 0.07 | `#F28E41` | 1 |
| 3 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.06 | `#F28E41`, `#F8696B` | 1 |
| 4 | **ONIVYDE** | IPSEN | Oncology/Pancreatic | 0.06 | `#63BE7B`, `#F8696B` | 1 |
| 5 | **RYBREVANT** | JJ | Oncology/NSCLC | 0.06 | `#FA9473`, `#FCBF7B` | 1 |

---

## DARZALEX
- **Client:** JJ
- **Therapy area:** Oncology/MM
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **RYBREVANT** | JJ | Oncology/NSCLC | 0.84 | `#E27E2E`, `#ED7D31`, `#F49AC2` | 1 |
| 2 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.78 | `#E27E2E`, `#ED7D31`, `#EEB500` | 1 |
| 3 | **CALQUENCE** | AZN | Oncology/CLL | 0.54 | `#0070C0`, `#00B5E2`, `#0FB4FF` | 1 |
| 4 | **LYNPARZA** | AZN | Oncology | 0.45 | `#7ED25C`, `#D7D200`, `#EEB500` | 1 |
| 5 | **JEMPERLI-ZEJULA** | GSK | Oncology/Endometrial | 0.34 | `#007179`, `#1E566D`, `#D81F1F` | 1 |

**Top match reasoning (RYBREVANT):**
- #FFA482 → #FBAB91 (color=near (17), candidate_rank=secondary, ta_overlap=0.5, minor_freq=2% (5/213))
- #E27E2E → #F75824 (color=loose (45), candidate_rank=primary, ta_overlap=0.5, minor_freq=1% (2/213))
- #F75824 → #F75824 (color=exact, candidate_rank=primary, ta_overlap=0.5, minor_freq=1% (2/213))

---

## DATROWAY
- **Client:** DSI_AZN
- **Therapy area:** Oncology/mBC
- **Decks analyzed:** 2

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **LIBTAYO** | REGENERON | Oncology/NMSC | 0.71 | `#002060`, `#033585`, `#1E5680` | 2 |
| 2 | **TRUQAP** | AZN | Oncology/Breast | 0.67 | `#002060`, `#250E62`, `#662D91` | 2 |
| 3 | **LYNPARZA** | AZN | Oncology | 0.66 | `#7CD214`, `#92D050`, `#F0AB00` | 2 |
| 4 | **CALQUENCE** | AZN | Oncology/CLL | 0.48 | `#002060`, `#00D1D6`, `#033585` | 2 |
| 5 | **DARZALEX** | JJ | Oncology/MM | 0.44 | `#68D2DF`, `#75E6E9`, `#79D1A1` | 2 |

**Top match reasoning (LIBTAYO):**
- #F73571 → #DC0077 (color=loose (60), candidate_rank=primary, ta_overlap=0.5, minor_freq=4% (4/108))
- #033585 → #004F6F (color=loose (34), candidate_rank=secondary, ta_overlap=0.5, minor_freq=3% (3/108))
- #304B4C → #004F6F (color=loose (60), candidate_rank=secondary, ta_overlap=0.5, minor_freq=3% (3/115))

---

## DUPIXENT_EoE
- **Client:** REGENERON_SANOFI
- **Therapy area:** Immunology/EoE
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **DARZALEX** | JJ | Oncology/MM | 0.29 | `#53A9A7`, `#59AFAD`, `#76C7E9` | 1 |
| 2 | **TEPEZZA** | JJ | Endocrinology/TED | 0.27 | `#0070C0`, `#76C7E9`, `#90C9C8` | 1 |
| 3 | **CCA_PP_TRACKER** | CCA | Cross-brand | 0.16 | `#76C7E9`, `#90C9C8`, `#90CAC8` | 1 |
| 4 | **TEZSPIRE** | AZN | Respiratory | 0.16 | `#2C4D75`, `#76C7E9`, `#90C9C8` | 1 |
| 5 | **OTEZLA** | AMGEN | Dermatology/Psoriasis | 0.15 | `#2C4D75`, `#53A9A7`, `#59AFAD` | 1 |

---

## EMPAVELI
- **Client:** APELLIS
- **Therapy area:** Hematology
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **ULTOMIRIS** | ALEXION | Hematology/PNH | 0.28 | `#00628B`, `#219491` | 1 |
| 2 | **DUPIXENT_EoE** | REGENERON_SANOFI | Immunology/EoE | 0.12 | `#219491`, `#F16464` | 1 |
| 3 | **TEPEZZA** | JJ | Endocrinology/TED | 0.12 | `#0059A6`, `#00628B`, `#0070C0` | 1 |
| 4 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.12 | `#C02C57`, `#D36B8A`, `#F16464` | 1 |
| 5 | **CALQUENCE** | AZN | Oncology/CLL | 0.12 | `#0070C0`, `#00B0F0` | 1 |

---

## ENHERTU
- **Client:** DSI_AZN
- **Therapy area:** Oncology/HER2
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **RYBREVANT** | JJ | Oncology/NSCLC | 1.12 | `#D99694`, `#EB8D8F`, `#EE7623` | 1 |
| 2 | **ONIVYDE** | IPSEN | Oncology/Pancreatic | 0.62 | `#34A355`, `#489F33`, `#4EBB57` | 1 |
| 3 | **BLENREP** | GSK | Oncology/MM | 0.43 | `#34A355`, `#397AA7`, `#489F33` | 1 |
| 4 | **DATROWAY** | DSI_AZN | Oncology/mBC | 0.32 | `#EE7623`, `#FF6600`, `#FF7500` | 1 |
| 5 | **JEMPERLI-ZEJULA** | GSK | Oncology/Endometrial | 0.29 | `#004F8E`, `#005AA4`, `#C72327` | 1 |

**Top match reasoning (RYBREVANT):**
- #FF6600 → #F75824 (color=loose (39), candidate_rank=primary, ta_overlap=0.5, minor_freq=0% (1/628))
- #D99694 → #FBAB91 (color=loose (40), candidate_rank=secondary, ta_overlap=0.5, minor_freq=1% (8/628))
- #FFB871 → #FBAB91 (color=loose (35), candidate_rank=secondary, ta_overlap=0.5, minor_freq=0% (1/628))

---

## JEMPERLI-ZEJULA
- **Client:** GSK
- **Therapy area:** Oncology/Endometrial
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **RYBREVANT** | JJ | Oncology/NSCLC | 0.40 | `#EA8D85`, `#F1B3AD`, `#F36633` | 1 |
| 2 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.22 | `#F36633` | 1 |
| 3 | **CALQUENCE** | AZN | Oncology/CLL | 0.15 | `#0070C0` | 1 |
| 4 | **DATROWAY** | DSI_AZN | Oncology/mBC | 0.13 | `#F36633` | 1 |
| 5 | **TEPEZZA** | JJ | Endocrinology/TED | 0.08 | `#0070C0`, `#ABC1EC` | 1 |

---

## LIBTAYO
- **Client:** REGENERON
- **Therapy area:** Oncology/NMSC
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.40 | `#EFAC30`, `#FF8205` | 1 |
| 2 | **LYNPARZA** | AZN | Oncology | 0.29 | `#7CBF33`, `#EFAC30`, `#FF8205` | 1 |
| 3 | **RYBREVANT** | JJ | Oncology/NSCLC | 0.22 | `#FF8205`, `#FFB469` | 1 |
| 4 | **DARZALEX** | JJ | Oncology/MM | 0.15 | `#99B3C7` | 1 |
| 5 | **DATROWAY** | DSI_AZN | Oncology/mBC | 0.14 | `#EFAC30`, `#FF8205` | 1 |

---

## LOKELMA
- **Client:** AZN
- **Therapy area:** Nephrology
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **ABILIFY_MAINTENA** | OTSUKA | Psychiatry | 0.13 | `#0070C0`, `#29629E`, `#FFB400` | 1 |
| 2 | **DUPIXENT_EoE** | REGENERON_SANOFI | Immunology/EoE | 0.12 | `#0099B3`, `#00AE9F`, `#29629E` | 1 |
| 3 | **ULTOMIRIS** | ALEXION | Hematology/PNH | 0.10 | `#00AE9F`, `#29629E`, `#594371` | 1 |
| 4 | **CALQUENCE** | AZN | Oncology/CLL | 0.10 | `#00493E`, `#0070C0`, `#0099B3` | 1 |
| 5 | **JEMPERLI-ZEJULA** | GSK | Oncology/Endometrial | 0.08 | `#0099B3`, `#00AE9F`, `#FF2916` | 1 |

---

## LYNPARZA
- **Client:** AZN
- **Therapy area:** Oncology
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **RYBREVANT** | JJ | Oncology/NSCLC | 0.41 | `#F2D076`, `#F7931D`, `#FA8F91` | 1 |
| 2 | **DARZALEX** | JJ | Oncology/MM | 0.41 | `#72CDF6`, `#85C899`, `#C400FA` | 1 |
| 3 | **ONIVYDE** | IPSEN | Oncology/Pancreatic | 0.35 | `#6AD66D`, `#71DB71`, `#A2668B` | 1 |
| 4 | **BLENREP** | GSK | Oncology/MM | 0.32 | `#2CA02C`, `#916CBC` | 1 |
| 5 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.25 | `#F7931D` | 1 |

---

## MIEBO
- **Client:** BL
- **Therapy area:** Ophthalmology/DED
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **CALQUENCE** | AZN | Oncology/CLL | 0.31 | `#0078BC`, `#007FB0`, `#0094AA` | 1 |
| 2 | **ULTOMIRIS** | ALEXION | Hematology/PNH | 0.29 | `#007DAD`, `#007FB0`, `#0094AA` | 1 |
| 3 | **TEPEZZA** | JJ | Endocrinology/TED | 0.26 | `#006EB9`, `#0078BC`, `#007DAD` | 1 |
| 4 | **ONIVYDE** | IPSEN | Oncology/Pancreatic | 0.25 | `#3EB07A`, `#4DA64D`, `#4DC885` | 1 |
| 5 | **RHAPSIDO** | NOVARTIS | Immunology | 0.24 | `#0078BC`, `#007DAD`, `#007FB0` | 1 |

---

## OJJAARA
- **Client:** GSK
- **Therapy area:** Hematology/MF
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **ULTOMIRIS** | ALEXION | Hematology/PNH | 0.12 | `#5F269E` | 1 |
| 2 | **BLENREP** | GSK | Oncology/MM | 0.09 | `#4DBC32`, `#5F269E` | 1 |
| 3 | **RYBREVANT** | JJ | Oncology/NSCLC | 0.06 | `#F49489`, `#F8BAB3` | 1 |
| 4 | **ONIVYDE** | IPSEN | Oncology/Pancreatic | 0.04 | `#4DBC32` | 1 |
| 5 | **LYNPARZA** | AZN | Oncology | 0.04 | `#9BBB59` | 1 |

---

## ONIVYDE
- **Client:** IPSEN
- **Therapy area:** Oncology/Pancreatic
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.48 | `#DE505E`, `#FF8813` | 1 |
| 2 | **BLENREP** | GSK | Oncology/MM | 0.25 | `#3FAE2A` | 1 |
| 3 | **RYBREVANT** | JJ | Oncology/NSCLC | 0.14 | `#FF8813` | 1 |
| 4 | **DARZALEX** | JJ | Oncology/MM | 0.14 | `#8AD79B` | 1 |
| 5 | **CALQUENCE** | AZN | Oncology/CLL | 0.14 | `#00C9E8` | 1 |

---

## OTEZLA
- **Client:** AMGEN
- **Therapy area:** Dermatology/Psoriasis
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **TEPEZZA** | JJ | Endocrinology/TED | 0.26 | `#2962A7`, `#4DBFE9`, `#8FAADC` | 1 |
| 2 | **DUPIXENT_EoE** | REGENERON_SANOFI | Immunology/EoE | 0.20 | `#006C75`, `#00717B`, `#2962A7` | 1 |
| 3 | **CCA_PP_TRACKER** | CCA | Cross-brand | 0.19 | `#8272E8`, `#8FAADC`, `#99DAF2` | 1 |
| 4 | **ALL_PET_KPI** | UNKNOWN | Cross-brand | 0.16 | `#006C75`, `#00717B`, `#2962A7` | 1 |
| 5 | **JEMPERLI-ZEJULA** | GSK | Oncology/Endometrial | 0.14 | `#006C75`, `#00717B`, `#C40E12` | 1 |

---

## PHYSICIANS_SFE_PET
- **Client:** UNKNOWN
- **Therapy area:** Cross-brand
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **ALL_PET_KPI** | UNKNOWN | Cross-brand | 0.30 | `#004D72`, `#6694AA` | 1 |
| 2 | **LYNPARZA** | AZN | Oncology | 0.11 | `#92D050`, `#FFC000` | 1 |
| 3 | **ABILIFY_MAINTENA** | OTSUKA | Psychiatry | 0.10 | `#FFC000` | 1 |
| 4 | **CCA_PP_TRACKER** | CCA | Cross-brand | 0.10 | `#B2B2B2` | 1 |
| 5 | **RYBREVANT** | JJ | Oncology/NSCLC | 0.10 | `#F7B0C0`, `#FF79A6`, `#FF9694` | 1 |

---

## RHAPSIDO
- **Client:** NOVARTIS
- **Therapy area:** Immunology
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **DUPIXENT_EoE** | REGENERON_SANOFI | Immunology/EoE | 0.50 | `#0C68B0`, `#168376`, `#4EB0AB` | 1 |
| 2 | **TEPEZZA** | JJ | Endocrinology/TED | 0.49 | `#0460A9`, `#0579D9`, `#0C68B0` | 1 |
| 3 | **CCA_PP_TRACKER** | CCA | Cross-brand | 0.25 | `#7BC3FC`, `#7FB7FE`, `#80C6C2` | 1 |
| 4 | **TEZSPIRE** | AZN | Respiratory | 0.25 | `#001750`, `#03487F`, `#4FA5FF` | 1 |
| 5 | **UPLIZNA** | AMGEN | Neurology/NMOSD | 0.22 | `#03487F`, `#168376`, `#842064` | 1 |

---

## RYBREVANT
- **Client:** JJ
- **Therapy area:** Oncology/NSCLC
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **LIBTAYO** | REGENERON | Oncology/NMSC | 0.38 | `#0B486B`, `#DC0086` | 1 |
| 2 | **LYNPARZA** | AZN | Oncology | 0.27 | `#9EDC70`, `#FFC000` | 1 |
| 3 | **ONIVYDE** | IPSEN | Oncology/Pancreatic | 0.22 | `#3B8686`, `#CC3399` | 1 |
| 4 | **JEMPERLI-ZEJULA** | GSK | Oncology/Endometrial | 0.22 | `#3B8686`, `#BF3119` | 1 |
| 5 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.22 | `#FF6A5A`, `#FFC000` | 1 |

---

## TEPEZZA
- **Client:** JJ
- **Therapy area:** Endocrinology/TED
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **TEZSPIRE** | AZN | Respiratory | 0.19 | `#151F6D`, `#73AADE`, `#7EE0DD` | 1 |
| 2 | **CCA_PP_TRACKER** | CCA | Cross-brand | 0.18 | `#73AADE`, `#7EE0DD`, `#80D1E2` | 1 |
| 3 | **RYBREVANT** | JJ | Oncology/NSCLC | 0.14 | `#C86B2D`, `#D9C5BB`, `#FF7C80` | 1 |
| 4 | **TRUQAP** | AZN | Oncology/Breast | 0.13 | `#151F6D`, `#605D75`, `#6C8495` | 1 |
| 5 | **DARZALEX** | JJ | Oncology/MM | 0.12 | `#73AADE`, `#7EE0DD`, `#80D1E2` | 1 |

---

## TEZSPIRE
- **Client:** AZN
- **Therapy area:** Respiratory
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **ABILIFY_MAINTENA** | OTSUKA | Psychiatry | 0.12 | `#DCD828`, `#F5CD39`, `#F89C0D` | 1 |
| 2 | **TEPEZZA** | JJ | Endocrinology/TED | 0.12 | `#68D2DF`, `#7F9BB2`, `#96E0E9` | 1 |
| 3 | **CCA_PP_TRACKER** | CCA | Cross-brand | 0.10 | `#68D2DF`, `#7F9BB2`, `#96E0E9` | 1 |
| 4 | **BONE_HCP_TRACKER** | UNKNOWN | Bone | 0.08 | `#F5CD39`, `#F89C0D` | 1 |
| 5 | **BLENREP** | GSK | Oncology/MM | 0.08 | `#13DF22`, `#5D9D39` | 1 |

---

## TRUQAP
- **Client:** AZN
- **Therapy area:** Oncology/Breast
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **ONIVYDE** | IPSEN | Oncology/Pancreatic | 0.50 | `#34CA8B`, `#408F69`, `#719743` | 1 |
| 2 | **DATROWAY** | DSI_AZN | Oncology/mBC | 0.50 | `#1E22AA`, `#461BBF` | 1 |
| 3 | **DARZALEX** | JJ | Oncology/MM | 0.42 | `#7FB49B`, `#93E59D`, `#99C3AF` | 1 |
| 4 | **LYNPARZA** | AZN | Oncology | 0.40 | `#92D050`, `#D6BC22`, `#F3BD35` | 1 |
| 5 | **BAVENCIO** | EMD | Oncology | 0.34 | `#1E22AA`, `#3E4062`, `#5C4A89` | 1 |

**Top match reasoning (ONIVYDE):**
- #34CA8B → #54AC65 (color=loose (58), candidate_rank=primary, ta_overlap=0.5, minor_freq=2% (6/269))
- #408F69 → #54AC65 (color=loose (35), candidate_rank=primary, ta_overlap=0.5, minor_freq=0% (1/269))
- #993366 → #C84874 (color=loose (53), candidate_rank=secondary, ta_overlap=0.5, minor_freq=0% (1/269))

---

## ULTOMIRIS
- **Client:** ALEXION
- **Therapy area:** Hematology/PNH
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **OJJAARA** | GSK | Hematology/MF | 0.42 | `#5C922A`, `#85CA46`, `#9BBB59` | 1 |
| 2 | **ABILIFY_MAINTENA** | OTSUKA | Psychiatry | 0.32 | `#20639B`, `#F8CF36`, `#FAB505` | 1 |
| 3 | **RYBREVANT** | JJ | Oncology/NSCLC | 0.30 | `#E47E19`, `#E9A697`, `#ED553B` | 1 |
| 4 | **CCA_PP_TRACKER** | CCA | Cross-brand | 0.30 | `#A5ACA5`, `#B2A4B6`, `#B6B6B6` | 1 |
| 5 | **EMPAVELI** | APELLIS | Hematology | 0.28 | `#16D0BA`, `#16D8C1`, `#ED553B` | 1 |

---

## UPLIZNA
- **Client:** AMGEN
- **Therapy area:** Neurology/NMOSD
- **Decks analyzed:** 1

| Rank | Candidate | Client | TA | Score | Matched Colors | Decks |
|------|-----------|--------|----|-------|-----------------|-------|
| 1 | **OTEZLA** | AMGEN | Dermatology/Psoriasis | 0.24 | `#003161`, `#082F64`, `#0A3D80` | 1 |
| 2 | **TEZSPIRE** | AZN | Respiratory | 0.24 | `#003161`, `#082F64`, `#0A3D80` | 1 |
| 3 | **ALL_PET_KPI** | UNKNOWN | Cross-brand | 0.16 | `#003161`, `#082F64`, `#0A3D80` | 1 |
| 4 | **ENHERTU** | DSI_AZN | Oncology/HER2 | 0.12 | `#DB2764`, `#ED7D31`, `#FFC000` | 1 |
| 5 | **ABILIFY_MAINTENA** | OTSUKA | Psychiatry | 0.12 | `#0A3D80`, `#FFC000` | 1 |

---
