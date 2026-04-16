# Layout Cluster Analysis

Shape signatures relaxed to count only content shapes (chart / table / picture). Decorative elements (auto_shape, line, freeform) ignored — they vary per client but don't define layout.

## Content Signature Distribution (top 20)

| Signature | Slides |
|---|---|
| `empty` | 291 |
| `1_chart_1_table` | 145 |
| `1_table` | 139 |
| `1_chart_2_table` | 110 |
| `1_chart_3_table` | 77 |
| `2_chart_2_table` | 68 |
| `2_chart_1_table` | 55 |
| `3_chart_1_table` | 45 |
| `2_chart` | 44 |
| `3_picture` | 40 |
| `1_chart_4_table` | 36 |
| `1_picture` | 31 |
| `2_chart_4_table` | 31 |
| `1_chart` | 29 |
| `2_chart_3_table` | 29 |
| `3_chart_3_table` | 28 |
| `3_chart_2_table` | 25 |
| `1_chart_1_table_1_picture` | 23 |
| `3_chart_4_table` | 22 |
| `6_picture` | 21 |

## Per-Signature Coordinate Stats (top 12)

### `1_chart_1_table` — 145 slides

| Shape Type | N observed | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `chart` | 145 | 1.62" | 2.02" | 5.41" | 4.29" |
| `table` | 145 | 1.48" | 1.82" | 6.79" | 4.55" |
| `text_box` | 1826 | 4.62" | 3.7" | 1.52" | 0.27" |

### `1_table` — 139 slides

| Shape Type | N observed | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `table` | 139 | 0.91" | 1.52" | 11.11" | 4.92" |
| `text_box` | 431 | 1.81" | 1.51" | 3.08" | 0.35" |

### `1_chart_2_table` — 110 slides

| Shape Type | N observed | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `chart` | 110 | 6.75" | 2.04" | 3.81" | 4.3" |
| `table` | 220 | 3.11" | 1.98" | 3.88" | 4.17" |
| `text_box` | 1260 | 6.46" | 2.54" | 1.57" | 0.29" |

### `1_chart_3_table` — 77 slides

| Shape Type | N observed | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `chart` | 77 | 7.76" | 2.15" | 3.59" | 4.34" |
| `table` | 231 | 4.71" | 2.06" | 3.32" | 3.92" |
| `text_box` | 934 | 6.93" | 2.96" | 1.38" | 0.29" |

### `2_chart_2_table` — 68 slides

| Shape Type | N observed | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `chart` | 136 | 6.84" | 2.42" | 3.0" | 3.69" |
| `table` | 136 | 2.6" | 2.25" | 2.92" | 3.66" |
| `text_box` | 928 | 5.57" | 3.77" | 1.57" | 0.3" |

### `2_chart_1_table` — 55 slides

| Shape Type | N observed | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `chart` | 110 | 6.67" | 2.18" | 3.15" | 3.99" |
| `table` | 55 | 1.59" | 2.02" | 6.25" | 4.2" |
| `text_box` | 770 | 4.94" | 4.05" | 1.43" | 0.27" |

### `3_chart_1_table` — 45 slides

| Shape Type | N observed | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `chart` | 135 | 7.02" | 2.36" | 2.57" | 3.87" |
| `table` | 45 | 0.39" | 2.17" | 3.85" | 3.76" |
| `text_box` | 858 | 8.25" | 4.06" | 1.26" | 0.27" |

### `2_chart` — 44 slides

| Shape Type | N observed | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `chart` | 88 | 2.4" | 2.33" | 4.19" | 3.67" |
| `text_box` | 819 | 5.45" | 3.53" | 1.3" | 0.3" |

### `3_picture` — 40 slides

| Shape Type | N observed | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `picture` | 120 | 0.92" | 2.88" | 0.5" | 0.51" |
| `text_box` | 246 | 2.04" | 2.95" | 3.06" | 0.4" |

### `1_chart_4_table` — 36 slides

| Shape Type | N observed | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `chart` | 36 | 6.7" | 2.12" | 4.58" | 4.36" |
| `table` | 144 | 5.71" | 2.12" | 2.16" | 3.53" |
| `text_box` | 480 | 8.46" | 3.21" | 1.17" | 0.29" |

### `1_picture` — 31 slides

| Shape Type | N observed | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `picture` | 31 | 2.88" | 1.76" | 2.39" | 0.95" |
| `text_box` | 72 | 6.0" | 3.37" | 2.31" | 0.45" |

---

## Chart Position Clusters

Clustered 410 distinct chart positions (grouping similar-position charts across all decks, tolerance 0.5 inch).

Each row here is a candidate `LAYOUTS{}` preset — a recurring chart position in client decks.

| # | Occurrences | Left | Top | Width | Height | Decks |
|---|---|---|---|---|---|---|
| 1 | 16 | 8.93" | 2.06" | 1.58" | 4.57" | JJ PET RYBREVANT+LAZCLUZE, J&J MM SFEA Q1'26_Full_Re, US Abrysvo Maternal HCP C |
| 2 | 16 | 5.44" | 2.41" | 2.51" | 4.04" | TRUQAP PET Quarterly Repo, [ZoomRx] AZN LOKELMA PET , ZoomRx - AZ - DSI ENHERTU |
| 3 | 15 | 7.95" | 1.94" | 2.1" | 4.53" | JJ PET RYBREVANT+LAZCLUZE, [ZoomRx] Q1'26 Otezla Der, ZoomRx Dupixent EoE PET Q |
| 4 | 13 | 8.38" | 2.08" | 1.44" | 4.53" | [ZoomRx] AZN LOKELMA PET , J&J MM SFEA Q1'26_Full_Re, [ZoomRx] B+L DED PET Jul- |
| 5 | 13 | 0.49" | 1.95" | 4.29" | 4.16" | DATROWAY EGFRm NSCLC Prom, OG Adbry PET SC Deck.pptx, V1_Adbry PET_W1'26 (Wave  |
| 6 | 12 | 8.6" | 2.17" | 3.17" | 4.45" | [ZoomRx] B+L DED PET Jul-, Physicians SFE (PET)_Q3'2, ZoomRx Dupixent EoE PET Q |
| 7 | 12 | 1.88" | 2.37" | 3.44" | 0.88" | Physicians SFE (PET)_Q3'2 |
| 8 | 11 | 11.85" | 2.08" | 1.44" | 4.53" | [ZoomRx] AZN LOKELMA PET , V1_Adbry PET_W1'26 (Wave , [ZoomRx] Alexion US PNH Q |
| 9 | 11 | 11.08" | 2.01" | 1.97" | 4.65" | [ZoomRx] Q1'26 Otezla Der, Physicians SFE (PET)_Q3'2, [ZoomRx] NVS Rhapsido_PET |
| 10 | 11 | 7.66" | 3.0" | 2.15" | 3.58" | TRUQAP PET Quarterly Repo, [ZoomRx] Tezspire SUA+CRS, TEPEZZA_PET_Q4 2025_Repor |
| 11 | 10 | 5.56" | 1.78" | 2.1" | 4.53" | [ZoomRx] Q1'26 Otezla Der, J&J MM SFEA Q1'26_Full_Re, [ZoomRx] Abilify LAI PET  |
| 12 | 10 | 8.48" | 1.93" | 3.9" | 4.37" | [ZoomRx] B+L DED PET Jul-, ZoomRx_Apellis EMPAVELI P, ZoomRx Dupixent EoE PET Q |
| 13 | 10 | 11.32" | 2.09" | 1.94" | 4.46" | GSK Jemperli Zejula PET F, [ZoomRx] Alexion US PNH Q, J&J MM SFEA Q1'26_Full_Re |
| 14 | 10 | 3.51" | 2.34" | 2.53" | 4.03" | [ZoomRx] B+L DED PET Jul-, [ZoomRx] Q1'26 Otezla Der, Physicians SFE (PET)_Q3'2 |
| 15 | 10 | 8.51" | 2.58" | 2.85" | 3.92" | DATROWAY EGFRm NSCLC Prom, TRUQAP PET Quarterly Repo, [ZoomRx] Tezspire SUA+CRS |
| 16 | 10 | 0.45" | 3.19" | 3.04" | 3.97" | Physicians SFE (PET)_Q3'2 |
| 17 | 9 | 11.5" | 2.09" | 1.57" | 4.53" | US Abrysvo Maternal HCP C, 03.12.2026_Q1'26 Bone HCP, J&J MM SFEA Q1'26_Full_Re |
| 18 | 9 | 10.53" | 2.06" | 1.39" | 4.44" | [ZoomRx] AZN LOKELMA PET , 03.12.2026_Q1'26 Bone HCP, [ZoomRx] Alexion US PNH Q |
| 19 | 9 | 7.02" | 2.04" | 1.97" | 4.46" | Physicians SFE (PET)_Q3'2, DATROWAY EGFRm NSCLC Prom, 03.12.2026_Q1'26 Bone HCP |
| 20 | 9 | 7.57" | 2.16" | 1.9" | 4.42" | J&J MM SFEA Q1'26_Full_Re, [ZoomRx] Abilify LAI PET , [ZoomRx] Q1'26 Otezla Der |
| 21 | 9 | 9.51" | 2.19" | 1.9" | 4.41" | J&J MM SFEA Q1'26_Full_Re, [ZoomRx] Abilify LAI PET , [ZoomRx] Q1'26 Otezla Der |
| 22 | 9 | 10.47" | 2.96" | 2.48" | 3.52" | TEPEZZA_PET_Q4 2025_Repor, [ZoomRx] Alexion US PNH Q, OG Adbry PET SC Deck.pptx |
| 23 | 9 | 9.12" | 2.04" | 1.97" | 4.5" | [ZoomRx] Q1'26 Otezla Der, [ZoomRx] Alexion US PNH Q, [ZoomRx] NVS Rhapsido_PET |
| 24 | 9 | 10.64" | 2.39" | 2.36" | 3.9" | TRUQAP PET Quarterly Repo, [ZoomRx] AZN LOKELMA PET , US Abrysvo Maternal HCP C |
| 25 | 9 | 7.48" | 1.99" | 5.41" | 4.59" | [ZoomRx] NVS Rhapsido_PET |
| 26 | 9 | 10.39" | 4.9" | 2.62" | 1.68" | [ZoomRx] Q1'26 Otezla Der, [ZoomRx] NVS Rhapsido_PET |
| 27 | 9 | 5.8" | 1.87" | 1.9" | 4.48" | JJ PET RYBREVANT+LAZCLUZE, J&J MM SFEA Q1'26_Full_Re, [ZoomRx] Q1'26 Otezla Der |
| 28 | 9 | 7.65" | 2.54" | 1.85" | 4.01" | TEPEZZA_PET_Q4 2025_Repor, ZoomRx_Apellis EMPAVELI P, V1_Adbry PET_W1'26 (Wave  |
| 29 | 9 | 8.04" | 1.15" | 1.24" | 4.83" | DATROWAY EGFRm NSCLC Prom |
| 30 | 9 | 6.54" | 2.29" | 2.41" | 4.02" | DATROWAY EGFRm NSCLC Prom, Physicians SFE (PET)_Q3'2, ZoomRx - AZ - DSI ENHERTU |

## Table Position Clusters

Clustered 490 distinct table positions.

| # | Occurrences | Left | Top | Width | Height | Decks |
|---|---|---|---|---|---|---|
| 1 | 40 | 12.52" | 2.17" | 0.58" | 4.5" | Lynparza aOC_PET_Q1 2026T, TEST [COPY] ZoomRx AMG UP, J&J MM SFEA Q1'26_Full_Re |
| 2 | 22 | 11.46" | 2.08" | 0.47" | 4.58" | TEST [COPY] ZoomRx AMG UP, [ZoomRx] Abilify LAI PET , V1_Adbry PET_W1'26 (Wave  |
| 3 | 21 | 0.23" | 1.34" | 2.31" | 0.28" | 03.12.2026_Q1'26 Bone HCP |
| 4 | 19 | 7.96" | 2.07" | 0.49" | 4.58" | J&J MM SFEA Q1'26_Full_Re, V1_Adbry PET_W1'26 (Wave , ZoomRx Dupixent EoE PET Q |
| 5 | 19 | 7.55" | 2.05" | 1.13" | 0.41" | [ZoomRx] B+L DED PET Jul- |
| 6 | 18 | 10.59" | 2.06" | 0.53" | 4.52" | JJ PET RYBREVANT+LAZCLUZE, J&J MM SFEA Q1'26_Full_Re, [ZoomRx] Q1'26 Otezla Der |
| 7 | 17 | 4.64" | 1.11" | 1.33" | 0.38" | 05022026_Bavencio Promoti |
| 8 | 17 | 12.07" | 2.03" | 0.64" | 4.47" | JJ PET RYBREVANT+LAZCLUZE, J&J MM SFEA Q1'26_Full_Re, Physicians SFE (PET)_Q3'2 |
| 9 | 17 | 0.11" | 1.8" | 13.12" | 0.47" | Physicians SFE (PET)_Q3'2, OG Adbry PET SC Deck.pptx, V1_Adbry PET_W1'26 (Wave  |
| 10 | 16 | 9.6" | 1.94" | 0.44" | 4.54" | JJ PET RYBREVANT+LAZCLUZE, V1_Adbry PET_W1'26 (Wave , [ZoomRx] Q1'26 Otezla Der |
| 11 | 15 | 4.59" | 2.43" | 0.97" | 0.57" | [ZoomRx] B+L DED PET Jul-, OG Adbry PET SC Deck.pptx, 05022026_Bavencio Promoti |
| 12 | 15 | 11.98" | 2.0" | 0.88" | 4.54" | Lynparza aOC_PET_Q1 2026T, J&J MM SFEA Q1'26_Full_Re, ZoomRx_Apellis EMPAVELI P |
| 13 | 15 | 11.22" | 2.5" | 0.97" | 0.57" | [ZoomRx] B+L DED PET Jul-, ZoomRx_Apellis EMPAVELI P, OG Adbry PET SC Deck.pptx |
| 14 | 15 | 5.99" | 2.39" | 0.63" | 4.01" | JJ PET RYBREVANT+LAZCLUZE, ZoomRx_Apellis EMPAVELI P, Physicians SFE (PET)_Q3'2 |
| 15 | 14 | 10.05" | 1.79" | 2.94" | 4.54" | 03.12.2026_Q1'26 Bone HCP, TEST [COPY] ZoomRx AMG UP, [ZoomRx] Q1'26 Otezla Der |
| 16 | 14 | 0.58" | 2.06" | 7.5" | 4.57" | V1_Adbry PET_W1'26 (Wave , AZ-DSI - DATROWAY mBC Pro, [ZoomRx] Alexion US PNH Q |
| 17 | 13 | 0.43" | 1.88" | 12.45" | 5.06" | [ZoomRx] B+L DED PET Jul-, [ZoomRx] Alexion US PNH Q, Physicians SFE (PET)_Q3'2 |
| 18 | 13 | 7.51" | 1.88" | 0.44" | 4.51" | J&J MM SFEA Q1'26_Full_Re, [ZoomRx] Q1'26 Otezla Der, ZoomRx Dupixent EoE PET Q |
| 19 | 13 | 5.12" | 6.45" | 0.96" | 0.4" | [ZoomRx] Libtayo (Nov ’25, [ZoomRx] B+L DED PET Jul-, [ZoomRx] Q1'26 Otezla Der |
| 20 | 13 | 0.24" | 1.78" | 5.2" | 4.64" | [ZoomRx] Q1'26 Otezla Der, Physicians SFE (PET)_Q3'2, ZoomRx_Apellis EMPAVELI P |
| 21 | 13 | 12.02" | 1.56" | 1.31" | 0.15" | DATROWAY EGFRm NSCLC Prom |
| 22 | 12 | 12.0" | 2.38" | 0.63" | 4.38" | [ZoomRx] B+L DED PET Jul-, V1_Adbry PET_W1'26 (Wave , Physicians SFE (PET)_Q3'2 |
| 23 | 12 | 12.67" | 1.56" | 0.66" | 0.15" | DATROWAY EGFRm NSCLC Prom |
| 24 | 12 | 0.11" | 1.7" | 13.12" | 0.48" | Physicians SFE (PET)_Q3'2 |
| 25 | 12 | 6.0" | 1.89" | 1.32" | 0.17" | US Abrysvo Maternal HCP C |
| 26 | 11 | 0.48" | 2.09" | 6.9" | 4.65" | [ZoomRx] Abilify LAI PET , [ZoomRx] B+L DED PET Jul-, 03.12.2026_Q1'26 Bone HCP |
| 27 | 11 | 0.47" | 1.58" | 9.09" | 5.01" | [ZoomRx] Alexion US PNH Q, 05022026_Bavencio Promoti, [ZoomRx] NVS Rhapsido_PET |
| 28 | 11 | 8.79" | 6.45" | 0.9" | 0.4" | [ZoomRx] B+L DED PET Jul-, [ZoomRx] Q1'26 Otezla Der, [ZoomRx] Tezspire SUA+CRS |
| 29 | 11 | 7.96" | 1.99" | 0.84" | 0.28" | [ZoomRx] B+L DED PET Jul-, DATROWAY EGFRm NSCLC Prom |
| 30 | 11 | 0.45" | 2.3" | 3.4" | 4.22" | JJ PET RYBREVANT+LAZCLUZE, TRUQAP PET Quarterly Repo, [ZoomRx] B+L DED PET Jul- |
