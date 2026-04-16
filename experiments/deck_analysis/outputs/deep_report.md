# Deep Deck Analysis Report

**Decks analyzed:** 32
**Total charts:** 4354
**Total headlines detected:** 3569
**Total tables:** 4950

---

## 1. Chart Configuration (OOXML-level)

### 1.1 Chart Tag + Grouping Combinations (top 20)

| Tag / Direction / Grouping | Count |
|---|---|
| `barChart(bar_dir=bar,grouping=clustered)` | 1511 |
| `scatterChart(bar_dir=None,grouping=None)` | 743 |
| `lineChart(bar_dir=None,grouping=standard)` | 595 |
| `barChart(bar_dir=col,grouping=percentStacked)` | 486 |
| `barChart(bar_dir=bar,grouping=percentStacked)` | 324 |
| `barChart(bar_dir=bar,grouping=stacked)` | 305 |
| `doughnutChart(bar_dir=None,grouping=None)` | 138 |
| `barChart(bar_dir=col,grouping=clustered)` | 119 |
| `barChart(bar_dir=col,grouping=stacked)` | 98 |
| `lineChart(bar_dir=None,grouping=stacked)` | 31 |
| `pieChart(bar_dir=None,grouping=None)` | 14 |
| `areaChart(bar_dir=None,grouping=percentStacked)` | 6 |
| `areaChart(bar_dir=None,grouping=stacked)` | 3 |
| `bubbleChart(bar_dir=None,grouping=None)` | 2 |

### 1.2 Bar/Column Gap Width

The `<c:gapWidth>` property controls space between bars. python-pptx does NOT expose this — requires lxml_helpers.

| Gap Width | Count |
|---|---|
| `50` | 393 |
| `100` | 355 |
| `80` | 257 |
| `70` | 230 |
| `150` | 200 |
| `110` | 138 |
| `60` | 114 |
| `75` | 114 |
| `90` | 75 |
| `71` | 70 |
| `61` | 58 |
| `182` | 56 |
| `40` | 49 |
| `57` | 49 |
| `140` | 36 |
| `66` | 34 |
| `214` | 32 |
| `160` | 31 |
| `48` | 28 |
| `130` | 26 |
| `52` | 25 |
| `400` | 21 |
| `68` | 19 |
| `109` | 17 |
| `120` | 16 |
| `62` | 15 |
| `65` | 15 |
| `72` | 15 |
| `95` | 14 |
| `45` | 14 |
| `139` | 14 |
| `122` | 14 |
| `125` | 13 |
| `30` | 13 |
| `33` | 13 |
| `20` | 11 |
| `105` | 11 |
| `92` | 10 |
| `101` | 10 |
| `200` | 9 |
| `86` | 9 |
| `51` | 8 |
| `0` | 8 |
| `35` | 8 |
| `102` | 8 |
| `41` | 8 |
| `83` | 8 |
| `67` | 7 |
| `112` | 7 |
| `219` | 7 |
| `25` | 6 |
| `103` | 6 |
| `21` | 6 |
| `42` | 6 |
| `300` | 6 |
| `350` | 5 |
| `34` | 5 |
| `79` | 5 |
| `43` | 4 |
| `96` | 4 |
| `74` | 4 |
| `44` | 4 |
| `250` | 4 |
| `210` | 4 |
| `132` | 4 |
| `39` | 3 |
| `59` | 3 |
| `78` | 3 |
| `118` | 3 |
| `180` | 3 |
| `85` | 3 |
| `148` | 3 |
| `46` | 3 |
| `106` | 2 |
| `93` | 2 |
| `31` | 2 |
| `54` | 2 |
| `165` | 2 |
| `55` | 2 |
| `164` | 2 |
| `37` | 1 |
| `91` | 1 |
| `108` | 1 |
| `63` | 1 |
| `121` | 1 |
| `128` | 1 |
| `28` | 1 |
| `36` | 1 |
| `88` | 1 |
| `73` | 1 |
| `298` | 1 |
| `220` | 1 |
| `87` | 1 |
| `69` | 1 |
| `410` | 1 |
| `240` | 1 |
| `24` | 1 |
| `325` | 1 |
| `312` | 1 |
| `296` | 1 |

### 1.3 Bar/Column Overlap

| Overlap | Count |
|---|---|
| `100` | 1219 |
| `-20` | 112 |
| `-5` | 53 |
| `-1` | 25 |
| `-50` | 23 |
| `-10` | 16 |
| `-27` | 14 |
| `-100` | 10 |
| `-6` | 6 |
| `-15` | 6 |
| `-78` | 5 |
| `-14` | 4 |
| `-19` | 4 |
| `50` | 3 |
| `-43` | 2 |
| `-35` | 2 |
| `35` | 2 |
| `-66` | 1 |
| `-47` | 1 |
| `-18` | 1 |
| `-46` | 1 |

### 1.4 Data Label Positions

| Position | Count |
|---|---|
| `ctr` | 4508 |
| `t` | 2778 |
| `outEnd` | 2618 |
| `b` | 176 |
| `r` | 166 |
| `inEnd` | 143 |
| `l` | 126 |
| `inBase` | 103 |
| `bestFit` | 2 |

### 1.5 Axis Orientation

`maxMin` = inverted axis (top-down for category axis in bar charts).

| Axis : Orientation | Count |
|---|---|
| `val:minMax` | 4835 |
| `cat:maxMin` | 1865 |
| `cat:minMax` | 1587 |
| `val:maxMin` | 115 |

### 1.6 Tick Label Positions

`none` = hidden axis labels (common when labels shown in separate table).

| Axis : Position | Count |
|---|---|
| `val:nextTo` | 4884 |
| `cat:nextTo` | 3057 |
| `cat:none` | 371 |
| `val:none` | 48 |
| `cat:low` | 24 |
| `val:high` | 18 |

### 1.7 Number Formats (top 20)

| Format | Count |
|---|---|
| `0%` | 4182 |
| `0` | 566 |
| `General` | 292 |
| `#,##0.0` | 59 |
| `0.0` | 50 |
| `#,##0` | 39 |
| `_ * #,##0.0_ ;_ * \-#,##0.0_ ;_ * "-"?_ ;_ @_ ` | 19 |
| `0%;\-0%;\ ` | 14 |
| `[>=0.05]0%;;;` | 13 |
| `0.00%` | 4 |
| `0.00` | 4 |
| `_(* #,##0.0_);_(* \(#,##0.0\);_(* "-"??_);_(@_)` | 3 |
| `#,##0.00` | 1 |
| `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)` | 1 |
| `0.0;[Red]0.0` | 1 |

### 1.8 Legend Positions

| Position | Count |
|---|---|
| `r` | 40 |
| `b` | 6 |

### 1.9 Series Marker Types (scatter/line charts)

| Marker | Count |
|---|---|
| `circle` | 3087 |
| `square` | 128 |
| `triangle` | 99 |
| `none` | 80 |
| `diamond` | 39 |
| `plus` | 20 |
| `star` | 3 |

### 1.10 Series Line Widths (EMU)

12700 EMU = 1pt. 19050 EMU = 1.5pt. 25400 EMU = 2pt.

| Width (EMU) | Count |
|---|---|
| `28575` | 1223 |
| `25400` | 827 |
| `12700` | 607 |
| `19050` | 504 |
| `9525` | 162 |
| `6350` | 18 |
| `15875` | 16 |
| `31750` | 16 |
| `38100` | 10 |
| `3175` | 2 |
| `47625` | 2 |
| `34925` | 2 |
| `22225` | 2 |
| `0` | 1 |
| `41275` | 1 |

### 1.11 Chart Features Summary

- Charts with title: **45** (1%)
- Charts with legend: **46** (1%)
- Charts with major gridlines: **706** (16%)
- Charts with minor gridlines: **11** (0%)
- Charts with trendlines: **1**
- Charts with error bars: **2**

---

## 2. OOXML Property Inventory

Every chart-XML tag seen across all decks. Tags python-pptx covers natively are NOT flagged as gaps — but this list is exhaustive reference for which lxml helpers might be needed.

**235** distinct chart XML tags encountered.

### 2.1 Top 40 Most-Used Tags

| Tag | Occurrences |
|---|---|
| `v` | 109176 |
| `pt` | 107859 |
| `noFill` | 74367 |
| `ext` | 71613 |
| `solidFill` | 56297 |
| `ln` | 55649 |
| `extLst` | 53688 |
| `effectLst` | 51984 |
| `idx` | 42424 |
| `uniqueId` | 42410 |
| `srgbClr` | 39048 |
| `showLegendKey` | 33321 |
| `showVal` | 33321 |
| `showCatName` | 33321 |
| `showSerName` | 33321 |
| `showPercent` | 33321 |
| `showBubbleSize` | 33321 |
| `pPr` | 31144 |
| `p` | 30507 |
| `bodyPr` | 30506 |
| `lstStyle` | 30506 |
| `defRPr` | 29605 |
| `ptCount` | 29317 |
| `dLblPos` | 29253 |
| `latin` | 25874 |
| `endParaRPr` | 24205 |
| `cs` | 24046 |
| `ea` | 22783 |
| `dLbl` | 21803 |
| `txPr` | 21636 |
| `spAutoFit` | 20695 |
| `manualLayout` | 18565 |
| `schemeClr` | 17652 |
| `tx` | 16927 |
| `axId` | 16856 |
| `strRef` | 16020 |
| `strCache` | 16009 |
| `x` | 14708 |
| `y` | 14708 |
| `dLbls` | 14148 |

---

## 3. Brand Extraction by Client

### AZN (7 decks)

**Decks:**
- `[ZoomRx] AZN LOKELMA PET Quarterly Report Q4 2025_Synapse_connector.pptx`
- `[ZoomRx] Tezspire SUA+CRSwNP PET Report Q1 '26_v1.pptx`
- `AZ-DSI - DATROWAY mBC Promotional Effectiveness Tracking Q1 '26_TEST.pptx`
- `Lynparza aOC_PET_Q1 2026TEST.pptx`
- `TRUQAP PET Quarterly Report Q4 '25 v2 (Sandbox v2).pptx`
- `ZoomRx - AZ - DSI ENHERTU NSCLC HER2 ERBB2mu Post Launch Wave 6 Report 3.1.pptx`
- `ZoomRx - AZN Calquence PET Q4 '25 Report v1 - Sandbox Copy.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#00B050` | 77 |
| `#FF8813` | 73 |
| `#94D448` | 53 |
| `#E15759` | 50 |
| `#250E62` | 42 |
| `#33A4FF` | 42 |
| `#59A14F` | 40 |
| `#9467BD` | 40 |
| `#1D7B87` | 40 |
| `#91C3D5` | 37 |
| `#00947F` | 32 |
| `#EE7623` | 31 |
| `#006937` | 31 |
| `#7F7F7F` | 30 |
| `#FFBA00` | 30 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Arial` | 9645 |
| `Calibri` | 3459 |
| `Manrope` | 1350 |
| `Arial(Body)` | 118 |
| `Century Gothic` | 62 |
| `Playfair Display` | 59 |
| `Arial (Headings)` | 32 |
| `+mn-lt` | 22 |

**Theme fonts (major/minor):**

- Major: ['Arial', 'Calibri']
- Minor: ['Arial', 'Calibri']


### JJ (4 decks)

**Decks:**
- `J&J MM SFEA Q1'26_Full_Report_V2_Migration.pptx`
- `JJ PET RYBREVANT+LAZCLUZE Q1'26 Report_Migration.pptx`
- `TEPEZZA_PET_Q4 2025_Report_JJ_NS_TP1latest_UseThis.pptx`
- `ZoomRx - GSK OJJAARA MF PET Research - February '26 Full Monthly Report v1 Migration Test.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#7FB1E1` | 61 |
| `#0063C3` | 50 |
| `#6FC6C1` | 35 |
| `#37CFCA` | 34 |
| `#BD05ED` | 30 |
| `#228D8A` | 29 |
| `#A0D0FF` | 19 |
| `#FF6A5A` | 18 |
| `#2A41C1` | 18 |
| `#CBF3F1` | 18 |
| `#F75824` | 16 |
| `#00A3C4` | 16 |
| `#FFC000` | 14 |
| `#0302C8` | 13 |
| `#ADADAD` | 13 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Johnson Text` | 4918 |
| `Johnson Display` | 702 |
| `+mn-lt` | 475 |
| `Century Gothic` | 457 |
| `+mj-lt` | 451 |
| `Arial` | 352 |

**Theme fonts (major/minor):**

- Major: ['Johnson Display', 'Century Gothic', 'Arial']
- Minor: ['Johnson Text', 'Century Gothic', 'Arial']


### Pfizer (2 decks)

**Decks:**
- `05022026_Bavencio Promotional Effectiveness Tracking Wave 4 Q1'26 - Full Final Report V2.0.pptx`
- `US Abrysvo Maternal HCP CFC Promotional Effectiveness Tracking (ZoomRx)_Q4 ‘25 - Full Report V3.0.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#0000C9` | 30 |
| `#43964A` | 14 |
| `#E356DC` | 12 |
| `#F49C34` | 10 |
| `#BA2A4C` | 9 |
| `#C00000` | 9 |
| `#6D5A7A` | 8 |
| `#4060AF` | 7 |
| `#7F7F7F` | 6 |
| `#081E3D` | 5 |
| `#0095FF` | 5 |
| `#00B050` | 4 |
| `#008A00` | 3 |
| `#808080` | 3 |
| `#6FBF4A` | 3 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Arial` | 1783 |
| `+mj-lt` | 321 |
| `Europa-Regular` | 32 |
| `Calibri (Body)` | 22 |
| `Calibri` | 13 |
| `Manrope` | 8 |
| `Invention(Body)` | 4 |
| `+mn-lt` | 2 |

**Theme fonts (major/minor):**

- Major: ['Arial']
- Minor: ['Arial']


### Regeneron (2 decks)

**Decks:**
- `[ZoomRx] Libtayo (Nov ’25 - Feb ’26) NMSC Message Recall & Effectiveness Report v3.0 - Synapse Migration.pptx`
- `ZoomRx Dupixent EoE PET Q4 '25_v3.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#219491` | 62 |
| `#F79646` | 18 |
| `#E46C0A` | 16 |
| `#DC0077` | 11 |
| `#8EB4E3` | 11 |
| `#00745A` | 11 |
| `#7030A0` | 11 |
| `#004F6F` | 10 |
| `#B370A4` | 10 |
| `#7ABFBD` | 9 |
| `#558ED5` | 9 |
| `#C83333` | 8 |
| `#BFBFBF` | 8 |
| `#96BBBB` | 7 |
| `#C19875` | 7 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Arial` | 3071 |
| `DM Sans 14pt` | 1585 |
| `+mj-lt` | 189 |
| `+mn-lt` | 99 |
| `Arial(Body)` | 51 |
| `Wingdings` | 27 |
| `Calibri` | 22 |
| `Trebuschet MS Body` | 1 |

**Theme fonts (major/minor):**

- Major: ['Trade Gothic LT Std', 'Arial']
- Minor: ['Arial']


### Amgen (2 decks)

**Decks:**
- `[ZoomRx] Q1'26 Otezla Derm PET_Full Report v0.1.pptx`
- `TEST [COPY] ZoomRx AMG UPLIZNA NMOSD PET Q4 '25 Final Report_cleaned.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#1F497D` | 48 |
| `#813F97` | 30 |
| `#007D6D` | 14 |
| `#6F8DB4` | 12 |
| `#003C71` | 11 |
| `#DA205F` | 8 |
| `#00A3DF` | 7 |
| `#4C35DE` | 7 |
| `#0063C3` | 7 |
| `#00B0F0` | 7 |
| `#E46C0A` | 5 |
| `#D9D9D9` | 5 |
| `#A5A5A5` | 4 |
| `#425563` | 4 |
| `#A6A6A6` | 4 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Century Gothic` | 4267 |
| `Calibri` | 214 |
| `+mj-lt` | 118 |
| `+mn-lt` | 74 |
| `Arial` | 3 |
| `Times New Roman` | 1 |

**Theme fonts (major/minor):**

- Major: ['Century Gothic', 'Calibri']
- Minor: ['Century Gothic', 'Arial']


### Unknown (2 decks)

**Decks:**
- `ALL PET Q1FY26 KPI Report 19DEC2025.pptx`
- `Physicians SFE (PET)_Q3'25 Report_V1.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#E7004C` | 37 |
| `#C7A013` | 30 |
| `#6BC9EF` | 25 |
| `#006838` | 24 |
| `#FF6969` | 21 |
| `#00607C` | 20 |
| `#4E79A7` | 18 |
| `#00A44A` | 14 |
| `#ED98C1` | 13 |
| `#E464A2` | 12 |
| `#00CC00` | 12 |
| `#D9D9D9` | 8 |
| `#33B66E` | 7 |
| `#FFC000` | 6 |
| `#B2B2B2` | 5 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Arial` | 2261 |
| `Century Gothic` | 185 |
| `+mj-lt` | 160 |
| `Arial(Body)` | 23 |
| `+mn-lt` | 5 |
| `Times New Roman` | 1 |

**Theme fonts (major/minor):**

- Major: ['Century Gothic', 'Arial']
- Minor: ['Century Gothic', 'Arial']


### GSK (2 decks)

**Decks:**
- `GSK BLENREP MM PET Monthly Report February 2026_Migration Copy.pptx`
- `GSK Jemperli Zejula PET Full Report February '26 Syn Check.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#3AB51D` | 36 |
| `#668EDD` | 28 |
| `#6658A6` | 26 |
| `#5C103B` | 22 |
| `#7030A0` | 21 |
| `#D51900` | 18 |
| `#17B3AF` | 17 |
| `#1B3B7A` | 16 |
| `#E21860` | 14 |
| `#4F447D` | 13 |
| `#F36633` | 12 |
| `#F25E29` | 11 |
| `#706352` | 11 |
| `#A9A197` | 11 |
| `#00847C` | 11 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Arial` | 763 |
| `+mn-lt` | 11 |
| `DM Sans 14pt` | 11 |
| `+mj-lt` | 9 |
| `Century Gothic` | 4 |
| `Times New Roman` | 1 |

**Theme fonts (major/minor):**

- Major: ['Arial']
- Minor: ['Arial']


### LEO (2 decks)

**Decks:**
- `OG Adbry PET SC Deck.pptx`
- `V1_Adbry PET_W1'26 (Wave 22)_migrated deck.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#C014A3` | 36 |
| `#D5685F` | 30 |
| `#6F439A` | 28 |
| `#FED006` | 14 |
| `#EAAFE0` | 14 |
| `#009B77` | 12 |
| `#D561C1` | 8 |
| `#90117A` | 8 |
| `#946A2C` | 6 |
| `#2E8082` | 6 |
| `#1E41EB` | 6 |
| `#0047BB` | 5 |
| `#00609C` | 5 |
| `#4E7BA9` | 4 |
| `#00B050` | 4 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Arial` | 1436 |
| `Franklin Gothic Book` | 17 |

**Theme fonts (major/minor):**

- Major: ['Arial']
- Minor: ['Arial']


### Bone-HCP (1 deck)

**Decks:**
- `03.12.2026_Q1'26 Bone HCP PET Full Report_v1.0.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#FF9933` | 35 |
| `#009201` | 13 |
| `#00B0F0` | 13 |
| `#58193D` | 11 |
| `#A6A6A6` | 5 |
| `#0063C3` | 4 |
| `#4BACC6` | 4 |
| `#E46C0A` | 4 |
| `#425563` | 4 |
| `#FFC285` | 4 |
| `#604A7B` | 3 |
| `#2C53A2` | 3 |
| `#F79646` | 2 |
| `#4F6228` | 2 |
| `#93CDDD` | 2 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Century Gothic` | 1663 |
| `+mn-lt` | 33 |
| `Times New Roman` | 6 |
| `Calibri` | 2 |
| `Arial` | 1 |

**Theme fonts (major/minor):**

- Major: ['Arial']
- Minor: ['Arial']


### Otsuka (1 deck)

**Decks:**
- `[ZoomRx] Abilify LAI PET - Full Report.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#FFC000` | 21 |
| `#C00000` | 15 |
| `#2D5CA2` | 15 |
| `#7F47AA` | 13 |
| `#7D44A9` | 11 |
| `#3B5998` | 11 |
| `#7E97CD` | 10 |
| `#B3C9EA` | 8 |
| `#0000FF` | 8 |
| `#D8DDE5` | 8 |
| `#1C1755` | 7 |
| `#FFFFFF` | 6 |
| `#DA0072` | 5 |
| `#7AB2B2` | 4 |
| `#F1E9E9` | 4 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Arial` | 1093 |
| `Johnson Text` | 13 |
| `Calibri (Body)` | 13 |
| `Tahoma` | 9 |
| `12` | 4 |
| `Gotham Book` | 1 |

**Theme fonts (major/minor):**

- Major: ['Calibri']
- Minor: ['Calibri']


### Alexion (1 deck)

**Decks:**
- `[ZoomRx] Alexion US PNH Q325 PET - v1.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#0E8779` | 29 |
| `#7030A0` | 28 |
| `#E35105` | 22 |
| `#799A01` | 21 |
| `#FFA300` | 21 |
| `#FF8181` | 10 |
| `#CCCCCC` | 10 |
| `#84CA68` | 8 |
| `#009886` | 6 |
| `#D34D2F` | 6 |
| `#008DAB` | 6 |
| `#F8CF36` | 6 |
| `#003B75` | 6 |
| `#E47E19` | 5 |
| `#B2A4B6` | 4 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Arial` | 2880 |
| `Arial(Body)` | 70 |
| `Montserrat` | 14 |
| `Montserrat Regular` | 4 |

**Theme fonts (major/minor):**

- Major: ['Arial Black']
- Minor: ['Arial']


### BL (1 deck)

**Decks:**
- `[ZoomRx] B+L DED PET Jul-Sep'25_Final Report v3.0.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#00A9EB` | 38 |
| `#400286` | 37 |
| `#A8109D` | 29 |
| `#00B050` | 29 |
| `#A6A6A6` | 25 |
| `#FF1119` | 15 |
| `#413A5F` | 10 |
| `#000F9F` | 9 |
| `#000000` | 8 |
| `#7F7F7F` | 8 |
| `#FFC000` | 7 |
| `#F78626` | 7 |
| `#8822FC` | 7 |
| `#BED32E` | 6 |
| `#04974F` | 6 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `+mn-lt` | 235 |
| `Century Gothic` | 212 |
| `Avenir Next LT Pro` | 19 |
| `Source Sans Pro` | 2 |
| `+mj-lt` | 1 |
| ` century gothic` | 1 |

**Theme fonts (major/minor):**

- Major: ['Avenir Next LT Pro']
- Minor: ['Century Gothic']


### Novartis (1 deck)

**Decks:**
- `[ZoomRx] NVS Rhapsido_PET_Wave 4 Full Report v1.0.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#018E86` | 14 |
| `#0070FE` | 11 |
| `#002068` | 11 |
| `#8F2DDE` | 8 |
| `#852065` | 7 |
| `#B56FA3` | 7 |
| `#61A8FF` | 6 |
| `#50E2D0` | 6 |
| `#A7A8AA` | 6 |
| `#C89ABA` | 5 |
| `#E7D2E0` | 5 |
| `#E4EECA` | 4 |
| `#A6C850` | 4 |
| `#830F66` | 4 |
| `#DAAB8D` | 4 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Arial` | 3168 |
| `+mj-lt` | 880 |
| `Calibri (Body)` | 306 |
| `+mn-lt` | 144 |
| `Times New Roman` | 6 |
| `Tahoma` | 2 |

**Theme fonts (major/minor):**

- Major: ['Arial']
- Minor: ['Arial']


### CCA (1 deck)

**Decks:**
- `CCA PET Q2FY26 Personal Promotion Synapse check.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#F28E2B` | 5 |
| `#FFFFFF` | 4 |
| `#A6A6A6` | 4 |
| `#8FAADC` | 4 |
| `#BDD7EE` | 3 |
| `#AFED5D` | 3 |
| `#FFABD5` | 3 |
| `#BFBFBF` | 2 |
| `#FF6700` | 2 |
| `#FF8F43` | 2 |
| `#FFBC8F` | 2 |
| `#002060` | 2 |
| `#2F5597` | 2 |
| `#B4C7E7` | 2 |
| `#F28E41` | 1 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Century Gothic` | 366 |
| `Arial` | 317 |
| `+mn-lt` | 6 |
| `Wingdings` | 4 |
| `+mj-lt` | 1 |
| `-apple-system` | 1 |

**Theme fonts (major/minor):**

- Major: ['Century Gothic']
- Minor: ['Century Gothic']


### DSI (1 deck)

**Decks:**
- `DATROWAY EGFRm NSCLC Promotional Effectiveness Tracking (PET) Q1 '26 PP and NPP Final Report_v1.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#1E22AA` | 22 |
| `#643466` | 11 |
| `#FBCCB1` | 10 |
| `#A6A6A6` | 5 |
| `#B253DE` | 5 |
| `#353236` | 5 |
| `#CC0066` | 5 |
| `#75E6E9` | 3 |
| `#00AE9B` | 3 |
| `#7CD214` | 3 |
| `#A74109` | 3 |
| `#304B4C` | 3 |
| `#662D91` | 3 |
| `#353EFF` | 3 |
| `#3D94AA` | 3 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `+mj-lt` | 814 |
| `Arial` | 61 |
| `Arial (Body)` | 6 |
| `Manrope` | 2 |

**Theme fonts (major/minor):**

- Major: ['Arial']
- Minor: ['Arial']


### Ipsen (1 deck)

**Decks:**
- `Onivyde SFE W3 Report.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#54AC65` | 41 |
| `#C84874` | 13 |
| `#8EC899` | 11 |
| `#595959` | 9 |
| `#BBDEC2` | 9 |
| `#D3D3D3` | 7 |
| `#D2E9D6` | 6 |
| `#A6A6A6` | 6 |
| `#A9D5B2` | 6 |
| `#00B039` | 4 |
| `#387343` | 4 |
| `#FA9F1A` | 4 |
| `#BFBFBF` | 4 |
| `#C2EAFF` | 3 |
| `#DE505E` | 1 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `+mn-lt` | 59 |
| `+mj-lt` | 34 |
| `Calibri` | 2 |

**Theme fonts (major/minor):**

- Major: ['Rethink Sans']
- Minor: ['Rethink Sans']


### Apellis (1 deck)

**Decks:**
- `ZoomRx_Apellis EMPAVELI PET_Wave 2 PET_v1.pptx`

**Top series colors (from chart XML):**

| Hex | Count |
|---|---|
| `#FC3B6E` | 37 |
| `#E7E6E6` | 33 |
| `#30CFD0` | 32 |
| `#81E3E3` | 9 |
| `#5A4986` | 8 |
| `#FD7FA0` | 8 |
| `#FEB0C4` | 8 |
| `#FFC000` | 7 |
| `#163A6B` | 6 |
| `#5E6C85` | 5 |
| `#C6F3F2` | 4 |
| `#00B0F0` | 4 |
| `#9A9A9A` | 3 |
| `#0070C0` | 3 |
| `#666666` | 2 |

**Top fonts (from slides):**

| Font | Count |
|---|---|
| `Calibri` | 868 |
| `Calibri (Body)` | 209 |
| `+mn-lt` | 95 |
| `Wingdings` | 16 |
| `+mj-lt` | 14 |
| `Arial` | 11 |

**Theme fonts (major/minor):**

- Major: ['Calibri Light']
- Minor: ['Calibri']


---

## 4. Layout Coordinate Analysis

For recurring shape signatures (≥5 occurrences), coordinate statistics by shape type.

The `left_median` and `width_median` values are directly usable as `LAYOUTS{}` entries in `pptx_utils/layout.py`.

### Signature: `chart=0,table=0,text_box=0,picture=0,auto_shape=0` (94 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `placeholder` | 152 | 1.04" | 2.74" | 8.49" | 1.22" |
| `embedded_ole_object` | 1 | 0.0" | 0.0" | 0.0" | 0.0" |
| `none` | 16 | 0.92" | 1.21" | 10.59" | 5.57" |

### Signature: `chart=0,table=0,text_box=1,picture=0,auto_shape=0` (33 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `placeholder` | 69 | 0.54" | 4.48" | 7.29" | 0.62" |
| `text_box` | 33 | 6.45" | 3.37" | 5.47" | 0.4" |

### Signature: `chart=0,table=1,text_box=2,picture=0,auto_shape=0` (23 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `table` | 23 | 1.02" | 1.3" | 10.75" | 4.98" |
| `text_box` | 46 | 0.61" | 1.1" | 4.82" | 0.54" |
| `placeholder` | 25 | 1.2" | 6.79" | 6.43" | 0.61" |
| `line` | 4 | 0.92" | 3.06" | 0.09" | 0.29" |
| `freeform` | 6 | 1.15" | 3.75" | 0.28" | 0.27" |

### Signature: `chart=0,table=0,text_box=0,picture=0,auto_shape=1` (14 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `placeholder` | 25 | 0.4" | 1.37" | 10.53" | 1.56" |
| `auto_shape` | 14 | 0.64" | 1.83" | 0.37" | 5.02" |

### Signature: `chart=0,table=0,text_box=1,picture=1,auto_shape=1` (14 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `embedded_ole_object` | 1 | 0.0" | 0.0" | 0.0" | 0.0" |
| `placeholder` | 27 | 5.1" | 1.09" | 3.13" | 0.44" |
| `auto_shape` | 14 | 11.32" | 1.76" | 1.34" | 0.41" |
| `picture` | 14 | 2.88" | 1.76" | 7.57" | 5.21" |
| `text_box` | 14 | 10.1" | 0.0" | 3.23" | 0.36" |

### Signature: `chart=0,table=0,text_box=10,picture=0,auto_shape=11` (12 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `auto_shape` | 132 | 1.8" | 1.35" | 0.16" | 0.16" |
| `text_box` | 120 | 5.55" | 3.7" | 4.64" | 0.4" |
| `group` | 82 | 4.65" | 2.95" | 0.83" | 0.16" |
| `line` | 72 | 1.13" | 1.43" | 0.76" | 0.0" |

### Signature: `chart=0,table=0,text_box=1,picture=0,auto_shape=1` (11 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `auto_shape` | 11 | 0.67" | 2.29" | 5.81" | 0.57" |
| `placeholder` | 11 | 0.77" | 0.27" | 8.6" | 1.22" |
| `text_box` | 11 | 0.46" | 0.85" | 8.37" | 2.32" |
| `embedded_ole_object` | 1 | 4.22" | 3.34" | 1.0" | 0.84" |

### Signature: `chart=0,table=1,text_box=3,picture=0,auto_shape=0` (10 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `placeholder` | 7 | 12.34" | 7.26" | 0.59" | 0.24" |
| `text_box` | 30 | 2.41" | 0.58" | 2.57" | 0.4" |
| `table` | 10 | 1.22" | 1.81" | 10.46" | 4.45" |

### Signature: `chart=0,table=0,text_box=5,picture=0,auto_shape=11` (10 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `text_box` | 50 | 2.67" | 3.11" | 5.28" | 0.4" |
| `group` | 50 | 1.58" | 2.86" | 0.93" | 0.91" |
| `auto_shape` | 110 | 0.88" | 1.37" | 0.97" | 0.75" |

### Signature: `chart=0,table=1,text_box=0,picture=0,auto_shape=1` (8 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `placeholder` | 15 | 0.4" | 0.53" | 10.53" | 0.47" |
| `auto_shape` | 8 | 0.64" | 1.83" | 0.37" | 3.24" |
| `table` | 8 | 2.08" | 2.84" | 6.41" | 3.38" |

### Signature: `chart=0,table=0,text_box=15,picture=0,auto_shape=8` (8 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `placeholder` | 8 | 0.49" | 0.53" | 12.35" | 0.93" |
| `auto_shape` | 64 | 0.4" | 1.95" | 0.55" | 0.55" |
| `text_box` | 120 | 1.37" | 3.23" | 3.19" | 0.27" |
| `line` | 33 | 1.12" | 2.17" | 5.03" | 0.0" |
| `group` | 84 | 0.38" | 1.88" | 5.81" | 0.65" |

### Signature: `chart=0,table=1,text_box=1,picture=0,auto_shape=0` (8 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `placeholder` | 8 | 0.4" | 0.22" | 12.53" | 0.64" |
| `table` | 8 | 0.98" | 1.18" | 11.42" | 4.84" |
| `text_box` | 8 | 5.77" | 0.38" | 2.27" | 0.18" |

### Signature: `chart=0,table=0,text_box=2,picture=0,auto_shape=1` (8 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `text_box` | 16 | 5.88" | 0.41" | 7.33" | 2.4" |
| `placeholder` | 8 | 0.01" | 7.15" | 0.58" | 0.22" |
| `auto_shape` | 8 | 0.24" | 2.8" | 8.81" | 0.6" |

### Signature: `chart=0,table=1,text_box=4,picture=0,auto_shape=0` (7 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `table` | 7 | 0.88" | 1.27" | 11.19" | 5.11" |
| `text_box` | 28 | 4.05" | 0.35" | 4.54" | 0.46" |
| `placeholder` | 4 | 12.34" | 7.1" | 0.59" | 0.27" |

### Signature: `chart=0,table=0,text_box=2,picture=0,auto_shape=3` (7 slides)

**Coordinate stats by shape type:**

| Shape Type | N | Left (med) | Top (med) | Width (med) | Height (med) |
|---|---|---|---|---|---|
| `text_box` | 14 | 1.51" | 0.58" | 8.18" | 1.45" |
| `auto_shape` | 21 | 2.2" | 1.65" | 0.16" | 0.15" |

---

## 5. Headline Patterns

**3569** headline-like text boxes detected (top of slide, wide, ≥8 chars).

**Char count:** min 8, median 72, max 2303

**Width (inches):** median 11.94, min 4.25, max 14.55

### Font sizes (top 10)

| Size (pt) | Count |
|---|---|
| 16 | 876 |
| 20 | 678 |
| 18 | 678 |
| 14 | 340 |
| 12 | 155 |
| 24 | 149 |
| 28 | 133 |
| 8 | 61 |
| 13 | 58 |
| 10 | 53 |

### Font colors (top 15)

| Hex | Count |
|---|---|
| `#000000` | 455 |
| `#0063C3` | 134 |
| `#001E60` | 126 |
| `#595959` | 110 |
| `#00A3DC` | 86 |
| `#002B5C` | 72 |
| `#E44405` | 59 |
| `#FF0000` | 53 |
| `#636466` | 46 |
| `#3E403F` | 46 |
| `#00264C` | 45 |
| `#404040` | 39 |
| `#4060AF` | 36 |
| `#242269` | 32 |
| `#3F4444` | 31 |

### Top positions (top of slide, inches)

| Top (in) | Count |
|---|---|
| -0.2 | 7 |
| -0.1 | 17 |
| 0.0 | 117 |
| 0.1 | 415 |
| 0.2 | 406 |
| 0.3 | 543 |
| 0.4 | 100 |
| 0.5 | 231 |
| 0.6 | 85 |
| 0.7 | 158 |
| 0.8 | 171 |
| 0.9 | 150 |
| 1.0 | 266 |
| 1.1 | 181 |
| 1.2 | 241 |

---

## 6. Table Patterns

**4950** tables across all decks.

**Width (inches):** median 1.95, min 0.23, max 13.26
**Height (inches):** median 2.99, min 0.15, max 22.35

### Top dimensions (rows × cols)

| Dimensions | Count |
|---|---|
| `1x1` | 1094 |
| `11x1` | 200 |
| `5x1` | 176 |
| `3x1` | 170 |
| `12x1` | 170 |
| `10x1` | 165 |
| `4x1` | 154 |
| `1x3` | 151 |
| `6x1` | 147 |
| `2x1` | 142 |
| `8x1` | 140 |
| `1x2` | 114 |
| `7x1` | 103 |
| `1x4` | 90 |
| `13x3` | 89 |

### Row count distribution (top 10)

| Rows | Count |
|---|---|
| 1 | 1574 |
| 11 | 355 |
| 2 | 322 |
| 10 | 317 |
| 12 | 314 |
| 4 | 258 |
| 5 | 254 |
| 3 | 252 |
| 8 | 240 |
| 6 | 226 |

### Col count distribution (top 10)

| Cols | Count |
|---|---|
| 1 | 2937 |
| 3 | 775 |
| 2 | 599 |
| 4 | 307 |
| 5 | 156 |
| 6 | 68 |
| 7 | 41 |
| 12 | 16 |
| 8 | 14 |
| 13 | 8 |
