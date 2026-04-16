# BRAND{} Proposal by Client

Derived from real deck analysis. For each client, proposes a `BRAND{}` entry structure using observed series colors and fonts.

Use as input for `pptx_utils/brand.py`.

## AZN

```python
"AZN": {
    "primary":    RGBColor(0x00, 0xB0, 0x50),  # most-used series color
    "secondary":  RGBColor(0xFF, 0x88, 0x13),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Arial",  # from theme
    "font_body":    "Arial",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#00B050`
- `#FF8813`
- `#94D448`
- `#E15759`
- `#250E62`
- `#33A4FF`
- `#59A14F`
- `#9467BD`
- `#1D7B87`
- `#91C3D5`

## JJ

```python
"JJ": {
    "primary":    RGBColor(0x7F, 0xB1, 0xE1),  # most-used series color
    "secondary":  RGBColor(0x00, 0x63, 0xC3),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Johnson Display",  # from theme
    "font_body":    "Johnson Text",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#7FB1E1`
- `#0063C3`
- `#6FC6C1`
- `#37CFCA`
- `#BD05ED`
- `#228D8A`
- `#A0D0FF`
- `#FF6A5A`
- `#2A41C1`
- `#CBF3F1`

## Pfizer

```python
"Pfizer": {
    "primary":    RGBColor(0x00, 0x00, 0xC9),  # most-used series color
    "secondary":  RGBColor(0x43, 0x96, 0x4A),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Arial",  # from theme
    "font_body":    "Arial",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#0000C9`
- `#43964A`
- `#E356DC`
- `#F49C34`
- `#BA2A4C`
- `#C00000`
- `#6D5A7A`
- `#4060AF`
- `#7F7F7F`
- `#081E3D`

## Regeneron

```python
"Regeneron": {
    "primary":    RGBColor(0x21, 0x94, 0x91),  # most-used series color
    "secondary":  RGBColor(0xF7, 0x96, 0x46),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Trade Gothic LT Std",  # from theme
    "font_body":    "Arial",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#219491`
- `#F79646`
- `#E46C0A`
- `#DC0077`
- `#8EB4E3`
- `#00745A`
- `#7030A0`
- `#004F6F`
- `#B370A4`
- `#7ABFBD`

## Amgen

```python
"Amgen": {
    "primary":    RGBColor(0x1F, 0x49, 0x7D),  # most-used series color
    "secondary":  RGBColor(0x81, 0x3F, 0x97),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Century Gothic",  # from theme
    "font_body":    "Century Gothic",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#1F497D`
- `#813F97`
- `#007D6D`
- `#6F8DB4`
- `#003C71`
- `#DA205F`
- `#00A3DF`
- `#4C35DE`
- `#0063C3`
- `#00B0F0`

## Unknown

```python
"Unknown": {
    "primary":    RGBColor(0xE7, 0x00, 0x4C),  # most-used series color
    "secondary":  RGBColor(0xC7, 0xA0, 0x13),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Century Gothic",  # from theme
    "font_body":    "Arial",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#E7004C`
- `#C7A013`
- `#6BC9EF`
- `#006838`
- `#FF6969`
- `#00607C`
- `#4E79A7`
- `#00A44A`
- `#ED98C1`
- `#E464A2`

## GSK

```python
"GSK": {
    "primary":    RGBColor(0x3A, 0xB5, 0x1D),  # most-used series color
    "secondary":  RGBColor(0x66, 0x8E, 0xDD),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Arial",  # from theme
    "font_body":    "Arial",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#3AB51D`
- `#668EDD`
- `#6658A6`
- `#5C103B`
- `#7030A0`
- `#D51900`
- `#17B3AF`
- `#1B3B7A`
- `#E21860`
- `#4F447D`

## LEO

```python
"LEO": {
    "primary":    RGBColor(0xC0, 0x14, 0xA3),  # most-used series color
    "secondary":  RGBColor(0xD5, 0x68, 0x5F),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Arial",  # from theme
    "font_body":    "Arial",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#C014A3`
- `#D5685F`
- `#6F439A`
- `#FED006`
- `#EAAFE0`
- `#009B77`
- `#D561C1`
- `#90117A`
- `#946A2C`
- `#2E8082`

## Bone-HCP

```python
"Bone-HCP": {
    "primary":    RGBColor(0xFF, 0x99, 0x33),  # most-used series color
    "secondary":  RGBColor(0x00, 0x92, 0x01),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Arial",  # from theme
    "font_body":    "Century Gothic",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#FF9933`
- `#009201`
- `#00B0F0`
- `#58193D`
- `#A6A6A6`
- `#0063C3`
- `#4BACC6`
- `#E46C0A`
- `#425563`
- `#FFC285`

## Otsuka

```python
"Otsuka": {
    "primary":    RGBColor(0xFF, 0xC0, 0x00),  # most-used series color
    "secondary":  RGBColor(0xC0, 0x00, 0x00),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Calibri",  # from theme
    "font_body":    "Arial",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#FFC000`
- `#C00000`
- `#2D5CA2`
- `#7F47AA`
- `#7D44A9`
- `#3B5998`
- `#7E97CD`
- `#B3C9EA`
- `#0000FF`
- `#D8DDE5`

## Alexion

```python
"Alexion": {
    "primary":    RGBColor(0x0E, 0x87, 0x79),  # most-used series color
    "secondary":  RGBColor(0x70, 0x30, 0xA0),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Arial Black",  # from theme
    "font_body":    "Arial",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#0E8779`
- `#7030A0`
- `#E35105`
- `#799A01`
- `#FFA300`
- `#FF8181`
- `#CCCCCC`
- `#84CA68`
- `#009886`
- `#D34D2F`

## BL

```python
"BL": {
    "primary":    RGBColor(0x00, 0xA9, 0xEB),  # most-used series color
    "secondary":  RGBColor(0x40, 0x02, 0x86),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Avenir Next LT Pro",  # from theme
    "font_body":    "Century Gothic",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#00A9EB`
- `#400286`
- `#A8109D`
- `#00B050`
- `#A6A6A6`
- `#FF1119`
- `#413A5F`
- `#000F9F`
- `#000000`
- `#7F7F7F`

## Novartis

```python
"Novartis": {
    "primary":    RGBColor(0x01, 0x8E, 0x86),  # most-used series color
    "secondary":  RGBColor(0x00, 0x70, 0xFE),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Arial",  # from theme
    "font_body":    "Arial",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#018E86`
- `#0070FE`
- `#002068`
- `#8F2DDE`
- `#852065`
- `#B56FA3`
- `#61A8FF`
- `#50E2D0`
- `#A7A8AA`
- `#C89ABA`

## CCA

```python
"CCA": {
    "primary":    RGBColor(0xF2, 0x8E, 0x2B),  # most-used series color
    "secondary":  RGBColor(0xFF, 0xFF, 0xFF),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Century Gothic",  # from theme
    "font_body":    "Century Gothic",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#F28E2B`
- `#FFFFFF`
- `#A6A6A6`
- `#8FAADC`
- `#BDD7EE`
- `#AFED5D`
- `#FFABD5`
- `#BFBFBF`
- `#FF6700`
- `#FF8F43`

## DSI

```python
"DSI": {
    "primary":    RGBColor(0x1E, 0x22, 0xAA),  # most-used series color
    "secondary":  RGBColor(0x64, 0x34, 0x66),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Arial",  # from theme
    "font_body":    "Arial",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#1E22AA`
- `#643466`
- `#FBCCB1`
- `#A6A6A6`
- `#B253DE`
- `#353236`
- `#CC0066`
- `#75E6E9`
- `#00AE9B`
- `#7CD214`

## Ipsen

```python
"Ipsen": {
    "primary":    RGBColor(0x54, 0xAC, 0x65),  # most-used series color
    "secondary":  RGBColor(0xC8, 0x48, 0x74),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Rethink Sans",  # from theme
    "font_body":    "Calibri",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#54AC65`
- `#C84874`
- `#8EC899`
- `#595959`
- `#BBDEC2`
- `#D3D3D3`
- `#D2E9D6`
- `#A6A6A6`
- `#A9D5B2`
- `#00B039`

## Apellis

```python
"Apellis": {
    "primary":    RGBColor(0xFC, 0x3B, 0x6E),  # most-used series color
    "secondary":  RGBColor(0xE7, 0xE6, 0xE6),  # 2nd most-used series color
    "positive":   RGBColor(0x00, 0xB0, 0x50),  # standard green for positive delta
    "negative":   RGBColor(0xFF, 0x00, 0x00),  # standard red for negative delta
    "font_heading": "Calibri Light",  # from theme
    "font_body":    "Calibri",  # most-used body font
},
```

**Observed color palette (top 10):**

- `#FC3B6E`
- `#E7E6E6`
- `#30CFD0`
- `#81E3E3`
- `#5A4986`
- `#FD7FA0`
- `#FEB0C4`
- `#FFC000`
- `#163A6B`
- `#5E6C85`
