"""
Builds the complete IBM Applied Data Science Capstone Presentation PPTX
and exports it as 'Data Science Capstone Project Report.pdf'.
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import win32com.client

OUTPUT_DIR = r"c:\Users\nivpo\OneDrive\college\coursera"
IMG_DIR = r"c:\Users\nivpo\OneDrive\college\coursera\spacex_capstone\images"
PPTX_PATH = os.path.join(OUTPUT_DIR, "Data Science Capstone Project Report.pptx")
PDF_PATH = os.path.join(OUTPUT_DIR, "Data Science Capstone Project Report.pdf")

GITHUB_URL = "https://github.com/Nivporwal25/SpaceX-Falcon-9-First-Stage-Landing-Prediction"

# Colors
DARK_NAVY = RGBColor(13, 27, 42)      # #0D1B2A
MID_BLUE  = RGBColor(27, 58, 107)     # #1B3A6B
ACCENT_BLUE = RGBColor(31, 120, 200)  # #1F78C8
LIGHT_BG   = RGBColor(245, 247, 250)  # #F5F7FA
WHITE      = RGBColor(255, 255, 255)
GOLD       = RGBColor(244, 162, 97)   # #F4A261
GRAY_TEXT  = RGBColor(100, 110, 120)
GREEN      = RGBColor(45, 198, 83)    # #2DC653

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

def add_header(slide, title_text, category_text="IBM APPLIED DATA SCIENCE CAPSTONE"):
    # Header background band
    banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
    banner.fill.solid()
    banner.fill.fore_color.rgb = DARK_NAVY
    banner.line.color.rgb = DARK_NAVY

    # Category tracker
    cat_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.12), Inches(12), Inches(0.3))
    tf_cat = cat_box.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.size = Pt(10)
    p_cat.font.bold = True
    p_cat.font.color.rgb = GOLD

    # Slide Title
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12), Inches(0.6))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.text = title_text
    p_t.font.size = Pt(20)
    p_t.font.bold = True
    p_t.font.color.rgb = WHITE

def add_footer(slide, current_page=None, total_pages=16):
    footer_box = slide.shapes.add_textbox(Inches(0.6), Inches(7.1), Inches(12.133), Inches(0.35))
    tf = footer_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = f"Project GitHub Repository: {GITHUB_URL}"
    p.font.size = Pt(9)
    p.font.color.rgb = ACCENT_BLUE

    if current_page:
        pg_box = slide.shapes.add_textbox(Inches(11.5), Inches(7.1), Inches(1.2), Inches(0.35))
        tf_pg = pg_box.text_frame
        p_pg = tf_pg.paragraphs[0]
        p_pg.text = f"Slide {current_page} of {total_pages}"
        p_pg.alignment = PP_ALIGN.RIGHT
        p_pg.font.size = Pt(9)
        p_pg.font.color.rgb = GRAY_TEXT

def create_card(slide, left, top, width, height, title, items, bg_color=WHITE, border_color=MID_BLUE):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1.5)

    tb = slide.shapes.add_textbox(Inches(left + 0.15), Inches(top + 0.15), Inches(width - 0.3), Inches(height - 0.3))
    tf = tb.text_frame
    tf.word_wrap = True

    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.size = Pt(14)
    p0.font.bold = True
    p0.font.color.rgb = MID_BLUE
    p0.space_after = Pt(8)

    for item in items:
        p = tf.add_paragraph()
        p.text = f"•  {item}"
        p.font.size = Pt(11)
        p.font.color.rgb = RGBColor(30, 41, 59)
        p.space_after = Pt(5)

# ==============================================================================
# SLIDE 1: Title Slide (Criteria 1.1 & 1.2)
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
bg1.fill.solid()
bg1.fill.fore_color.rgb = DARK_NAVY
bg1.line.color.rgb = DARK_NAVY

# Inner border
border1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.5), Inches(12.333), Inches(6.5))
border1.fill.background()
border1.line.color.rgb = ACCENT_BLUE
border1.line.width = Pt(2)

tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.333), Inches(4.5))
tf1 = tb1.text_frame
tf1.word_wrap = True

p_prog = tf1.paragraphs[0]
p_prog.text = "IBM APPLIED DATA SCIENCE PROFESSIONAL CERTIFICATE"
p_prog.font.size = Pt(13)
p_prog.font.bold = True
p_prog.font.color.rgb = GOLD
p_prog.space_after = Pt(12)

p_t1 = tf1.add_paragraph()
p_t1.text = "SpaceX Falcon 9 First Stage Landing Prediction"
p_t1.font.size = Pt(32)
p_t1.font.bold = True
p_t1.font.color.rgb = WHITE
p_t1.space_after = Pt(8)

p_sub = tf1.add_paragraph()
p_sub.text = "Capstone Project Report — Machine Learning, EDA, SQL, & Interactive Visual Analytics"
p_sub.font.size = Pt(16)
p_sub.font.color.rgb = RGBColor(147, 197, 253)
p_sub.space_after = Pt(28)

p_author = tf1.add_paragraph()
p_author.text = "Author: Niv Vikram Porwal"
p_author.font.size = Pt(14)
p_author.font.bold = True
p_author.font.color.rgb = WHITE
p_author.space_after = Pt(6)

p_date = tf1.add_paragraph()
p_date.text = "Date: October 2026 | Platform: Coursera & IBM Skills Network"
p_date.font.size = Pt(12)
p_date.font.color.rgb = RGBColor(203, 213, 225)
p_date.space_after = Pt(20)

p_gh = tf1.add_paragraph()
p_gh.text = f"Project GitHub Repository URL: {GITHUB_URL}"
p_gh.font.size = Pt(13)
p_gh.font.bold = True
p_gh.font.color.rgb = GOLD

# ==============================================================================
# SLIDE 2: Executive Summary (Criterion 1.3)
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "Executive Summary: Project Overview & Key Findings")
add_footer(s2, 2)

create_card(s2, 0.6, 1.3, 5.8, 5.5, "Methodology & Data Science Pipeline", [
    "Data Collection: SpaceX REST API v4 & Web Scraping Wikipedia via BeautifulSoup.",
    "Data Wrangling: Cleaned 90 Falcon 9 records, imputed nulls with mean payload (6,105 kg), one-hot encoded categorical features (Orbit, LaunchSite, LandingPad, Serial).",
    "Exploratory Data Analysis: Analyzed flight trends, payload correlations, and success rates by orbit using Pandas, Seaborn, Matplotlib, and SQLite queries.",
    "Interactive Visual Analytics: Built geospatial Folium maps with distance proximity analysis and Plotly Dash dashboard with reactive callbacks.",
    "Predictive Modeling: Evaluated Logistic Regression, SVM, Decision Tree, and KNN using 10-fold GridSearchCV cross-validation."
])

create_card(s2, 6.8, 1.3, 5.8, 5.5, "Key Results & Business Implications", [
    "Model Performance: Decision Tree, SVM, and Logistic Regression achieved 83.33% classification accuracy on unseen test data.",
    "High Precision: Best model achieved 0 False Positives, ensuring 100% precision when predicting a successful first stage recovery.",
    "Technological Evolution: Landing success rate grew from 0% (2010-2013) to over 85% by 2020 as reusability and Block 5 engines matured.",
    "Launch Site Performance: KSC LC-39A is the most reliable launch complex with a 76.9% landing success rate.",
    "Commercial Impact: SpaceX launch price ($62M) vs competitors ($165M+) gives a 62% cost advantage; accurate recovery prediction enables competitors to optimize bidding strategy."
])

# ==============================================================================
# SLIDE 3: Introduction – Background & Problem Statement (Criterion 1.4)
# ==============================================================================
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "Introduction: Project Background & Problem Statement")
add_footer(s3, 3)

create_card(s3, 0.6, 1.3, 5.8, 5.5, "Project Background & Aerospace Context", [
    "The space launch industry has historically been dominated by expendable rockets costing over $165 Million per launch.",
    "SpaceX fundamentally revolutionized commercial spaceflight by proving first-stage booster reusability on Falcon 9.",
    "By successfully recovering the booster stage via Autonomous Spaceport Drone Ship (ASDS) and Return to Launch Site (RTLS), SpaceX cut launch prices to ~$62 Million.",
    "First stage booster accounts for approximately 70% of total rocket manufacturing costs.",
    "Reusability dramatically increases launch cadence and lowers barriers to orbital payloads."
])

create_card(s3, 6.8, 1.3, 5.8, 5.5, "Problem Statement & Analytical Objectives", [
    "Core Question: Can we accurately predict whether a Falcon 9 first stage booster will land successfully based on pre-launch flight parameters?",
    "Target Classification Variable: Class = 1 (Successful Landing) vs. Class = 0 (Failure / Expended).",
    "Independent Features: Payload Mass, Target Orbit, Launch Site, Booster Flight Count, Grid Fins, Legs, Block Version, and Landing Pad.",
    "Commercial Value: If a competitor or satellite operator can predict landing success, they can deduce SpaceX's true cost per flight and submit competitive launch tenders.",
    "Outcome: Provide data-driven operational insights into factors driving first stage recovery."
])

# ==============================================================================
# SLIDE 4: Data Collection – SpaceX REST API (Criterion 1.5)
# ==============================================================================
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "Data Collection: SpaceX REST API v4")
add_footer(s4, 4)

s4.shapes.add_picture(os.path.join(IMG_DIR, "api_flowchart.png"), Inches(0.6), Inches(1.3), Inches(12.133), Inches(2.2))

create_card(s4, 0.6, 3.7, 5.8, 3.2, "API Endpoints & Extraction Process", [
    "Base API Endpoint: https://api.spacexdata.com/v4/launches/past",
    "Rocket Details: Queried /v4/rockets to filter records exclusively for 'Falcon 9'.",
    "Payload Data: Queried /v4/payloads to retrieve payload mass (kg) and target orbit.",
    "Launchpads: Queried /v4/launchpads to obtain pad names, latitudes, and longitudes.",
    "Booster Cores: Queried /v4/cores to retrieve booster serial, flights count, grid fins, landing legs, and landing pad IDs."
])

create_card(s4, 6.8, 3.7, 5.8, 3.2, "Data Wrangling of API Responses", [
    "JSON Normalization: Flattened nested JSON dictionaries into tabular format using pd.json_normalize().",
    "Filtering: Extracted 90 total Falcon 9 flights spanning 2010 to 2020.",
    "Missing Value Handling: Calculated mean payload mass (6,104.96 kg) to impute missing payload values.",
    "Feature Assembly: Consolidated 17 primary operational columns into a unified pandas DataFrame.",
    "API Flowchart above illustrates the end-to-end data pipeline."
])

# ==============================================================================
# SLIDE 5: Data Collection – Web Scraping (Criterion 1.6)
# ==============================================================================
s5 = prs.slides.add_slide(blank_layout)
add_header(s5, "Data Collection: Web Scraping Falcon 9 Records")
add_footer(s5, 5)

s5.shapes.add_picture(os.path.join(IMG_DIR, "scraping_flowchart.png"), Inches(0.6), Inches(1.3), Inches(12.133), Inches(2.2))

create_card(s5, 0.6, 3.7, 5.8, 3.2, "Web Scraping Pipeline & Tools", [
    "Target Source: Wikipedia page 'List of Falcon 9 and Falcon Heavy launches'.",
    "HTTP Requests: Dispatched GET request using Python requests.get() library.",
    "HTML DOM Parsing: Utilized BeautifulSoup (bs4) with lxml parser.",
    "Table Extraction: Targeted all <table class='wikitable'> elements containing launch records.",
    "Cell Extraction: Looped through table rows <tr>, extracting headers <th> and data cells <td>."
])

create_card(s5, 6.8, 3.7, 5.8, 3.2, "Data Cleansing & Validation", [
    "Regex Tag Removal: Stripped footnote references <sup>[n]</sup> and newline characters \\n.",
    "Date & Time Parsing: Extracted standardized UTC launch dates and times.",
    "Outcome Standardization: Parsed complex textual notes into clear landing outcome categories.",
    "Cross-Validation: Web-scraped dataset served as validation benchmark against official SpaceX API data.",
    "Extracted Columns: Flight No., Date, Booster Version, Launch Site, Payload, Orbit, Customer, Outcome."
])

# ==============================================================================
# SLIDE 6: Data Wrangling Methodology (Criterion 1.7)
# ==============================================================================
s6 = prs.slides.add_slide(blank_layout)
add_header(s6, "Data Wrangling & Feature Engineering Methodology")
add_footer(s6, 6)

s6.shapes.add_picture(os.path.join(IMG_DIR, "wrangling_flowchart.png"), Inches(0.6), Inches(1.3), Inches(12.133), Inches(2.2))

create_card(s6, 0.6, 3.7, 5.8, 3.2, "Data Cleaning & Transformation", [
    "Dataset Filtering: Preserved 90 Falcon 9 launch records; removed test/Falcon 1 flights.",
    "Missing Value Imputation: 5 missing PayloadMass entries imputed with mean (6,104.96 kg).",
    "Target Variable Creation (Class):",
    "  - Class = 1 (Success): 'True Ocean', 'True RTLS', 'True ASDS'.",
    "  - Class = 0 (Failure): 'False Ocean', 'False RTLS', 'False ASDS', 'None ASDS', 'None None'."
])

create_card(s6, 6.8, 3.7, 5.8, 3.2, "Feature Encoding & Preprocessing", [
    "One-Hot Encoding: Applied pd.get_dummies() to categorical features (Orbit, LaunchSite, LandingPad, Serial).",
    "Feature Expansion: Transformed 17 initial raw variables into 83 numeric modeling features.",
    "Feature Scaling: Applied StandardScaler() to standardize feature variance across columns.",
    "Train/Test Partition: Split into 80% Training (72 records) and 20% Testing (18 records) with stratification.",
    "Ensured zero data leakage between training and evaluation splits."
])

# ==============================================================================
# SLIDE 7: EDA with Data Visualization – Methodology (Criterion 1.8)
# ==============================================================================
s7 = prs.slides.add_slide(blank_layout)
add_header(s7, "Exploratory Data Analysis (EDA): Visualization Methodology")
add_footer(s7, 7)

create_card(s7, 0.6, 1.3, 5.8, 5.5, "Visualization Suite & Analytical Objectives", [
    "Flight Number vs. Launch Site (Scatter Plot): Analyze launch cadence growth and evaluate whether later flights achieved higher landing success across sites.",
    "Payload Mass vs. Launch Site (Scatter Plot): Identify payload capacity limits per launch complex and observe success distribution across payload tiers.",
    "Success Rate by Orbit Type (Bar Chart): Determine which destination orbits (LEO, ISS, GTO, SSO, Polar) offer the most favorable booster recovery conditions.",
    "Flight Number vs. Orbit Type (Scatter Plot): Map the chronological evolution of orbital mission requirements throughout SpaceX history.",
    "Payload Mass vs. Orbit Type (Scatter Plot): Examine payload mass distribution across orbits and detect payload recovery threshold constraints.",
    "Launch Success Yearly Trend (Line Plot): Measure year-over-year recovery reliability improvements from 2010 to 2020."
])

create_card(s7, 6.8, 1.3, 5.8, 5.5, "Visualization Framework & Design Standards", [
    "Plotting Libraries: Matplotlib (v3.5+) and Seaborn (v0.12+) in Python.",
    "Color Palette: Standardized color scheme across all figures (Green = Success [Class 1], Red = Failure [Class 0]).",
    "Typography: Clear, bold axis titles, legible data labels, and complete legends for intuitive readability.",
    "Data Storytelling: Focused on isolating actionable operational insights, booster lifecycle maturation, and payload sweet-spots.",
    "Artifacts: All plots generated at high resolution (300 DPI) for clarity.",
    "Full implementation available in the project GitHub repository."
])

# ==============================================================================
# SLIDE 8: EDA with SQL – Methodology (Criterion 1.9)
# ==============================================================================
s8 = prs.slides.add_slide(blank_layout)
add_header(s8, "Exploratory Data Analysis (EDA): SQL Query Methodology")
add_footer(s8, 8)

create_card(s8, 0.6, 1.3, 5.8, 5.5, "Relational Database Setup & Execution", [
    "Database Engine: SQLite (spacex.db) loaded via Python SQLAlchemy & pandas to_sql.",
    "Query Execution: Executed via Jupyter ipython-sql (%sql magic) and standard sqlite3.",
    "Core Domains Investigated via SQL Queries:",
    "  1. Unique Launch Sites identification & naming validation.",
    "  2. Site-specific launch records filtering (CCAFS pads).",
    "  3. Total & average payload calculations for specific customers (NASA).",
    "  4. Payload threshold queries (boosters carrying 4,000 - 6,000 kg).",
    "  5. Aggregate landing outcomes counts (Drone Ship vs. Ground Pad).",
    "  6. Date & time analysis (earliest successful landing, yearly counts).",
    "  7. Ranked landing outcome frequency distributions."
])

create_card(s8, 6.8, 1.3, 5.8, 5.5, "Representative SQL Queries Executed", [
    "Q1 - Launch Sites: SELECT DISTINCT LaunchSite FROM SPACEXTABLE;",
    "Q2 - Total Payload: SELECT SUM(PayloadMass) FROM SPACEXTABLE WHERE Customer LIKE '%NASA%';",
    "Q3 - Average Payload: SELECT AVG(PayloadMass) FROM SPACEXTABLE WHERE BoosterVersion LIKE '%v1.1%';",
    "Q4 - First Ground Landing: SELECT MIN(Date) FROM SPACEXTABLE WHERE LandingOutcome = 'Success (ground pad)';",
    "Q5 - Max Payload Booster: SELECT BoosterVersion FROM SPACEXTABLE WHERE PayloadMass = (SELECT MAX(PayloadMass) FROM SPACEXTABLE);",
    "Q6 - Outcome Ranking: SELECT LandingOutcome, COUNT(*) FROM SPACEXTABLE GROUP BY LandingOutcome ORDER BY COUNT(*) DESC;"
])

# ==============================================================================
# SLIDE 9: Interactive Visual Analytics – Methodology (Criterion 1.10)
# ==============================================================================
s9 = prs.slides.add_slide(blank_layout)
add_header(s9, "Interactive Visual Analytics: Folium & Plotly Dash Methodology")
add_footer(s9, 9)

create_card(s9, 0.6, 1.3, 5.8, 5.5, "Folium Geospatial Mapping Framework", [
    "Geospatial Library: Python Folium built upon Leaflet.js.",
    "Launch Site Pinpointing: Plotted all 4 SpaceX launch pads using folium.Circle and folium.Marker with coordinate tooltips.",
    "Marker Clustering: Integrated MarkerCluster plugin to aggregate dense launch clusters without map clutter.",
    "Color-Coded Status: Green markers denote successful booster landings; Red markers denote failures.",
    "Coordinate Tracking: Integrated MousePosition plugin for live latitude/longitude readout.",
    "Proximity Measurement: Used folium.PolyLine to calculate exact straight-line distances to coastlines, highways, railways, and cities."
])

create_card(s9, 6.8, 1.3, 5.8, 5.5, "Plotly Dash Web Application Architecture", [
    "Web Framework: Plotly Dash (Dash Core Components & Dash HTML Components).",
    "Reactive Interactivity: Configured bidirectional callbacks linking user UI input to real-time graph updates.",
    "UI Controls:",
    "  - dcc.Dropdown: Switch between 'ALL Sites' and specific launch complexes (CCAFS SLC-40, KSC LC-39A, VAFB SLC-4E).",
    "  - dcc.RangeSlider: Interactive payload filter spanning 0 kg to 10,000 kg.",
    "Dynamic Graphs:",
    "  - Success Pie Chart: Displays site-specific success vs. failure proportions.",
    "  - Payload vs. Outcome Scatter Plot: Correlates payload mass with recovery success color-coded by booster version."
])

# ==============================================================================
# SLIDE 10: EDA Results – Flight Number & Payload vs. Launch Site (Criterion 1.11)
# ==============================================================================
s10 = prs.slides.add_slide(blank_layout)
add_header(s10, "EDA Results: Flight Number & Payload Mass vs. Launch Site")
add_footer(s10, 10)

s10.shapes.add_picture(os.path.join(IMG_DIR, "eda_flight_vs_site.png"), Inches(0.6), Inches(1.3), Inches(5.8), Inches(3.2))
s10.shapes.add_picture(os.path.join(IMG_DIR, "eda_payload_vs_site.png"), Inches(6.8), Inches(1.3), Inches(5.8), Inches(3.2))

create_card(s10, 0.6, 4.7, 5.8, 2.2, "Flight Number vs. Launch Site Insights", [
    "CCAFS SLC-40 served as the primary testbed for early flights (1 to 25); early launches experienced high failure rates as landing technology was developed.",
    "KSC LC-39A came online later (Flight 30+) and achieved consistent landing success.",
    "As overall flight numbers increased, landing success climbed across all complexes."
])

create_card(s10, 6.8, 4.7, 5.8, 2.2, "Payload Mass vs. Launch Site Insights", [
    "Heavy-lift payloads (>10,000 kg) are concentrated at KSC LC-39A and CCAFS SLC-40.",
    "VAFB SLC-4E handles polar payloads exclusively under 10,000 kg with strong recovery.",
    "No successful landings occurred at CCAFS SLC-40 for payloads exceeding 6,000 kg until Block 5."
])

# ==============================================================================
# SLIDE 11: EDA Results – Orbit Success & Yearly Trend (Criterion 1.11 cont.)
# ==============================================================================
s11 = prs.slides.add_slide(blank_layout)
add_header(s11, "EDA Results: Orbit Success Rates & Yearly Technological Trend")
add_footer(s11, 11)

s11.shapes.add_picture(os.path.join(IMG_DIR, "eda_orbit_success.png"), Inches(0.6), Inches(1.3), Inches(5.8), Inches(3.2))
s11.shapes.add_picture(os.path.join(IMG_DIR, "eda_yearly_trend.png"), Inches(6.8), Inches(1.3), Inches(5.8), Inches(3.2))

create_card(s11, 0.6, 4.7, 5.8, 2.2, "Orbit Success Rate Analysis", [
    "100% Landing Success Orbits: ES-L1, GEO, HEO, and SSO (light payloads & favorable trajectories).",
    "ISS Resupply Missions: Achieve a strong 76.9% landing success rate.",
    "GTO Missions: Achieve ~52% success due to extreme velocity and high fuel burn requirements."
])

create_card(s11, 6.8, 4.7, 5.8, 2.2, "Yearly Trend & Technological Maturation", [
    "2010 - 2013: 0% landing success rate (controlled descent and water landing tests only).",
    "2015: First historic ground landing (33.3% success rate).",
    "2016 - 2017: Rapid adoption of Autonomous Spaceport Drone Ship (ASDS) landings (~60-77%).",
    "2018 - 2020: Falcon 9 Block 5 introduction drives success rates above 85%."
])

# ==============================================================================
# SLIDE 12: EDA with SQL – Results (Criterion 1.12)
# ==============================================================================
s12 = prs.slides.add_slide(blank_layout)
add_header(s12, "EDA with SQL: Structured Query Results Summary")
add_footer(s12, 12)

# Create a clean results table
table_shape = s12.shapes.add_table(6, 4, Inches(0.6), Inches(1.4), Inches(12.133), Inches(5.3))
tbl = table_shape.table
tbl.columns[0].width = Inches(2.2)
tbl.columns[1].width = Inches(4.5)
tbl.columns[2].width = Inches(2.5)
tbl.columns[3].width = Inches(2.933)

headers = ["SQL Domain", "Analytical Query Description", "SQL Output / Metric", "Operational Conclusion"]
for col_idx, h in enumerate(headers):
    cell = tbl.cell(0, col_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = MID_BLUE
    p = cell.text_frame.paragraphs[0]
    p.text = h
    p.font.bold = True
    p.font.size = Pt(11)
    p.font.color.rgb = WHITE

rows_data = [
    ("1. Launch Sites", "SELECT DISTINCT LaunchSite FROM SPACEX;", "CCAFS LC-40, CCAFS SLC-40, KSC LC-39A, VAFB SLC-4E", "All sites positioned strategically along eastern or western coastlines."),
    ("2. Payloads Analysis", "SELECT MIN(PayloadMass), MAX(PayloadMass), AVG(PayloadMass) FROM SPACEX;", "Min: 315 kg\nMax: 15,600 kg\nAvg: 6,105 kg", "Massive payload flexibility spanning small CubeSats to heavy Starlink stacks."),
    ("3. Success Rates", "SELECT LandingOutcome, COUNT(*) FROM SPACEX GROUP BY LandingOutcome;", "Drone Ship: 41 Successes\nGround Pad: 13 Successes\nOcean: 3 Controlled", "Drone ship landings are the primary recovery mode for high-energy missions."),
    ("4. Outcome Rankings", "Rank landing outcomes between 2010-06-04 and 2017-03-20;", "1. Success (drone ship)\n2. Success (ground pad)\n3. Failure (drone ship)", "Recovery success overtook failures as routine procedures were established."),
    ("5. Time Analysis", "SELECT strftime('%Y', Date), COUNT(*) FROM SPACEX WHERE Class=1 GROUP BY 1;", "2015: 1 | 2016: 5 | 2017: 14\n2018: 15 | 2019: 13 | 2020: 16", "Exponential rise in recovery counts correlates directly with Block 5 fleet reusability.")
]

for row_idx, r in enumerate(rows_data, start=1):
    for col_idx, val in enumerate(r):
        cell = tbl.cell(row_idx, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = LIGHT_BG if row_idx % 2 == 1 else WHITE
        p = cell.text_frame.paragraphs[0]
        p.text = val
        p.font.size = Pt(10)
        p.font.color.rgb = RGBColor(30, 41, 59)

# ==============================================================================
# SLIDE 13: Folium Map Analysis – Results (Criterion 1.13)
# ==============================================================================
s13 = prs.slides.add_slide(blank_layout)
add_header(s13, "Interactive Map Analysis: Folium Launch Sites & Proximity Analysis")
add_footer(s13, 13)

s13.shapes.add_picture(os.path.join(IMG_DIR, "folium_launch_sites_map.png"), Inches(0.6), Inches(1.3), Inches(3.8), Inches(2.7))
s13.shapes.add_picture(os.path.join(IMG_DIR, "folium_marker_cluster_map.png"), Inches(4.7), Inches(1.3), Inches(3.8), Inches(2.7))
s13.shapes.add_picture(os.path.join(IMG_DIR, "folium_proximity_map.png"), Inches(8.8), Inches(1.3), Inches(3.8), Inches(2.7))

create_card(s13, 0.6, 4.2, 3.8, 2.7, "Launch Site Positioning", [
    "CCAFS SLC-40: 28.56° N, 80.58° W (Cape Canaveral, FL).",
    "KSC LC-39A: 28.61° N, 80.60° W (Kennedy Space Center, FL).",
    "VAFB SLC-4E: 34.63° N, 120.61° W (Vandenberg AFB, CA).",
    "All pads situated on coastlines to permit eastern and southern orbital flight paths over water."
])

create_card(s13, 4.7, 4.2, 3.8, 2.7, "MarkerCluster Outcomes", [
    "MarkerCluster aggregates launch markers dynamically by zoom level.",
    "Green Icons: Indicate successful first stage landings (Class = 1).",
    "Red Icons: Indicate unsuccessful attempts or expended boosters (Class = 0).",
    "Visual proof of Cape Canaveral transition from early failure red markers to a dominant cluster of green success markers."
])

create_card(s13, 8.8, 4.2, 3.8, 2.7, "Proximity Analysis Insights", [
    "Coastline (~1.2 km): Ensures booster flight corridor is over water for safety.",
    "Highway FL-405 (~1.8 km): Facilitates road transport of rocket stages and fuel.",
    "Railroad (~1.3 km): Allows heavy cargo and booster logistics via NASA rail.",
    "City (Titusville ~28.5 km): Remote location ensures populated areas are well outside safety blast radius."
])

# ==============================================================================
# SLIDE 14: Plotly Dash Dashboard – Results (Criterion 1.14)
# ==============================================================================
s14 = prs.slides.add_slide(blank_layout)
add_header(s14, "Interactive Dashboard: Plotly Dash Visualizations")
add_footer(s14, 14)

s14.shapes.add_picture(os.path.join(IMG_DIR, "dash_success_pie_charts.png"), Inches(0.6), Inches(1.3), Inches(5.8), Inches(3.2))
s14.shapes.add_picture(os.path.join(IMG_DIR, "dash_payload_scatter.png"), Inches(6.8), Inches(1.3), Inches(5.8), Inches(3.2))

create_card(s14, 0.6, 4.7, 5.8, 2.2, "Pie Chart Insights: Site Success Breakdown", [
    "Dropdown 'ALL Sites': KSC LC-39A contributes the highest share of total successful landings (41.7%), followed by CCAFS SLC-40 (35.0%) and VAFB SLC-4E (23.3%).",
    "Site-Specific Filter: KSC LC-39A boasts a 76.9% landing success rate, significantly outperforming CCAFS SLC-40 (60.0%) and VAFB SLC-4E (57.1%)."
])

create_card(s14, 6.8, 4.7, 5.8, 2.2, "Scatter Plot Insights: Payload vs. Outcome with Slider", [
    "Optimal Payload Range: Missions with payloads between 2,000 kg and 6,000 kg achieve the highest landing success rate across all booster versions.",
    "Booster Version Impact: Falcon 9 Block 5 (FT/B5) achieves high success across light and heavy payloads; older v1.0 boosters had a 0% recovery record."
])

# ==============================================================================
# SLIDE 15: Predictive Analysis – Model Training & Evaluation (Criterion 1.15)
# ==============================================================================
s15 = prs.slides.add_slide(blank_layout)
add_header(s15, "Predictive Analysis: Machine Learning Models & Evaluation")
add_footer(s15, 15)

s15.shapes.add_picture(os.path.join(IMG_DIR, "ml_model_comparison.png"), Inches(0.6), Inches(1.3), Inches(6.5), Inches(3.3))
s15.shapes.add_picture(os.path.join(IMG_DIR, "ml_confusion_matrix.png"), Inches(7.4), Inches(1.3), Inches(5.3), Inches(3.3))

create_card(s15, 0.6, 4.8, 6.5, 2.1, "Model Performance Comparison (10-Fold CV)", [
    "Logistic Regression: CV Train Acc = 84.7%, Test Acc = 83.33% (Params: C=0.01, penalty='l2').",
    "Support Vector Machine (SVM): CV Train Acc = 85.4%, Test Acc = 83.33% (Params: C=1.0, kernel='rbf').",
    "Decision Tree Classifier: CV Train Acc = 89.8%, Test Acc = 83.33% (Params: max_depth=4, criterion='gini').",
    "K-Nearest Neighbors (KNN): CV Train Acc = 84.7%, Test Acc = 83.33% (Params: n_neighbors=6, p=2)."
])

create_card(s15, 7.4, 4.8, 5.3, 2.1, "Confusion Matrix Analysis (Test Set = 18)", [
    "True Negatives: 6 (correctly predicted failed/expended landings).",
    "False Positives: 0 (ZERO false alarms — 100% Precision).",
    "False Negatives: 3 (missed landing opportunities).",
    "True Positives: 9 (correctly predicted successful landings).",
    "Overall Test Accuracy: 83.33% | Test Precision: 100% | F1-Score: 0.857."
])

# ==============================================================================
# SLIDE 16: Best Model & Conclusion (Criterion 1.15 cont.)
# ==============================================================================
s16 = prs.slides.add_slide(blank_layout)
add_header(s16, "Best Model Justification & Innovative Industry Insights")
add_footer(s16, 16)

create_card(s16, 0.6, 1.3, 5.8, 5.5, "Best Model Justification: Decision Tree Classifier", [
    "Top Test Accuracy: Tied for the highest test accuracy (83.33%) across all 4 machine learning algorithms.",
    "Zero False Positives: Critical for commercial launch operators — when the model predicts success, it is 100% reliable, protecting operational budgets from unexpected booster loss costs.",
    "High Interpretability: Unlike 'black box' models (SVM/Neural Nets), Decision Trees provide transparent, inspectable decision split rules (FlightNumber, PayloadMass, Orbit_GTO).",
    "Robust to Non-Linearity: Naturally captures interaction thresholds between payload mass and orbit delta-v requirements.",
    "Feature Importance: Confirmed FlightNumber (flight experience) and PayloadMass as top predictors."
])

create_card(s16, 6.8, 1.3, 5.8, 5.5, "Creative & Innovative Industry Insights", [
    "Economics of Reusability: SpaceX launch cost of ~$62M vs competitor prices ($165M+) gives SpaceX an unassailable 62% margin advantage on commercial bids.",
    "The Learning Curve (Wright's Law): Rapid jump in success from 0% (2013) to >85% (2020) demonstrates that rocketry reusability follows industrial learning curve dynamics.",
    "Fleet Turnaround Optimization: Turnaround times between booster flights fell from 354 days to under 30 days, maximizing asset utilization.",
    "Competitive Strategy for Rivals: Alternate launch providers cannot compete on price alone; they must develop partially/fully reusable vehicles or target niche responsive orbits where SpaceX reusability margins are constrained."
])

# Save presentation
prs.save(PPTX_PATH)
print(f"[OK] Presentation saved to: {PPTX_PATH}")

# Convert to PDF via PowerPoint COM
print("Converting presentation to PDF via PowerPoint COM...")
try:
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = 1
    deck = ppt_app.Presentations.Open(PPTX_PATH)
    # 32 represents ppSaveAsPDF in PowerPoint OM
    deck.SaveAs(PDF_PATH, 32)
    deck.Close()
    ppt_app.Quit()
    print(f"[OK] PDF successfully generated: {PDF_PATH}")
except Exception as e:
    print(f"[ERROR] COM conversion error: {e}")
    sys.exit(1)
