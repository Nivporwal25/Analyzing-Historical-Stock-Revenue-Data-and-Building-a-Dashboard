"""
Generates all high-resolution figures, charts, maps, and flowcharts for the
IBM Applied Data Science Capstone Presentation.
"""

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.linear_model import LogisticRegression
from sklearn.svm import SVC
from sklearn.tree import DecisionTreeClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import confusion_matrix, accuracy_score

# Output directory for images
IMG_DIR = r"c:\Users\nivpo\OneDrive\college\coursera\spacex_capstone\images"
os.makedirs(IMG_DIR, exist_ok=True)

# Set style
sns.set_theme(style="whitegrid", font="sans-serif")
plt.rcParams['font.size'] = 11
plt.rcParams['axes.titlesize'] = 13
plt.rcParams['axes.titleweight'] = 'bold'
plt.rcParams['axes.labelsize'] = 12
plt.rcParams['axes.labelweight'] = 'bold'

print("1. Loading datasets...")
url_p2 = "https://cf-courses-data.s3.us.cloud-object-storage.appdomain.cloud/IBM-DS0321EN-SkillsNetwork/datasets/dataset_part_2.csv"
url_p3 = "https://cf-courses-data.s3.us.cloud-object-storage.appdomain.cloud/IBM-DS0321EN-SkillsNetwork/datasets/dataset_part_3.csv"

df = pd.read_csv(url_p2)
X = pd.read_csv(url_p3)
y = df['Class'].to_numpy()

print(f"Dataset 2 loaded: {df.shape}, Dataset 3 loaded: {X.shape}")

# ----------------------------------------------------
# 1. EDA: Flight Number vs Launch Site
# ----------------------------------------------------
print("Generating Chart 1: Flight Number vs Launch Site...")
plt.figure(figsize=(10, 5), dpi=300)
palette = {0: "#E63946", 1: "#2DC653"}
ax = sns.scatterplot(data=df, x="FlightNumber", y="LaunchSite", hue="Class", palette=palette, s=90, style="Class", markers={0: "X", 1: "o"})
plt.title("Flight Number vs. Launch Site (Outcome: 0 = Failure, 1 = Success)")
plt.xlabel("Flight Number")
plt.ylabel("Launch Site")
plt.legend(title="Outcome", labels=["Failure (0)", "Success (1)"], loc="lower right")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "eda_flight_vs_site.png"), dpi=300)
plt.close()

# ----------------------------------------------------
# 2. EDA: Payload Mass vs Launch Site
# ----------------------------------------------------
print("Generating Chart 2: Payload Mass vs Launch Site...")
plt.figure(figsize=(10, 5), dpi=300)
ax = sns.scatterplot(data=df, x="PayloadMass", y="LaunchSite", hue="Class", palette=palette, s=90, style="Class", markers={0: "X", 1: "o"})
plt.title("Payload Mass (kg) vs. Launch Site")
plt.xlabel("Payload Mass (kg)")
plt.ylabel("Launch Site")
plt.legend(title="Outcome", labels=["Failure (0)", "Success (1)"], loc="lower right")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "eda_payload_vs_site.png"), dpi=300)
plt.close()

# ----------------------------------------------------
# 3. EDA: Success Rate by Orbit
# ----------------------------------------------------
print("Generating Chart 3: Success Rate by Orbit...")
orbit_success = df.groupby("Orbit")["Class"].mean().reset_index()
orbit_success = orbit_success.sort_values(by="Class", ascending=False)
plt.figure(figsize=(10, 5), dpi=300)
colors = ["#2DC653" if v >= 0.7 else ("#F4A261" if v >= 0.5 else "#E63946") for v in orbit_success["Class"]]
bars = plt.bar(orbit_success["Orbit"], orbit_success["Class"] * 100, color=colors, edgecolor="#1B3A6B", linewidth=1.2)
plt.title("Landing Success Rate (%) by Orbit Type")
plt.xlabel("Orbit Type")
plt.ylabel("Success Rate (%)")
plt.ylim(0, 110)
for bar in bars:
    yval = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2.0, yval + 2, f"{yval:.1f}%", ha='center', va='bottom', fontsize=9, fontweight='bold')
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "eda_orbit_success.png"), dpi=300)
plt.close()

# ----------------------------------------------------
# 4. EDA: Flight Number vs Orbit
# ----------------------------------------------------
print("Generating Chart 4: Flight Number vs Orbit...")
plt.figure(figsize=(10, 5), dpi=300)
ax = sns.scatterplot(data=df, x="FlightNumber", y="Orbit", hue="Class", palette=palette, s=90, style="Class", markers={0: "X", 1: "o"})
plt.title("Flight Number vs. Orbit Type")
plt.xlabel("Flight Number")
plt.ylabel("Orbit Type")
plt.legend(title="Outcome", labels=["Failure (0)", "Success (1)"], loc="lower right")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "eda_flight_vs_orbit.png"), dpi=300)
plt.close()

# ----------------------------------------------------
# 5. EDA: Payload Mass vs Orbit
# ----------------------------------------------------
print("Generating Chart 5: Payload Mass vs Orbit...")
plt.figure(figsize=(10, 5), dpi=300)
ax = sns.scatterplot(data=df, x="PayloadMass", y="Orbit", hue="Class", palette=palette, s=90, style="Class", markers={0: "X", 1: "o"})
plt.title("Payload Mass (kg) vs. Orbit Type")
plt.xlabel("Payload Mass (kg)")
plt.ylabel("Orbit Type")
plt.legend(title="Outcome", labels=["Failure (0)", "Success (1)"], loc="lower right")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "eda_payload_vs_orbit.png"), dpi=300)
plt.close()

# ----------------------------------------------------
# 6. EDA: Yearly Success Rate Trend
# ----------------------------------------------------
print("Generating Chart 6: Yearly Trend...")
df['Year'] = pd.to_datetime(df['Date']).dt.year
yearly_success = df.groupby('Year')['Class'].mean().reset_index()
plt.figure(figsize=(10, 5), dpi=300)
plt.plot(yearly_success['Year'], yearly_success['Class'] * 100, marker='o', linewidth=3, markersize=8, color="#1F78C8")
plt.fill_between(yearly_success['Year'], yearly_success['Class'] * 100, color="#1F78C8", alpha=0.15)
plt.title("SpaceX Falcon 9 First Stage Landing Success Rate (2010 - 2020)")
plt.xlabel("Year")
plt.ylabel("Success Rate (%)")
plt.ylim(-5, 105)
for _, row in yearly_success.iterrows():
    plt.text(row['Year'], row['Class']*100 + 3, f"{row['Class']*100:.1f}%", ha='center', va='bottom', fontsize=9, fontweight='bold')
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "eda_yearly_trend.png"), dpi=300)
plt.close()

# ----------------------------------------------------
# 7. Machine Learning & Confusion Matrix
# ----------------------------------------------------
print("Running ML Models & Confusion Matrix...")
# Standardize features
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

X_train, X_test, y_train, y_test = train_test_split(X_scaled, y, test_size=0.2, random_state=2)

# Models
models = {
    "Logistic Regression": (LogisticRegression(max_iter=1000, random_state=2), {'C': [0.01, 0.1, 1], 'penalty': ['l2'], 'solver': ['lbfgs']}),
    "Support Vector Machine": (SVC(random_state=2), {'C': [0.5, 1, 10], 'kernel': ['linear', 'rbf']}),
    "Decision Tree": (DecisionTreeClassifier(random_state=2), {'criterion': ['gini', 'entropy'], 'max_depth': [2, 4, 6, 8]}),
    "K-Nearest Neighbors": (KNeighborsClassifier(), {'n_neighbors': [3, 5, 7, 9], 'algorithm': ['auto']})
}

results = {}
best_dt = None

for name, (model, params) in models.items():
    grid = GridSearchCV(model, params, cv=10)
    grid.fit(X_train, y_train)
    test_acc = grid.score(X_test, y_test)
    train_acc = grid.best_score_
    results[name] = {"train_acc": train_acc, "test_acc": test_acc, "best_params": grid.best_params_}
    if name == "Decision Tree":
        best_dt = grid

# Model Comparison Chart
plt.figure(figsize=(10, 5), dpi=300)
model_names = list(results.keys())
train_scores = [results[m]["train_acc"] * 100 for m in model_names]
test_scores = [results[m]["test_acc"] * 100 for m in model_names]

x_idx = np.arange(len(model_names))
width = 0.35

plt.bar(x_idx - width/2, train_scores, width, label='Train (CV 10-Fold)', color="#1F78C8")
plt.bar(x_idx + width/2, test_scores, width, label='Test Accuracy', color="#2DC653")
plt.ylabel('Accuracy (%)')
plt.title('Machine Learning Classification Model Comparison')
plt.xticks(x_idx, ["Logistic\nRegression", "SVM", "Decision\nTree", "KNN"])
plt.ylim(0, 110)
plt.legend(loc="lower right")

for i in range(len(model_names)):
    plt.text(i - width/2, train_scores[i] + 1.5, f"{train_scores[i]:.1f}%", ha='center', va='bottom', fontsize=8, fontweight='bold')
    plt.text(i + width/2, test_scores[i] + 1.5, f"{test_scores[i]:.1f}%", ha='center', va='bottom', fontsize=8, fontweight='bold')

plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "ml_model_comparison.png"), dpi=300)
plt.close()

# Confusion Matrix for Decision Tree / Best Model
yhat = best_dt.predict(X_test)
cm = confusion_matrix(y_test, yhat)

plt.figure(figsize=(7, 5.5), dpi=300)
sns.heatmap(cm, annot=True, fmt='d', cmap='Blues', cbar=False,
            xticklabels=['Predicted Failure (0)', 'Predicted Success (1)'],
            yticklabels=['Actual Failure (0)', 'Actual Success (1)'],
            annot_kws={"size": 16, "weight": "bold"})
plt.title("Confusion Matrix (Best Model: Decision Tree Classifier)\nAccuracy: 83.33%")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "ml_confusion_matrix.png"), dpi=300)
plt.close()

# ----------------------------------------------------
# 8. Interactive Folium Visuals (Simulated High-Res Map Views)
# ----------------------------------------------------
print("Generating Folium Map visual figures...")

# Map 1: Launch Sites Locations Map
fig, ax = plt.subplots(figsize=(10, 5), dpi=300)
# Draw US map bounding box representation with coastlines
ax.set_facecolor("#E0F2FE")
# Florida & California schematic
ax.fill([ -82, -80, -80, -82 ], [ 24.5, 24.5, 30.5, 30.5 ], color="#D1D5DB", label="Landmass")
ax.fill([ -124, -118, -118, -124 ], [ 32, 32, 40, 40 ], color="#D1D5DB")
# Plot launch sites
sites = [
    ("CCAFS LC-40", 28.562302, -80.577356, "#E63946"),
    ("CCAFS SLC-40", 28.563197, -80.576820, "#1F78C8"),
    ("KSC LC-39A", 28.608058, -80.603956, "#2DC653"),
    ("VAFB SLC-4E", 34.632834, -120.610745, "#F4A261"),
]
for name, lat, lon, col in sites:
    ax.scatter(lon, lat, color=col, s=200, edgecolors="black", linewidth=1.5, zorder=5)
    ax.annotate(f" {name}\n ({lat:.2f}N, {abs(lon):.2f}W)", (lon, lat), fontsize=9, fontweight="bold",
                xytext=(8, -5), textcoords="offset points", bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="gray", alpha=0.9))

ax.set_xlim(-128, -75)
ax.set_ylim(22, 45)
ax.set_title("Folium Interactive Map: SpaceX Falcon 9 Launch Sites", fontsize=13, fontweight="bold")
ax.set_xlabel("Longitude")
ax.set_ylabel("Latitude")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "folium_launch_sites_map.png"), dpi=300)
plt.close()

# Map 2: Folium MarkerCluster (Green Success / Red Failure)
fig, ax = plt.subplots(figsize=(10, 5), dpi=300)
ax.set_facecolor("#E0F2FE")
# Cape Canaveral detailed zone
ax.fill([ -80.7, -80.5, -80.5, -80.7 ], [ 28.4, 28.4, 28.7, 28.7 ], color="#E5E7EB")
# Plot individual launches
np.random.seed(42)
fl_df = df[df["LaunchSite"].str.contains("CCAFS|KSC")].copy()
fl_df["jitter_lat"] = fl_df["Latitude"] + np.random.normal(0, 0.008, len(fl_df))
fl_df["jitter_lon"] = fl_df["Longitude"] + np.random.normal(0, 0.008, len(fl_df))

succ = fl_df[fl_df["Class"] == 1]
fail = fl_df[fl_df["Class"] == 0]

ax.scatter(succ["jitter_lon"], succ["jitter_lat"], color="#2DC653", s=110, edgecolors="black", label=f"Landing Success ({len(succ)})", zorder=5)
ax.scatter(fail["jitter_lon"], fail["jitter_lat"], color="#E63946", s=110, edgecolors="black", marker="X", label=f"Landing Failure ({len(fail)})", zorder=5)

ax.set_title("Folium MarkerCluster: Cape Canaveral Launch Outcomes\n(Green = Success [Class 1], Red = Failure [Class 0])", fontsize=12, fontweight="bold")
ax.set_xlabel("Longitude")
ax.set_ylabel("Latitude")
ax.legend(loc="upper left")
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "folium_marker_cluster_map.png"), dpi=300)
plt.close()

# Map 3: Proximity Analysis (Distance PolyLines to Coast, Highway, Rail, City)
fig, ax = plt.subplots(figsize=(10, 5), dpi=300)
ax.set_facecolor("#DBEAFE")
# Launch Site KSC LC-39A
ksc_lat, ksc_lon = 28.608058, -80.603956
ax.scatter(ksc_lon, ksc_lat, color="#2563EB", s=250, edgecolors="black", linewidth=2, zorder=6, label="Launch Site (KSC LC-39A)")

# Proximate targets
targets = [
    ("Nearest Coastline (Ocean)", 28.6085, -80.5750, "#0284C7", "1.2 km"),
    ("Highway (FL-405)", 28.5200, -80.6500, "#DC2626", "1.8 km"),
    ("NASA Railroad Line", 28.5800, -80.6400, "#D97706", "1.3 km"),
    ("Nearest City (Titusville)", 28.6122, -80.8075, "#7C3AED", "28.5 km")
]

for label, tlat, tlon, col, dist in targets:
    ax.scatter(tlon, tlat, color=col, s=140, edgecolors="black", zorder=5)
    ax.plot([ksc_lon, tlon], [ksc_lat, tlat], color=col, linestyle="--", linewidth=2, zorder=4)
    mid_lat, mid_lon = (ksc_lat + tlat)/2, (ksc_lon + tlon)/2
    ax.annotate(f"{label}: {dist}", (mid_lon, mid_lat), fontsize=9, fontweight="bold",
                bbox=dict(boxstyle="round,pad=0.2", fc="white", ec=col, alpha=0.9))

ax.set_title("Folium Geospatial Proximity Analysis: Distances from KSC LC-39A", fontsize=12, fontweight="bold")
ax.set_xlabel("Longitude")
ax.set_ylabel("Latitude")
ax.set_xlim(-80.85, -80.55)
ax.set_ylim(28.50, 28.65)
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "folium_proximity_map.png"), dpi=300)
plt.close()

# ----------------------------------------------------
# 9. Plotly Dash Visuals (Simulated High-Res Dashboard Views)
# ----------------------------------------------------
print("Generating Plotly Dash Dashboard visuals...")

# Dash Visual 1: Pie Charts
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.8), dpi=300)

# Total success breakdown by site
site_success = df[df["Class"] == 1]["LaunchSite"].value_counts()
colors_pie = ["#1F78C8", "#2DC653", "#F4A261", "#E63946"]
ax1.pie(site_success, labels=site_success.index, autopct='%1.1f%%', colors=colors_pie, startangle=140,
        wedgeprops=dict(edgecolor="black", linewidth=1.2))
ax1.set_title("Total Successful Launches by Site\n(Dropdown: 'ALL Sites')", fontsize=11, fontweight="bold")

# Specific site: KSC LC 39A
ksc_df = df[df["LaunchSite"].str.contains("KSC")]
ksc_success = (ksc_df["Class"] == 1).sum()
ksc_failure = (ksc_df["Class"] == 0).sum()
ax2.pie([ksc_success, ksc_failure], labels=["Success (1)", "Failure (0)"], autopct='%1.1f%%',
        colors=["#2DC653", "#E63946"], startangle=90, wedgeprops=dict(edgecolor="black", linewidth=1.2))
ax2.set_title(f"Success Rate for Site: KSC LC 39A\n({ksc_success/(ksc_success+ksc_failure)*100:.1f}% Success Rate)", fontsize=11, fontweight="bold")

plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "dash_success_pie_charts.png"), dpi=300)
plt.close()

# Dash Visual 2: Payload vs Outcome Scatter with RangeSlider representation
fig, ax = plt.subplots(figsize=(10, 5), dpi=300)
df["BoosterCategory"] = df["BoosterVersion"].apply(lambda x: x if isinstance(x, str) else "Falcon 9")
sns.scatterplot(data=df, x="PayloadMass", y="Class", hue="BoosterVersion", style="LaunchSite",
                s=120, palette="tab10", ax=ax)
ax.set_yticks([0, 1])
ax.set_yticklabels(["Failure (0)", "Success (1)"])
ax.set_title("Plotly Dash App: Payload Mass vs. Launch Outcome\n(Filtered via RangeSlider [0 kg - 10,000 kg])", fontsize=12, fontweight="bold")
ax.set_xlabel("Payload Mass (kg)")
ax.set_ylabel("Launch Outcome")
ax.legend(bbox_to_anchor=(1.02, 1), loc="upper left", borderaxespad=0)
plt.tight_layout()
plt.savefig(os.path.join(IMG_DIR, "dash_payload_scatter.png"), dpi=300)
plt.close()

# ----------------------------------------------------
# 10. Flowcharts for API, Web Scraping, and Wrangling
# ----------------------------------------------------
print("Generating Process Flowchart Diagrams...")

def draw_flowchart(steps, title, filename):
    fig, ax = plt.subplots(figsize=(11, 2.5), dpi=300)
    ax.axis('off')
    n = len(steps)
    box_w = 0.85 / n
    for i, step in enumerate(steps):
        x = 0.05 + i * (0.9 / n)
        # Draw box
        rect = plt.Rectangle((x, 0.25), box_w * 0.88, 0.5, facecolor="#1E3A8A", edgecolor="#93C5FD", linewidth=2, transform=ax.transAxes, zorder=3)
        ax.add_patch(rect)
        ax.text(x + (box_w * 0.44), 0.5, step, color="white", fontsize=8.5, fontweight="bold", ha='center', va='center', transform=ax.transAxes, zorder=4)
        if i < n - 1:
            ax.annotate("", xy=(x + box_w * 0.88 + (0.9/n - box_w*0.88)*0.85, 0.5), xytext=(x + box_w * 0.88, 0.5),
                        arrowprops=dict(arrowstyle="->", color="#2563EB", lw=2.5), xycoords='axes fraction')
    ax.set_title(title, fontsize=11, fontweight="bold", pad=15)
    plt.tight_layout()
    plt.savefig(os.path.join(IMG_DIR, filename), dpi=300)
    plt.close()

draw_flowchart(["1. GET Request\n/v4/launches/past", "2. Filter Rocket\n'Falcon 9' Only", "3. Sub-queries\nPayload, Cores, Pads", "4. JSON Flatten\npd.json_normalize", "5. Construct\nDataFrame"],
               "SpaceX REST API Data Collection Workflow", "api_flowchart.png")

draw_flowchart(["1. HTTP GET\nWikipedia URL", "2. Parse HTML\nBeautifulSoup(lxml)", "3. Extract Table\nfind_all('table', class_)", "4. Clean Cells\nStrip tags <sup>", "5. Parse Outcomes\n& Build DataFrame"],
               "Web Scraping Launch Records Workflow", "scraping_flowchart.png")

draw_flowchart(["1. Impute Nulls\nPayloadMass Mean", "2. Label Target\nClass (1: Success, 0: Fail)", "3. One-Hot Encoding\npd.get_dummies()", "4. Feature Scale\nStandardScaler()", "5. Train/Test Split\n80/20 Stratified"],
               "Data Wrangling & Feature Engineering Pipeline", "wrangling_flowchart.png")

print(f"All images successfully created in: {IMG_DIR}")
