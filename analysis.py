import pandas as pd
import numpy as np
from geopy.distance import geodesic
import folium
from folium.plugins import MarkerCluster
import re
import os
import webbrowser
from datetime import datetime


# 1.LOAD DATA
file_path = "Data.xlsx"

df = pd.read_excel(file_path, sheet_name="Order")
rates_df = pd.read_excel(file_path, sheet_name="Rates_Master")

# clear column names from hidden spaces
df.columns = df.columns.str.strip()
rates_df.columns = rates_df.columns.str.strip()

print("column Rates_Master:", rates_df.columns.tolist())


# 2.CLEAR RATES
def clean_rupiah(x):
    if pd.isna(x):
        return 0
    x = str(x)
    x = re.sub(r"[^\d]", "", x)
    return int(x) if x != "" else 0

# auto fetch all zone columns (safer)
zona_cols = [col for col in rates_df.columns if "Zona" in col]

for col in zona_cols:
    rates_df[col] = rates_df[col].apply(clean_rupiah)


# 3.STORE COORDINATES
store_locations = {
    "Bandung": (-6.9175, 107.6191),
    "Batam": (1.1301, 104.0529),
    "Bekasi": (-6.2383, 106.9756),
    "Denpasar": (-8.6500, 115.2167),
    "Depok": (-6.4025, 106.7942),
    "Jakarta": (-6.2088, 106.8456),
    "Jayapura": (-2.5916, 140.6690),
    "Madiun": (-7.6298, 111.5239),
    "Makassar": (-5.1477, 119.4327),
    "Medan": (3.5952, 98.6722),
    "Palembang": (-2.9761, 104.7754),
    "Pekanbaru": (0.5071, 101.4478),
    "Pontianak": (-0.0263, 109.3425),
    "Samarinda": (-0.4948, 117.1436),
    "Semarang": (-6.9667, 110.4167),
    "Surabaya": (-7.2575, 112.7521),
    "Surakarta": (-7.5666, 110.8167),
    "Tangerang": (-6.1783, 106.6319),
    "Ternate": (0.7905, 127.3753),
    "Tual": (-5.6311, 132.7441),
    "Yogyakarta": (-7.7956, 110.3695)

}


# 4.COORDINATE VALIDATION
df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")

def valid_coord(lat, lon):
    return (
        pd.notnull(lat) and pd.notnull(lon)
        and -90 <= lat <= 90
        and -180 <= lon <= 180
    )


# 5.CALCULATE DISTANCE & ZONE
def get_nearest_store(lat, lon):
    min_dist = float("inf")
    nearest_store = None
    
    for store, coord in store_locations.items():
        dist = geodesic((lat, lon), coord).km
        if dist < min_dist:
            min_dist = dist
            nearest_store = store
            
    return nearest_store, min_dist

def get_zona(distance):
    if distance < 300:
        return zona_cols[0]
    elif distance <= 1500:
        return zona_cols[1]
    else:
        return zona_cols[2]


# 6.CALCULATE SHIPPING COSTS
def hitung_ongkir(row):

    if not valid_coord(row["Latitude"], row["Longitude"]):
        return pd.Series([None, 0, None, 0, 0, 0])
    
    store, distance = get_nearest_store(row["Latitude"], row["Longitude"])
    zona_col = get_zona(distance)
    
    weight = np.ceil(pd.to_numeric(row["Weight"], errors="coerce"))
    if pd.isna(weight):
        weight = 0
    
    match = rates_df[
        (rates_df["Courier"] == row["Courier"]) &
        (rates_df["Service"] == row["Service"])
    ]
    
    if not match.empty and zona_col in rates_df.columns:
        rate = match.iloc[0][zona_col]
        ongkir = rate * weight
    else:
        rate = 0
        ongkir = 0
    
    return pd.Series([store, distance, zona_col, weight, rate, ongkir])

df[[
    "Nearest_Store",
    "Distance_km",
    "Zona",
    "Weight_Bulat",
    "Rates_kg",
    "Shipping_Cost"
]] = df.apply(hitung_ongkir, axis=1)


# 7.RUPIAH FORMAT
df["Shipping_Cost"] = pd.to_numeric(df["Shipping_Cost"], errors="coerce").fillna(0)

def format_rupiah(x):
    return "Rp {:,.0f}".format(x).replace(",", ".")

df["Shipping_Cost_Rp"] = df["Shipping_Cost"].apply(format_rupiah)

# delete the Shipping Cost column
df.drop(columns=["Shipping_Cost"], inplace=True)


# 8.SAVE EXCEL
output_file = "Product_Sales_Analysis.xlsx"
if os.path.exists(output_file):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"Product_Sales_Analysis{timestamp}.xlsx"
df.to_excel(output_file, index=False)


# 9.MAP VISUALIZATION
m = folium.Map(location=[-2.5,118], zoom_start=5, tiles="CartoDB positron")

marker_cluster = MarkerCluster().add_to(m)

#shop marker
for store, coord in store_locations.items():
    folium.Marker(
        location=coord,
        popup=f"<b>Toko {store}</b>",
        icon=folium.Icon(color="red", icon="home")
    ).add_to(m)
warna_courier = {
    "JNE": "blue",
    "JNT": "green",
    "Shopee Express": "purple",
    "Anteraja": "orange",
    "SiCepat": "cadetblue",
    "Pos Indonesia": "darkred"
}
for _, row in df.iterrows():
    if valid_coord(row["Latitude"], row["Longitude"]):

        warna = warna_courier.get(row["Courier"], "gray")

        popup_html = f"""
        <b>Courier:</b> {row['Courier']}<br>
        <b>Service:</b> {row['Service']}<br>
        <b>Ongkir:</b> {row['Shipping_Cost_Rp']}<br>
        <b>Jarak:</b> {row['Distance_km']:.1f} km
        """
        folium.CircleMarker(
            location=[row["Latitude"], row["Longitude"]],
            radius=6,
            color=warna,
            fill=True,
            fill_opacity=0.7,
            popup=popup_html
        ).add_to(marker_cluster)
m.save("Interactive_Shipping_Map.html")
map_path = os.path.abspath("Interactive_Shipping_Map.html")
webbrowser.open(f"file://{map_path}") 

print("✅ Completed Without Error.")
print(f"File Excel: {output_file}")
print("File Maps: Interactive_Shipping_Map.html")
