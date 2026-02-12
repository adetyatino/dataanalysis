# Shipping Cost and Geospatial Analysis

## Overview

This project analyzes shipping costs based on geographic distance and courier service zones. It integrates data processing, geospatial distance calculation, and interactive visualization to support logistics cost evaluation and operational decision-making.

The system calculates delivery distances, classifies shipments into predefined zones, assigns shipping costs based on courier and service type, and generates both an analytical Excel report and an interactive map visualization.

This project demonstrates practical skills in data analytics, geospatial analysis, and business-focused problem solving.

---

## Objectives

- Analyze shipping cost distribution by distance zone  
- Compare courier service performance  
- Visualize shipment distribution geographically  
- Support data-driven logistics decision making  

---

## Tools and Technologies

- Python  
- Pandas  
- NumPy  
- Geopy  
- Folium  
- OpenPyXL  


---

## Methodology

### 1. Data Preparation
The script loads order data from Excel and performs cleaning and formatting.

### 2. Distance Calculation
Geographic distance between origin and destination is calculated using latitude and longitude coordinates with the geodesic formula.

### 3. Zone Classification
Shipments are categorized into:

- Zone 1: < 300 km  
- Zone 2: 300 – 1500 km  
- Zone 3: Inter-island deliveries  

### 4. Cost Assignment
Shipping costs are determined based on courier, service type, and distance zone.

### 5. Output Generation
The script generates:

- `Analysis.xlsx` containing structured analysis results  
- `Interactive_Shipping_Map.html` containing an interactive map visualization  

---

## How to Run

Install dependencies:

```bash
pip install pandas numpy geopy folium openpyxl

Run the Script : analysis.py


## Project Structure

