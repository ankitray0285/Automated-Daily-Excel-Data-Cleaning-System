# Automated-Daily-Excel-Data-Cleaning-System
This repository contains an Excel + Power Query automation system that cleans and consolidates daily Excel files dropped in a folder.

The system automatically standardizes, validates, and prepares data extracted from e-HRMS, saving manual effort and ensuring consistency.

🔹 Project Overview

Many HR and business systems generate multiple Excel files daily. These files often contain inconsistent data, missing values, or duplicates.

This project automates the cleaning and consolidation process:

Extract files from e-HRMS into a folder
Automatically clean and standardize them using Power Query
Save the processed data in a master dataset for further analysis
🔹 Features
Automated File Processing: Reads all Excel files from a designated folder
Data Cleaning & Transformation:
Standardizes column names (Date, Employee, Department, etc.)
Removes duplicates
Handles null and negative values
Trims spaces in text fields
Standardizes date formats (DD-MM-YYYY)
Optional: add calculated columns (e.g., Weekday, Month)
Master Dataset Generation: Consolidated cleaned data is saved automatically
🔹 Folder Structure
project-folder/
│
├── daily_files/          # Folder to drop extracted Excel sheets from e-HRMS
├── CleanedMaster.xlsx    # Excel file where Power Query saves cleaned data
├── README.md             # This file
└── screenshots/          # Screenshots of Power Query steps
🔹 Setup Instructions
Create Daily Files Folder
Example: C:\Daily_Files\
Drop all extracted Excel files from e-HRMS here
Open CleanedMaster.xlsx
Contains Power Query connections for cleaning
Set Folder Path in Power Query
Data → Get Data → From File → From Folder → Select daily_files folder
Combine & Transform Data
Power Query will automatically:
Standardize column names
Remove duplicates
Trim spaces
Handle nulls / negative values
Standardize date formats
Load Cleaned Master Dataset
Click Close & Load → Master dataset tab updates automatically
Refresh Daily
Drop new files → Open workbook → Click Refresh All
Power Query cleans and updates master dataset automatically
🔹 Dataset Cleaning Rules
Column	Source Variants	Data Handling
Date	Transaction Date, Entry Date	Standardize → DD-MM-YYYY
Employee	Name, Emp Name	Trim spaces
Department	Dept, Team	Standardize names
Salary	Amount, Pay	Null/negative → 0
Comments	Remarks, Notes	Optional

Validation:

Numeric fields ≥ 0
Required fields (Date, Employee, Department) must not be blank
🔹 Key Benefits
Saves 2–5 hours daily of manual cleaning
Eliminates human error
Standardized and validated data for further analysis
Scalable and repeatable
🔹 Screenshots
Include images showing:
Folder structure
Power Query steps (combine, transform, clean)
Master dataset after cleaning
🔹 How to Use
Extract files from e-HRMS into the daily_files folder
Open CleanedMaster.xlsx
Click Refresh All → Master dataset updates automatically
Use the cleaned master dataset for reporting or further processing
