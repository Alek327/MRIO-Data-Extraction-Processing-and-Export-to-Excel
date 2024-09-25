MRIO Data Extraction, Processing, and Export to Excel for MUST4WATER

Overview

This script is a small part of the larger MUST4WATER project, focusing on the extraction and processing of Multi-Regional Input-Output (MRIO) data. The goal is to analyze trade flows and economic interactions between different regions using Input-Output Tables (IOT) and trade data. The processed data is exported to Excel for further analysis and reporting within the scope of the project.

Key Objectives:
Extract Input-Output Tables (IOT): Retrieve IOT matrices for five regions (NW, NE, Toscana, Centro, Sud).
Process Trade Data: Prepare trade matrices for intermediate and final demand, based on regional sectoral interactions.
Export to Excel: Generate an Excel file with all matrices for easy review and integration into other parts of the MUST4WATER project.
Files

1. MRIO.RData
Contains the core MRIO data, including:

Input-Output Tables (IOT)
Trade data for intermediate and final demand
2. all_matrices.xlsx
The output Excel file, which contains the following sheets:

Trade_Intermedio: Intermediate trade data matrix (215x5)
Trade_Finale: Final demand trade data matrix (215x5)
IOT_NW, IOT_NE, IOT_Toscana, IOT_Centro, IOT_Sud: IOT matrices (49x52) for each region
How to Run the Script

1. Prerequisites:
R installed on your computer
writexl package: Install by running:

install.packages("writexl")

2. Load MRIO Data:
Make sure you have the MRIO dataset in your working directory, and load it into R using:

load("~/Desktop/MRIO.RData")

3. Process the Data:
The script extracts and processes:

a) Input-Output Tables (IOT) for the five regions
b) Intermediate and Final Trade Data for regional sectors

4. Export Results:
The processed matrices are exported into an Excel file using:

write_xlsx(all_matrices, path = "all_matrices.xlsx")

This file contains all the matrices, with each region and trade dataset saved in separate sheets.

How to Use the Output

Economic Analysis: These matrices can be used to assess inter-regional economic flows and dependencies within the scope of the MUST4WATER project.
Integration into Broader Work: The Excel file can easily be shared with other team members for further modeling or analysis related to water and economic systems in the target regions.

Acknowledgments

This script is part of the MUST4WATER project. Contributions and further analysis will feed into broader research objectives aimed at understanding regional interdependencies and their impact on water resources.
