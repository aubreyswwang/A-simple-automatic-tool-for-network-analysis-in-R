# A_simple_automatic_tool_for_network_analysis_in_R
Simply edit the CSV file, run the provided R script, and the results will be automatically generated and stored in the specified results folder.

This tool enables network analysis for two groups. For each network, the results include：
- network estimation and visualization
- centrality estimation and visualization
- stability testing
- accuracy assessment

A network comparison test report between the two groups is also generated.

## ✨ Key Features

- **Weighted Data Ready**: Supports imputed datasets with sample weights or without, extending applicability beyond traditional methods.
- **Flexible Network Analysis**: Handles one or two networks; includes full comparison tests when two networks are analyzed.
- **End-to-End Workflow**: From estimation and visualization to stability, accuracy, and automated reporting.
- **Easy to Use**: One config file, no coding required.

## 📋 Functional Modules
1. **Network Estimation & Visualization**  
   EBICglasso estimation, customizable grouping/colors, high-resolution plots, node predictability pie charts  
2. **Centrality Analysis**  
   Strength, Expected Influence, Betweenness, Closeness; plots + CSV export  
3. **Stability Analysis**  
   Case-dropping bootstrap, CS coefficient, multiple metric stability plots  
4. **Accuracy Analysis**  
   Nonparametric bootstrap CIs, edge difference tests, accuracy assessment  
5. **Node Predictability**  
   MGM-based predictability, R² visualization, CSV export
6. **Network Comparison**
   Statistical comparison between networks using NCT

   ## 🛠️ Requirements
- **R ≥ 4.0**
- Required packages (auto-installed if missing):  
  `"readxl", "dplyr", "haven", "psych", "car",
  "ggplot2", "mgm", "bootnet", "qgraph",
  "glmnet", "readr", "networktools", "Hmisc",
  "writexl", "patchwork", "openxlsx", "huge",
  "officer", "flextable", "gridExtra", "tidyr",
  "forcats", "Matrix", "NetworkComparisonTest"`

  ## 📁 Project Structure
  ```
  network-analysis/
  ├── main_analysis_dual_network.R # Main script
  ├── config_dual_network.csv # Configuration file
  ├── data/ # Data folder
  │ └── your_data.xlsx
  ├── results/ # Output results
  │ ├── networks/
  │ ├── centrality/
  │ ├── stability/
  │ ├── accuracy/
  │ ├── comparison/
  │ └── predictability/
  └── README.md
  ```

  ## 🚀 Quick Start
1. **Clone Repository**

2. **Prepare Data**
   Place .xlsx, .xls, or .csv file in `data/`.

   * Rows = participants
   * Columns = variables
   * Optional: weight column

3. **Configure Parameters**
   Edit `config_dual_network.csv`.

4. **Run Analysis**
   source("main\_analysis\_dual\_network.R")

5. **View Results**
   Results saved in `results/`: plots, CSVs, and a text report.
   ```
   results/
   ├── networks/              # Network visualizations and legends
   │   ├── network_plot_A.png
   │   ├── network_plot_B.png (dual network only)
   │   ├── legend_A_*.png
   │   └── legend_B_*.png (dual network only)
   ├── centrality/            # Centrality analysis results
   │   ├── centrality_plot_A.png
   │   ├── centrality_values_A.csv
   │   └── ... (corresponding B files for dual network)
   ├── stability/             # Stability analysis results
   │   ├── stability_plot_A.png
   │   ├── strength_stability_plot_A.png
   │   ├── cs_coefficient_A.txt
   │   └── ... (corresponding B files for dual network)
   ├── accuracy/              # Accuracy analysis results
   │   ├── accuracy_plot_A.png
   │   ├── edge_difference_plot_A.png
   │   └── ... (corresponding B files for dual network)
   ├── predictability/        # Node predictability results
   │   ├── predictability_results_A.csv
   │   └── predictability_results_B.csv (dual network only)
   ├── comparison/            # Network comparison results (dual network only)
   │   ├── nct_result.rds
   │   ├── nct_summary_output_raw.xlsx
   │   ├── nct_plot_global.png
   │   ├── nct_plot_edges.png
   │   └── significant_edge_differences.csv
   └── analysis_report.txt    # Comprehensive analysis report
   ```

## ⚙️ Configuration

The config file (`config_dual_network.csv`) uses three columns:

* **parameter**: parameter name
* **value**: parameter value
* **description**: explanation

Example:
| parameter         | value                                                   | description                                                   |
|-------------------|---------------------------------------------------------|---------------------------------------------------------------|
| output_folder     | /your/result/directory                                  | Output folder for all results                                 |
| data_file         | /your/data/directory/file                               | Data file name (supports .xlsx/.xls/.csv)                     |
| sheet_name        | your sheet name                                         | Excel worksheet name (leave empty for CSV files)              |
| has_weights       | TRUE                                                    | Whether to use weighted analysis (TRUE/FALSE)                 |
| weight_col        | IPW_weight                                              | Weight column name or number (leave empty if not using weights)|
| start_col_A       | 2                                                       | Starting column number for network A                          |
| end_col_A         | 11                                                      | Ending column number for network A                            |
| node_labels_A     | A;B;C;D;E;F;G;H;I;J                                     | Node labels for Network A (separated by semicolons)           |
| node_fullnames_A  | aaa;bbb;ccc;ddd;eee;fff;ggg;hhh;iii;jjj                 | Full node names for A                                         |
| group_info_A      | "ABC:1,2,3:red;DEFG:4,5,6,7:blue;HIJ:8,9,10:green"      | Group info for A: GroupName:NodeList:Color                    |
| network_title_A   | Network A: Control Group                                | Title for Network A                                           |
| start_col_B       | 20                                                      | Starting column number for network B                          |
| end_col_B         | 29                                                      | Ending column number for network B                            |
| node_labels_B     | A;B;C;D;E;F;G;H;I;J                                     | Node labels for Network B (must match A!)                     |
| node_fullnames_B  | aaa;bbb;ccc;ddd;eee;fff;ggg;hhh;iii;jjj                 | Full node names for B                                         |
| group_info_B      | "ABC:1,2,3:red;DEFG:4,5,6,7:blue;HIJ:8,9,10:green"      | Group info for B                                              |
| network_title_B   | Network B: Patient Group                                | Title for Network B                                           |
| gamma_value       | 0.5                                                     | EBIC gamma parameter (usually 0.25-0.5)                       |
| n_boots_stability | 1000                                                    | Bootstrap iterations for stability analysis                   |
| n_boots_accuracy  | 5000                                                    | Bootstrap iterations for accuracy analysis                    |
| n_boots_compare   | 1000                                                    | Bootstrap iterations for network comparison test              |


## ⚠️ Notes

* This script is designed for datasets where Groups A and B share the same weight column.
* Recommended: ≥ 3× nodes as sample size
* Bootstrap is time-intensive; adjust iterations accordingly

