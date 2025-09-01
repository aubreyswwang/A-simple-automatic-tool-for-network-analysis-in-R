# A-simple-automatic-tool-for-network-analysis-in-R
Just edit one CSV file, run the provided R script, and the results will be automatically generated and saved into the specified result folder.

This tool performs **single network analysis**, including:
-	network estimation and visualization
-	centrality estimation and visualization
-	stability testing 
-	accuracy assessment

## ✨ Key Features

- **One-Click Analysis**: Modify the config file, no coding required  
- **Complete Workflow**: Estimation, visualization, centrality, stability, accuracy  
- **Flexible Configuration**: Custom labels, groups, colors, parameters  
- **Weighted Sample Support**: Built-in support for sample weights  
- **Automated Reports**: All results saved in a structured `results/` folder  

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

   ## 🛠️ Requirements
- **R ≥ 4.0**
- Required packages (auto-installed if missing):  
  `readxl, dplyr, haven, readr, tidyr, forcats, psych, car, Hmisc, Matrix, mgm, bootnet, qgraph, glmnet, networktools, huge, ggplot2, patchwork, gridExtra, writexl, openxlsx, officer, flextable`

  ## 📁 Project Structure
  ```
  network-analysis/
  ├── main_analysis_single_network.R # Main script
  ├── config_single_network.csv # Configuration file
  ├── data/ # Data folder
  │ └── your_data.xlsx
  ├── results/ # Output results
  │ ├── networks/
  │ ├── centrality/
  │ ├── stability/
  │ ├── accuracy/
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
   Edit `config_single_network.csv`.

4. **Run Analysis**
   source("main\_analysis\_single\_network.R")

5. **View Results**
   Results saved in `results/`: plots, CSVs, and a text report.

## ⚙️ Configuration

The config file (`config_single_network.csv`) uses three columns:

* **parameter**: parameter name
* **value**: parameter value
* **description**: explanation

Example:
| parameter           | value                                                 | description                                                     |
| ------------------- | ----------------------------------------------------- | --------------------------------------------------------------- |
| output\_folder      | /your/result/directory                                | Output folder for all results                                   |
| data\_file          | /your/data/directory/file                             | Data file name (supports .xlsx/.xls/.csv)                       |
| sheet\_name         | your sheet name                                       | Excel worksheet name (leave empty for CSV files)                |
| start\_col          | 2                                                     | Starting column number for network analysis                     |
| end\_col            | 12                                                    | Ending column number for network analysis                       |
| has\_weights        | TRUE                                                  | Whether to use weighted analysis (TRUE/FALSE)                   |
| weight\_col         | IPW\_weight                                           | Weight column name or number (leave empty if not using weights) |
| node\_labels        | A;B;C;D;E;F;G;H;I;J                                   | Node labels (separated by semicolons)                           |
| node\_fullnames     | aaa;bbb;ccc;ddd;eee;fff;ggg;hhh;iii;jjj               | Full node names (for grouping and legends)                      |
| group\_info         | "ABC:1,2,3\:red;DEFG:4,5,6,7\:blue;HIJ:8,9,10\:green" | Group information (format: GroupName\:NodeList\:Color)          |
| network\_title      | Single Network Analysis                               | Network plot title                                              |
| gamma\_value        | 0.5                                                   | EBIC gamma parameter (usually 0.25-0.5)                         |
| n\_boots\_stability | 1000                                                  | Bootstrap iterations for stability analysis                     |
| n\_boots\_accuracy  | 5000                                                  | Bootstrap iterations for accuracy analysis                      |

## 📊 Output Files

* **networks/**: network plots + legends
* **centrality/**: centrality plots + CSV
* **stability/**: bootstrap plots + CS coefficient
* **accuracy/**: confidence intervals + edge difference plots
* **predictability/**: predictability CSV
* **analysis\_report.txt**: summary report

## ⚠️ Notes

* Data should be continuous; missing values removed automatically
* Recommended: ≥ 3× nodes as sample size
* Bootstrap is time-intensive; adjust iterations accordingly

