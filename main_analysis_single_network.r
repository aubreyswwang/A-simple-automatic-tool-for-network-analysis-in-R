# main_analysis_single_network.R
# Automated Single Network Analysis Project
# Author: Aubrey WANG
# Purpose: Generate single network analysis reports by modifying only the config_single_network.csv file.
# =============================================================================

# Clear environment
rm(list = ls())

# Load required packages
cat("Loading R packages...\n")
pkgs <- c("readxl", "dplyr", "haven", "psych", "car",
          "ggplot2", "mgm", "bootnet", "qgraph",
          "glmnet", "readr", "networktools", "Hmisc",
          "writexl", "patchwork", "openxlsx", "huge",
          "officer", "flextable", "gridExtra", "tidyr",
          "forcats", "Matrix")

# Install and load all packages
invisible(lapply(pkgs, function(pkg) {
  if (!require(pkg, character.only = TRUE)) {
    install.packages(pkg, dependencies = TRUE)
    library(pkg, character.only = TRUE)
  }
}))

# Manually specify working directory
# setwd("/your/project/directory")  # Please replace with your project path
cat("Current working directory:\n", getwd(), "\n") # Please confirm your working directory is correct

# Read configuration file
cat("Reading configuration file...\n")
if (!file.exists("config_single_network.csv")) {
  stop("Error: Configuration file 'config_single_network.csv' not found. Please create configuration file first!")
}

config <- read.csv("config_single_network.csv", stringsAsFactors = FALSE, encoding = "UTF-8")
config_list <- setNames(config$value, config$parameter)

# Parse configuration
output_path <- config_list[["output_folder"]]
data_file <- config_list[["data_file"]]
sheet_name <- config_list[["sheet_name"]]
start_col <- as.numeric(config_list[["start_col"]])
end_col <- as.numeric(config_list[["end_col"]])
has_weights <- as.logical(config_list[["has_weights"]])
weight_col <- if(has_weights && !is.na(config_list[["weight_col"]])) config_list[["weight_col"]] else NULL
node_labels <- strsplit(config_list[["node_labels"]], ";")[[1]]
node_fullnames <- if(!is.na(config_list[["node_fullnames"]]) && config_list[["node_fullnames"]] != "") {
  strsplit(config_list[["node_fullnames"]], ";")[[1]]
} else {
  node_labels
}
group_info <- config_list[["group_info"]]
network_title <- config_list[["network_title"]]
gamma_value <- as.numeric(config_list[["gamma_value"]])
n_boots_stability <- as.numeric(config_list[["n_boots_stability"]])
n_boots_accuracy <- as.numeric(config_list[["n_boots_accuracy"]])

# Define subfolder names and corresponding abbreviation suffixes
sub_folders <- c(
  "networks"     = "net",
  "centrality"   = "cen",
  "stability"    = "sta",
  "accuracy"     = "acc",
  "predictability" = "pre"
)

# Create folders and dynamically generate save_path_xxx variables
for (folder_name in names(sub_folders)) {
  full_path <- file.path(output_path, folder_name)
  
  # Ensure folder exists
  if (!dir.exists(full_path)) {
    dir.create(full_path, recursive = TRUE)
    cat("Created directory:", full_path, "\n")
  }

  var_name <- paste0("save_path_", sub_folders[folder_name]) # Generate variable name, e.g., save_path_net
  assign(var_name, full_path, envir = .GlobalEnv) # Assign path to variable (in global environment)
  cat("✅", var_name, "=", full_path, "\n")
}

# Read data
cat("Reading data file:", data_file, "\n")
if (!file.exists(data_file)) {
  data_file <- file.path("data", data_file)
  if (!file.exists(data_file)) {
    stop("Error: Data file not found!")
  }
}

# Choose reading method based on file extension
if (grepl("\\.xlsx?$", data_file)) {
  if (!is.na(sheet_name) && sheet_name != "") {
    raw_data <- read_excel(data_file, sheet = sheet_name)
  } else {
    raw_data <- read_excel(data_file)
  }
} else if (grepl("\\.csv$", data_file)) {
  raw_data <- read.csv(data_file)
} else {
  stop("Error: Unsupported file format! Only .xlsx, .xls, .csv are supported")
}

# Data preprocessing
cat("Processing data...\n")

# Extract network analysis data
network_data <- raw_data[, start_col:end_col, drop = FALSE]

# Extract weights (if available)
if (has_weights && !is.null(weight_col)) {
  if (is.numeric(weight_col)) {
    weights <- raw_data[, weight_col, drop = FALSE][[1]]
  } else {
    weights <- raw_data[[weight_col]]
  }
  if (any(is.na(weights))) {
    cat("Warning: Missing values in weight column, using equal weights\n")
    weights <- rep(1, nrow(network_data))
  }
} else {
  weights <- rep(1, nrow(network_data))
}

# Set node labels
if (length(node_labels) == ncol(network_data)) {
  colnames(network_data) <- node_labels
} else {
  cat("Warning: Number of node labels doesn't match number of variables, using default labels\n")
  colnames(network_data) <- paste0("Node_", 1:ncol(network_data))
}

# Data cleaning
network_data <- network_data %>%
  mutate(across(everything(), ~ as.numeric(as.character(.)))) %>%
  na.omit()

# Synchronously remove corresponding rows in weights
if (nrow(network_data) < nrow(raw_data)) {
  valid_rows <- as.numeric(rownames(network_data))
  weights <- weights[valid_rows]
}

cat("Data dimensions:", nrow(network_data), "rows", ncol(network_data), "columns\n")

# Parse group information
groups_vec <- NULL
group_colors <- NULL
node_colors <- NULL

if (length(node_labels) != length(node_fullnames)) {
  stop("❌ node_labels and node_fullnames counts don't match!")
}

# Create two mappings:
label_to_fullname <- setNames(node_fullnames, trimws(node_labels)) # abbreviation → full name (for legend display)
fullname_to_label <- setNames(node_labels, trimws(node_fullnames)) # full name → abbreviation (for parsing group_info)

# Create results folder (save_path_net already defined)
save_path_net <- file.path(config_list[["output_folder"]], "networks")
if (!dir.exists(save_path_net)) {
  dir.create(save_path_net, recursive = TRUE)
}

groups <- list()
group_colors <- character()
names(group_colors) <- character()

if (!is.na(config_list[["group_info"]]) && config_list[["group_info"]] != "") {
  group_parts <- strsplit(config_list[["group_info"]], ";")[[1]]
  
  for (part in group_parts) {
    part <- trimws(part)
    if (part == "") next
    
    # Regex extract: group_name:node_list:color
    match_pattern <- regmatches(part, regexec("^(.+?):([^:]+):(#(?:[0-9A-Fa-f]{6}))$", part))
    if (length(match_pattern[[1]]) < 4) {
      warning("Group format error, skipping: ", part)
      next
    }
    
    group_name <- trimws(match_pattern[[1]][2])
    index_str  <- trimws(match_pattern[[1]][3])
    color      <- match_pattern[[1]][4]
    
    items <- trimws(unlist(strsplit(index_str, ","))) # Parse nodes: supports numbers, abbreviations, full names
    # Determine type and convert to column indices
    if (grepl("^[0-9]+$", items[1])) {
      node_indices <- as.numeric(items) # Numeric indices
    } else if (items[1] %in% node_labels) {
      node_indices <- match(items, node_labels) # Abbreviation indices
    } else {
      # Full name → map to abbreviation → indices
      # Correct: use full name to find abbreviation
      short_names <- fullname_to_label[items]
      if (any(is.na(short_names))) {
        not_found <- items[is.na(short_names)]
        stop("Could not find labels corresponding to these full names:", paste(not_found, collapse = ", "))
      }
      node_indices <- match(short_names, node_labels)
    }
    # Check bounds
    if (any(is.na(node_indices)) || any(node_indices < 1) || any(node_indices > length(node_labels))) {
      stop("Group '", group_name, "' contains invalid nodes")
    }
    
    # Store as node labels, not numbers
    groups[[group_name]] <- node_labels[node_indices]
    group_colors[group_name] <- color
  }
} else {
  cat("Note: No group information provided\n")
}
  # After the for loop ends, construct groups_vec and node_colors
  if (length(groups) > 0) {
    groups_vec <- rep(NA, ncol(network_data))
    names(groups_vec) <- colnames(network_data)
    
    for (grp_name in names(groups)) {
      matched <- match(groups[[grp_name]], colnames(network_data))
      if (any(!is.na(matched))) {
        groups_vec[matched] <- grp_name
      }
    }
    
    node_colors <- group_colors[groups_vec]
} else {
  cat("Note: No group information provided\n")
}

# Data standardization (Nonparanormal transformation)
cat("Performing data transformation...\n")
network_data_npn <- huge.npn(network_data)

# Network estimation
cat("Estimating network...\n")

# Calculate weighted covariance and correlation matrices
if (has_weights && !is.null(weight_col)) {
  cov_matrix <- cov.wt(network_data_npn, wt = weights, cor = FALSE)$cov
  cor_matrix <- cov2cor(cov_matrix)
  effective_n <- sum(weights)^2 / sum(weights^2)  # Effective sample size
} else {
  cor_matrix <- cor(network_data_npn)
  effective_n <- nrow(network_data_npn)
}

# Estimate network using EBICglasso
network <- qgraph::EBICglasso(cor_matrix, n = effective_n, gamma = gamma_value)

# Node predictability analysis
cat("Computing node predictability...\n")
p <- ncol(network_data_npn)
mgm_fit <- mgm(
  data = network_data_npn,
  type = rep('g', p),
  level = rep(1, p),
  lambdaSel = "CV",
  ruleReg = "OR"
)
predictability <- predict(
  object = mgm_fit, 
  data = network_data_npn,
  errorCon = "R2"
)

mean_predictability <- round(mean(predictability$error$R2), digits = 3)

# Save predictability results
pred_results <- data.frame(
  Node = colnames(network_data_npn),
  R2 = round(predictability$error$R2, 3)
)
pred_results$Mean_R2 <- mean_predictability

write.csv(pred_results, file = file.path(save_path_pre, "predictability_results.csv"), row.names = FALSE)

# Network visualization
cat("Generating network plot...\n")

# Basic network plot
png(filename = file.path(save_path_net, "network_plot.png"),
 width = 8, height = 8, units = "in", res = 300)
par(mar = c(2, 2, 3, 2))

if (!is.null(groups_vec)) {
  network_plot <- qgraph(
    cor_matrix,
    graph = "glasso",
    gamma = gamma_value,
    layout = "spring",
    groups = groups_vec,
    sampleSize = effective_n,
    labels = colnames(network_data_npn),
    color = node_colors,
    pie = as.numeric(predictability$error$R2),
    pieColor = rep('grey', ncol(network_data_npn)),
    border.color = "black",
    border.width = 1,
    theme = "Borkulo",
    legend = FALSE,
    vsize = 8.5,
    edge.labels = TRUE,
    edge.label.cex = 0.7,
    fade = TRUE,
    title = network_title
  )
} else {
  network_plot <- qgraph(
    cor_matrix,
    graph = "glasso",
    gamma = gamma_value,
    layout = "spring",
    sampleSize = effective_n,
    labels = colnames(network_data_npn),
    pie = as.numeric(predictability$error$R2),
    pieColor = rep('grey', ncol(network_data_npn)),
    border.color = "black",
    border.width = 1,
    theme = "Borkulo",
    legend = FALSE,
    vsize = 8.5,
    edge.labels = TRUE,
    edge.label.cex = 0.7,
    fade = TRUE,
    title = network_title
  )
}

dev.off()

# Generate corresponding custom legend
save_single_group_legend <- function(group_name, items, group_color, item_fullnames_map, save_path, filename, width = 4.5, res = 300) {
  n_items <- length(items)
  item_height     <- 0.35
  title_gap       <- 0.30
  bottom_padding  <- 0.15
  top_padding     <- 0.15
  text_y_offset   <- 0.01
  group_name_cex  <- 1.6
  item_text_cex   <- 1.3
  bar_width       <- 3.3 # Change here to adjust bar width
  
  total_height <- top_padding + 0.25 + title_gap + n_items * item_height + bottom_padding
  total_width  <- width
  
  png(file.path(save_path, filename), total_width, total_height, "in", res = res)
  par(mar = c(0.03, 0.03, 0, 0.03), xpd = NA)
  
  plot(1, type = "n", axes = FALSE, xlab = "", ylab = "", xlim = c(0, 4), ylim = c(0, total_height))
  
  # Group name
  y_pos <- total_height - top_padding - 0.25
  text(0.05, y_pos, group_name, font = 2, cex = group_name_cex, adj = c(0, 0.5))
  y_pos <- y_pos - title_gap
  
  # Each node: display "full_name (abbreviation)"
  for (item in items) {
    full_name <- item_fullnames_map[item]
    if (is.na(full_name)) full_name <- item
    label_text <- paste0(full_name, " (", item, ")")
    
    rect(0.015, y_pos - item_height, 0.05 + bar_width, y_pos, col = group_color, border = "black")
    text(0.05, y_pos - item_height / 2 + text_y_offset,
         label_text,
         adj = c(0, 0.5), cex = item_text_cex)
    
    y_pos <- y_pos - item_height
  }
  
  dev.off()
  cat("Legend saved:", file.path(save_path, filename), "\n")
}

# Generate legend for each group
if (length(groups) > 0) {
  for (grp_name in names(groups)) {
    save_single_group_legend(
      group_name = grp_name,
      items = groups[[grp_name]],         
      group_color = group_colors[grp_name],
      item_fullnames_map = label_to_fullname, 
      save_path = save_path_net,
      filename = paste0("legend_", gsub("[[:space:]]+", "_", grp_name), ".png"),
      width = 5.0,
      res = 300
    )
  }
} else {
  cat("No legend generated: no group information\n")
}

# Centrality analysis
cat("Computing centrality metrics...\n")
png(filename = file.path(save_path_cen, "centrality_plot.png"),
  width = 8, height = 8, units = "in", res = 300)

centrality_results <- centralityPlot(
  network_plot,
  weighted = TRUE,
  signed = TRUE,
  scale = "z-scores",
  labels = colnames(network_data_npn),
  include = "all",
  orderBy = "ExpectedInfluence"
)

dev.off()

# Save centrality values
write.csv(centrality_results$data, file = file.path(save_path_cen, "centrality_values.csv"), row.names = FALSE)

# Stability analysis
cat("Performing stability analysis...\n")

# Create custom estimation function
myWeightedGlasso_bootnet <- function(data, gamma, ...) {
  if (has_weights && !is.null(weight_col)) {
    # Match row indices to get corresponding weights
    data_str <- apply(data, 1, paste, collapse = "_")
    orig_str <- apply(network_data_npn, 1, paste, collapse = "_")
    row_ids <- match(data_str, orig_str)
    
    if (any(is.na(row_ids))) {
      stop("Unable to match weights")
    }
    
    curr_weights <- weights[row_ids]
    covm <- cov.wt(data, wt = curr_weights, cor = FALSE)$cov
    C <- cov2cor(covm)
    C_pd <- as.matrix(nearPD(C, corr = TRUE)$mat)
    net <- EBICglasso(C_pd, n = length(curr_weights), gamma = gamma)
  } else {
    C <- cor(data)
    net <- EBICglasso(C, n = nrow(data), gamma = gamma)
  }
  return(net)
}

# Case-dropping bootstrap for stability
set.seed(1234)
casedrop_boot <- bootnet(
  data = as.matrix(network_data_npn),
  default = "none",
  fun = myWeightedGlasso_bootnet,
  type = "case",
  statistics = c("edge", "expectedInfluence", "strength", "closeness", "betweenness"),
  nBoots = n_boots_stability,
  gamma = gamma_value,
  caseN = nrow(network_data_npn),
  maxErrors = 10,
  errorHandling = "continue",
  nCores = 1
)

# Plot stability graphs
# Expected Influence stability plot
stability_plot <- plot(casedrop_boot, statistics = "expectedInfluence", order = "sample") +
  theme(legend.position = "none") +
  ggtitle("Expected Influence Stability")

ggsave(
  filename = file.path(save_path_sta, "stability_plot.png"),
  plot = stability_plot,
  width = 8, height = 8, units = "in", dpi = 300
)

# Strength stability plot
strength_plot <- plot(casedrop_boot, statistics = "strength", order = "sample") +
  theme(legend.position = "none") +
  ggtitle("Strength Stability")

ggsave(
  filename = file.path(save_path_sta, "strength_stability_plot.png"),
  plot = strength_plot,
  width = 8, height = 8, units = "in", dpi = 300
)

# Betweenness stability plot
betweenness_plot <- plot(casedrop_boot, statistics = "betweenness", order = "sample") +
  theme(legend.position = "none") +
  ggtitle("Betweenness Stability")

ggsave(
  filename = file.path(save_path_sta, "betweenness_stability_plot.png"),
  plot = betweenness_plot,
  width = 8, height = 8, units = "in", dpi = 300
)

# Closeness stability plot
closeness_plot <- plot(casedrop_boot, statistics = "closeness", order = "sample") +
  theme(legend.position = "none") +
  ggtitle("Closeness Stability")

ggsave(
  filename = file.path(save_path_sta, "closeness_stability_plot.png"),
  plot = closeness_plot,
  width = 8, height = 8, units = "in", dpi = 300
)

# CS coefficient calculation
cs_results <- capture.output(corStability(casedrop_boot, cor = 0.7, statistics = "expectedInfluence"))
writeLines(cs_results, file.path(save_path_sta, "cs_coefficient.txt"))

# Accuracy analysis
cat("Performing accuracy analysis...\n")

# Nonparametric bootstrap for accuracy
set.seed(1234)
nonpar_boot <- bootnet(
  data = as.matrix(network_data_npn),
  default = "none",
  fun = myWeightedGlasso_bootnet,
  nBoots = n_boots_accuracy,
  gamma = gamma_value,
  nCores = 1,
  type = "nonparametric",
  maxErrors = 10,
  errorHandling = "continue",
  weighted = TRUE,
  signed = TRUE
)

# Confidence interval plot
accuracy_plot <- plot(nonpar_boot, order = "sample", plot = "area")
ggsave(
  filename = file.path(save_path_acc, "accuracy_plot.png"),
  plot = accuracy_plot,
  width = 6, height = 8, units = "in", dpi = 300
)

# Edge weight difference plot
edge_diff_plot <- plot(nonpar_boot, "edge", plot = "difference", onlyNonZero = TRUE, order = "sample")
ggsave(
  filename = file.path(save_path_acc, "edge_difference_plot.png"),
  plot = edge_diff_plot,
  width = 8, height = 8, units = "in", dpi = 300
)

# Generate comprehensive report
cat("Generating analysis report...\n")

report_content <- paste0(
  "Network Analysis Report\n",
  "=====================================\n\n",
  "Analysis Time: ", Sys.time(), "\n",
  "Data File: ", data_file, "\n",
  "Network Title: ", network_title, "\n\n",
  "Data Overview:\n",
  "- Sample Size: ", nrow(network_data_npn), "\n",
  "- Number of Nodes: ", ncol(network_data_npn), "\n",
  "- Using Weights: ", ifelse(has_weights && !is.null(weight_col), "Yes", "No"), "\n",
  "- Effective Sample Size: ", round(effective_n, 2), "\n\n",
  "Network Parameters:\n",
  "- EBIC gamma: ", gamma_value, "\n",
  "- Stability Analysis Bootstrap Iterations: ", n_boots_stability, "\n",
  "- Accuracy Analysis Bootstrap Iterations: ", n_boots_accuracy, "\n\n",
  "Output Files Description:\n",
  "- results/networks/: Network plots\n",
  "- results/centrality/: Centrality analysis results\n",
  "- results/stability/: Stability analysis results\n", 
  "- results/accuracy/: Accuracy analysis results\n",
  "- results/predictability/: Node predictability results\n"
)

writeLines(report_content, "results/analysis_report.txt")

# Completion
cat("\n=====================================\n")
cat("Network analysis completed!\n")
cat("All results have been saved to the results/ folder\n")
cat("Please check results/analysis_report.txt for detailed report\n")
cat("=====================================\n")