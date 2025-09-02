# main_analysis_dual_network.R
# Automated Dual Network Analysis Project
# Author: Aubrey WANG
# Purpose: Generate dual network analysis reports by modifying only the config_dual_network.csv file.
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
          "forcats", "Matrix", "NetworkComparisonTest")

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
if (!file.exists("config_dual_network.csv")) {
  stop("Error: Configuration file 'config_dual_network.csv' not found. Please create configuration file first!")
}

config <- read.csv("config_dual_network.csv", stringsAsFactors = FALSE, encoding = "UTF-8")
config_list <- setNames(config$value, config$parameter)

# Parse configuration
output_path <- config_list[["output_folder"]]
data_file <- config_list[["data_file"]]
sheet_name <- config_list[["sheet_name"]]
has_weights <- as.logical(config_list[["has_weights"]])
weight_col <- if(has_weights && !is.na(config_list[["weight_col"]])) config_list[["weight_col"]] else NULL

# Network A
start_col_A <- as.numeric(config_list[["start_col_A"]])
end_col_A <- as.numeric(config_list[["end_col_A"]])
node_labels_A <- strsplit(config_list[["node_labels_A"]], ";")[[1]]
node_fullnames_A <- if(!is.na(config_list[["node_fullnames_A"]]) && config_list[["node_fullnames_A"]] != "") {
  strsplit(config_list[["node_fullnames_A"]], ";")[[1]]
} else {node_labels_A}
group_info_A <- config_list[["group_info_A"]]
network_title_A <- config_list[["network_title_A"]]

# Network B
start_col_B <- if(!is.na(config_list[["start_col_B"]]) && config_list[["start_col_B"]] != "") as.numeric(config_list[["start_col_B"]]) else NULL
end_col_B <- if(!is.na(config_list[["end_col_B"]]) && config_list[["end_col_B"]] != "") as.numeric(config_list[["end_col_B"]]) else NULL
node_labels_B <- if(!is.na(config_list[["node_labels_B"]]) && config_list[["node_labels_B"]] != "") {
  strsplit(config_list[["node_labels_B"]], ";")[[1]]
} else {NULL}
node_fullnames_B <- if(!is.na(config_list[["node_fullnames_B"]]) && config_list[["node_fullnames_B"]] != "") {
  strsplit(config_list[["node_fullnames_B"]], ";")[[1]]
} else {node_labels_B}
group_info_B <- if(!is.na(config_list[["group_info_B"]]) && config_list[["group_info_B"]] != "") config_list[["group_info_B"]] else NULL
network_title_B <- if(!is.na(config_list[["network_title_B"]]) && config_list[["network_title_B"]] != "") config_list[["network_title_B"]] else NULL

# Common parameters
gamma_value <- as.numeric(config_list[["gamma_value"]])
n_boots_stability <- as.numeric(config_list[["n_boots_stability"]])
n_boots_accuracy <- as.numeric(config_list[["n_boots_accuracy"]])
n_boots_compare <- as.numeric(config_list[["n_boots_compare"]])

# Define subfolder names and corresponding abbreviation suffixes
sub_folders <- c(
  "networks"       = "net",
  "centrality"     = "cen",
  "stability"      = "sta",
  "accuracy"       = "acc",
  "predictability" = "pre",
  "comparison"     = "cmp" 
)

# Create folders and dynamically generate save_path_xxx variables
for (folder_name in names(sub_folders)) {
  full_path <- file.path(output_path, folder_name)
  
  # Ensure folder exists
  if (!dir.exists(full_path)) {
    dir.create(full_path, recursive = TRUE)
  }

  var_name <- paste0("save_path_", sub_folders[folder_name]) # Generate variable name, e.g., save_path_net
  assign(var_name, full_path, envir = .GlobalEnv) # Assign path to variable (in global environment)
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

# Data preprocessing for Network A
cat("Processing Network A data...\n")
network_data_A <- raw_data[, start_col_A:end_col_A, drop = FALSE]

# Extract weights for Network A (if available)
if (has_weights && !is.null(weight_col)) {
  if (is.numeric(weight_col)) {
    weights <- raw_data[, weight_col, drop = FALSE][[1]]
  } else {
    weights <- raw_data[[weight_col]]
  }
  if (any(is.na(weights))) {
    cat("Warning: Missing values in weight column for Network A, using equal weights\n")
    weights <- rep(1, nrow(network_data_A))
  }
} else {
  weights <- rep(1, nrow(network_data_A))
}

# Set node labels for Network A
if (length(node_labels_A) == ncol(network_data_A)) {
  colnames(network_data_A) <- node_labels_A
} else {
  cat("Warning: Number of node labels doesn't match number of variables for Network A, using default labels\n")
  colnames(network_data_A) <- paste0("Node_A", 1:ncol(network_data_A))
}

# Data cleaning for Network A
network_data_A <- network_data_A %>%
  mutate(across(everything(), ~ as.numeric(as.character(.)))) %>%
  na.omit()

# Synchronously remove corresponding rows in weights for Network A
if (nrow(network_data_A) < nrow(raw_data)) {
  valid_rows_A <- as.numeric(rownames(network_data_A))
  weights <- weights[valid_rows_A]
}

cat("Network A data dimensions:", nrow(network_data_A), "rows", ncol(network_data_A), "columns\n")

# Data preprocessing for Network B (if exists)
network_data_B <- NULL

if (!is.null(start_col_B) && !is.null(end_col_B)) {
  cat("Processing Network B data...\n")
  network_data_B <- raw_data[, start_col_B:end_col_B, drop = FALSE]
    
  # Set node labels for Network B
  if (!is.null(node_labels_B) && length(node_labels_B) == ncol(network_data_B)) {
    colnames(network_data_B) <- node_labels_B
  } else {
    cat("Warning: Number of node labels doesn't match number of variables for Network B, using default labels\n")
    colnames(network_data_B) <- paste0("Node_B", 1:ncol(network_data_B))
  }
  
  # Data cleaning for Network B
  network_data_B <- network_data_B %>%
    mutate(across(everything(), ~ as.numeric(as.character(.)))) %>%
    na.omit()
  
  # Synchronously remove corresponding rows in weights for Network B
  if (nrow(network_data_B) < nrow(raw_data)) {
    valid_rows_B <- as.numeric(rownames(network_data_B))
    weights <- weights[valid_rows_B]
  }
  
  cat("Network B data dimensions:", nrow(network_data_B), "rows", ncol(network_data_B), "columns\n")
}

# Function to parse group information
parse_group_info <- function(group_info_str, node_labels, node_fullnames) {
  groups <- list()
  group_colors <- character()
  
  if (is.na(group_info_str) || group_info_str == "") {
    return(list(groups = groups, group_colors = group_colors))
  }
  
  # Create mappings
  label_to_fullname <- setNames(node_fullnames, trimws(node_labels))
  fullname_to_label <- setNames(node_labels, trimws(node_fullnames))
  
  group_parts <- strsplit(group_info_str, ";")[[1]]
  
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
    
    items <- trimws(unlist(strsplit(index_str, ",")))
    
    # Determine type and convert to column indices
    if (grepl("^[0-9]+$", items[1])) {
      node_indices <- as.numeric(items) # Numeric indices
    } else if (items[1] %in% node_labels) {
      node_indices <- match(items, node_labels) # Abbreviation indices
    } else {
      # Full name → map to abbreviation → indices
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
  
  return(list(groups = groups, group_colors = group_colors))
}

# Parse group information for Network A
group_result_A <- parse_group_info(group_info_A, node_labels_A, node_fullnames_A)
groups_A <- group_result_A$groups
group_colors_A <- group_result_A$group_colors

# Create groups_vec and node_colors for Network A
groups_vec_A <- NULL
node_colors_A <- NULL
if (length(groups_A) > 0) {
  groups_vec_A <- rep(NA, ncol(network_data_A))
  names(groups_vec_A) <- colnames(network_data_A)
  
  for (grp_name in names(groups_A)) {
    matched <- match(groups_A[[grp_name]], colnames(network_data_A))
    if (any(!is.na(matched))) {
      groups_vec_A[matched] <- grp_name
    }
  }
  
  node_colors_A <- group_colors_A[groups_vec_A]
}

# Parse group information for Network B (if exists)
groups_B <- list()
group_colors_B <- character()
groups_vec_B <- NULL
node_colors_B <- NULL

if (!is.null(network_data_B) && !is.null(group_info_B)) {
  group_result_B <- parse_group_info(group_info_B, node_labels_B, node_fullnames_B)
  groups_B <- group_result_B$groups
  group_colors_B <- group_result_B$group_colors
  
  if (length(groups_B) > 0) {
    groups_vec_B <- rep(NA, ncol(network_data_B))
    names(groups_vec_B) <- colnames(network_data_B)
    
    for (grp_name in names(groups_B)) {
      matched <- match(groups_B[[grp_name]], colnames(network_data_B))
      if (any(!is.na(matched))) {
        groups_vec_B[matched] <- grp_name
      }
    }
    
    node_colors_B <- group_colors_B[groups_vec_B]
  }
}

# Data standardization (Nonparanormal transformation)
cat("Performing data transformation...\n")
npn_A <- huge.npn(network_data_A)

npn_B <- NULL
if (!is.null(network_data_B)) {
  npn_B <- huge.npn(network_data_B)
}

# Network estimation for Network A
cat("Estimating Network A...\n")

# Calculate weighted covariance and correlation matrices for Network A
if (has_weights && !is.null(weight_col)) {
  cov_matrix_A <- cov.wt(npn_A, wt = weights, cor = FALSE)$cov
  cor_matrix_A <- cov2cor(cov_matrix_A)
  effective_n_A <- sum(weights)^2 / sum(weights^2)  # Effective sample size
} else {
  cor_matrix_A <- cor(npn_A)
  effective_n_A <- nrow(npn_A)
}

# Estimate network using EBICglasso for Network A
network_A <- qgraph::EBICglasso(cor_matrix_A, n = effective_n_A, gamma = gamma_value)

# Network estimation for Network B (if exists)
network_B <- NULL
cor_matrix_B <- NULL
effective_n_B <- NULL

if (!is.null(network_data_B)) {
  cat("Estimating Network B...\n")
  
  # Calculate weighted covariance and correlation matrices for Network B
  if (has_weights && !is.null(weight_col)) {
    cov_matrix_B <- cov.wt(npn_B, wt = weights, cor = FALSE)$cov
    cor_matrix_B <- cov2cor(cov_matrix_B)
    effective_n_B <- sum(weights)^2 / sum(weights^2)  # Effective sample size
  } else {
    cor_matrix_B <- cor(npn_B)
    effective_n_B <- nrow(npn_B)
  }
  
  # Estimate network using EBICglasso for Network B
  network_B <- qgraph::EBICglasso(cor_matrix_B, n = effective_n_B, gamma = gamma_value)
}

# Node predictability analysis for Network A
cat("Computing node predictability for Network A...\n")
p_A <- ncol(npn_A)
mgm_fit_A <- mgm(
  data = npn_A,
  type = rep('g', p_A),
  level = rep(1, p_A),
  lambdaSel = "CV",
  ruleReg = "OR"
)
predictability_A <- predict(
  object = mgm_fit_A, 
  data = npn_A,
  errorCon = "R2"
)

mean_predictability_A <- round(mean(predictability_A$error$R2), digits = 3)

# Save predictability results for Network A
pred_results_A <- data.frame(
  Node = colnames(npn_A),
  R2 = round(predictability_A$error$R2, 3)
)
pred_results_A$Mean_R2 <- mean_predictability_A

write.csv(pred_results_A, file = file.path(save_path_pre, "predictability_results_A.csv"), row.names = FALSE)

# Node predictability analysis for Network B (if exists)
predictability_B <- NULL
mean_predictability_B <- NULL

if (!is.null(network_data_B)) {
  cat("Computing node predictability for Network B...\n")
  p_B <- ncol(npn_B)
  mgm_fit_B <- mgm(
    data = npn_B,
    type = rep('g', p_B),
    level = rep(1, p_B),
    lambdaSel = "CV",
    ruleReg = "OR"
  )
  predictability_B <- predict(
    object = mgm_fit_B, 
    data = npn_B,
    errorCon = "R2"
  )
  
  mean_predictability_B <- round(mean(predictability_B$error$R2), digits = 3)
  
  # Save predictability results for Network B
  pred_results_B <- data.frame(
    Node = colnames(npn_B),
    R2 = round(predictability_B$error$R2, 3)
  )
  pred_results_B$Mean_R2 <- mean_predictability_B
  
  write.csv(pred_results_B, file = file.path(save_path_pre, "predictability_results_B.csv"), row.names = FALSE)
}

# Network visualization function
plot_network <- function(cor_matrix, network_data, groups_vec, node_colors, predictability, effective_n, network_title, filename) {
  png(filename = filename, width = 8, height = 8, units = "in", res = 300)
  par(mar = c(2, 2, 3, 2))
  
  if (!is.null(groups_vec)) {
    network_plot <- qgraph(
      cor_matrix,
      graph = "glasso",
      gamma = gamma_value,
      layout = "spring",
      groups = groups_vec,
      sampleSize = effective_n,
      labels = colnames(network_data),
      color = node_colors,
      pie = as.numeric(predictability$error$R2),
      pieColor = rep('grey', ncol(network_data)),
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
      labels = colnames(network_data),
      pie = as.numeric(predictability$error$R2),
      pieColor = rep('grey', ncol(network_data)),
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
  return(network_plot)
}

# Network visualization
cat("Generating network plots...\n")

# Plot Network A
network_plot_A <- plot_network(
  cor_matrix_A, npn_A, groups_vec_A, node_colors_A, 
  predictability_A, effective_n_A, network_title_A,
  file.path(save_path_net, "network_plot_A.png")
)

# Plot Network B (if exists)
network_plot_B <- NULL
if (!is.null(network_data_B)) {
  network_plot_B <- plot_network(
    cor_matrix_B, npn_B, groups_vec_B, node_colors_B, 
    predictability_B, effective_n_B, network_title_B,
    file.path(save_path_net, "network_plot_B.png")
  )
}

# Generate legend function
save_single_group_legend <- function(group_name, items, group_color, item_fullnames_map, save_path, filename, width = 4.5, res = 300) {
  n_items <- length(items)
  item_height     <- 0.35
  title_gap       <- 0.30
  bottom_padding  <- 0.15
  top_padding     <- 0.15
  text_y_offset   <- 0.01
  group_name_cex  <- 1.6
  item_text_cex   <- 1.3
  bar_width       <- 3.3 # Change here if you want to adjust the bar width
  
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

# Generate legend for Network A groups
if (length(groups_A) > 0) {
  label_to_fullname_A <- setNames(node_fullnames_A, trimws(node_labels_A))
  for (grp_name in names(groups_A)) {
    save_single_group_legend(
      group_name = grp_name,
      items = groups_A[[grp_name]],         
      group_color = group_colors_A[grp_name],
      item_fullnames_map = label_to_fullname_A, 
      save_path = save_path_net,
      filename = paste0("legend_A_", gsub("[[:space:]]+", "_", grp_name), ".png"),
      width = 5.0,
      res = 300
    )
  }
}

# Generate legend for Network B groups (if exists)
if (!is.null(network_data_B) && length(groups_B) > 0) {
  label_to_fullname_B <- setNames(node_fullnames_B, trimws(node_labels_B))
  for (grp_name in names(groups_B)) {
    save_single_group_legend(
      group_name = grp_name,
      items = groups_B[[grp_name]],         
      group_color = group_colors_B[grp_name],
      item_fullnames_map = label_to_fullname_B, 
      save_path = save_path_net,
      filename = paste0("legend_B_", gsub("[[:space:]]+", "_", grp_name), ".png"),
      width = 5.0,
      res = 300
    )
  }
}

# Centrality analysis function
analyze_centrality <- function(network_plot, network_data, suffix) {
  cat("Computing centrality metrics for", suffix, "...\n")
  
  filename <- file.path(save_path_cen, paste0("centrality_plot_", suffix, ".png"))
  png(filename = filename, width = 8, height = 8, units = "in", res = 300)
  
  centrality_results <- centralityPlot(
    network_plot,
    weighted = TRUE,
    signed = TRUE,
    scale = "z-scores",
    labels = colnames(network_data),
    include = "all",
    orderBy = "ExpectedInfluence"
  )
  
  dev.off()
  
  # Save centrality values
  write.csv(centrality_results$data, file = file.path(save_path_cen, paste0("centrality_values_", suffix, ".csv")), row.names = FALSE)
  
  return(centrality_results)
}

# Centrality analysis
centrality_results_A <- analyze_centrality(network_plot_A, npn_A, "A")

centrality_results_B <- NULL
if (!is.null(network_data_B)) {
  centrality_results_B <- analyze_centrality(network_plot_B, npn_B, "B")
}

# Stability analysis function
analyze_stability <- function(network_data, weights, suffix, gamma_val, n_boots) {
  cat("Performing stability analysis for", suffix, "...\n")
  
  # Create custom estimation function
  myWeightedGlasso_bootnet <- function(data, gamma, ...) {
    if (has_weights && !is.null(weight_col)) {
      # Match row indices to get corresponding weights
      data_str <- apply(data, 1, paste, collapse = "_")
      orig_str <- apply(network_data, 1, paste, collapse = "_")
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
    data = as.matrix(network_data),
    default = "none",
    fun = myWeightedGlasso_bootnet,
    type = "case",
    statistics = c("edge", "expectedInfluence", "strength", "closeness", "betweenness"),
    nBoots = n_boots,
    gamma = gamma_val,
    caseN = nrow(network_data),
    maxErrors = 10,
    errorHandling = "continue",
    nCores = 1
  )
  
  # Plot stability graphs
  # Expected Influence stability plot
  stability_plot <- plot(casedrop_boot, statistics = "expectedInfluence", order = "sample") +
    theme(legend.position = "none") +
    ggtitle(paste("Expected Influence Stability -", suffix))
  
  ggsave(
    filename = file.path(save_path_sta, paste0("stability_plot_", suffix, ".png")),
    plot = stability_plot,
    width = 8, height = 8, units = "in", dpi = 300
  )
  
  # Strength stability plot
  strength_plot <- plot(casedrop_boot, statistics = "strength", order = "sample") +
    theme(legend.position = "none") +
    ggtitle(paste("Strength Stability -", suffix))
  
  ggsave(
    filename = file.path(save_path_sta, paste0("strength_stability_plot_", suffix, ".png")),
    plot = strength_plot,
    width = 8, height = 8, units = "in", dpi = 300
  )
  
  # Betweenness stability plot
  betweenness_plot <- plot(casedrop_boot, statistics = "betweenness", order = "sample") +
    theme(legend.position = "none") +
    ggtitle(paste("Betweenness Stability -", suffix))
  
  ggsave(
    filename = file.path(save_path_sta, paste0("betweenness_stability_plot_", suffix, ".png")),
    plot = betweenness_plot,
    width = 8, height = 8, units = "in", dpi = 300
  )
  
  # Closeness stability plot
  closeness_plot <- plot(casedrop_boot, statistics = "closeness", order = "sample") +
    theme(legend.position = "none") +
    ggtitle(paste("Closeness Stability -", suffix))
  
  ggsave(
    filename = file.path(save_path_sta, paste0("closeness_stability_plot_", suffix, ".png")),
    plot = closeness_plot,
    width = 8, height = 8, units = "in", dpi = 300
  )
  
  # CS coefficient calculation
  cs_results <- capture.output(corStability(casedrop_boot, cor = 0.7, statistics = "expectedInfluence"))
  writeLines(cs_results, file.path(save_path_sta, paste0("cs_coefficient_", suffix, ".txt")))
  
  return(casedrop_boot)
}

# Stability analysis
casedrop_boot_A <- analyze_stability(npn_A, weights, "A", gamma_value, n_boots_stability)

casedrop_boot_B <- NULL
if (!is.null(network_data_B)) {
  casedrop_boot_B <- analyze_stability(npn_B, weights, "B", gamma_value, n_boots_stability)
}

# Accuracy analysis function
analyze_accuracy <- function(network_data, weights, suffix, gamma_val, n_boots) {
  cat("Performing accuracy analysis for", suffix, "...\n")
  
  # Create custom estimation function
  myWeightedGlasso_bootnet <- function(data, gamma, ...) {
    if (has_weights && !is.null(weight_col)) {
      # Match row indices to get corresponding weights
      data_str <- apply(data, 1, paste, collapse = "_")
      orig_str <- apply(network_data, 1, paste, collapse = "_")
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
  
  # Nonparametric bootstrap for accuracy
  set.seed(1234)
  nonpar_boot <- bootnet(
    data = as.matrix(network_data),
    default = "none",
    fun = myWeightedGlasso_bootnet,
    nBoots = n_boots,
    gamma = gamma_val,
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
    filename = file.path(save_path_acc, paste0("accuracy_plot_", suffix, ".png")),
    plot = accuracy_plot,
    width = 6, height = 8, units = "in", dpi = 300
  )
  
  # Edge weight difference plot
  edge_diff_plot <- plot(nonpar_boot, "edge", plot = "difference", onlyNonZero = TRUE, order = "sample")
  ggsave(
    filename = file.path(save_path_acc, paste0("edge_difference_plot_", suffix, ".png")),
    plot = edge_diff_plot,
    width = 8, height = 8, units = "in", dpi = 300
  )
  
  return(nonpar_boot)
}

# Accuracy analysis
nonpar_boot_A <- analyze_accuracy(npn_A, weights, "A", gamma_value, n_boots_accuracy)

nonpar_boot_B <- NULL
if (!is.null(network_data_B)) {
  nonpar_boot_B <- analyze_accuracy(npn_B, weights, "B", gamma_value, n_boots_accuracy)
}

# Network Comparison Test (NCT) - Only if both networks exist
if (!is.null(network_data_B)) {
  cat("Preparing Network Comparison Test (NCT) with dynamic weight support...\n")
  
  # Define the weighted estimator
  myWeightedGlasso <- function(data,
                               weights = NULL,
                               gamma,
                               verbose = FALSE,
                               ...) {
    if (!is.null(weights)) {
      # Use weighted covariance
      covm <- cov.wt(data, wt = weights, cor = FALSE)$cov
      C    <- cov2cor(covm)
      n_eff <- sum(weights)^2 / sum(weights^2)  # Effective N
    } else {
      # Use regular correlation matrix
      C <- cor(data)
      n_eff <- nrow(data)
    }
    # Use EBICglasso to estimate network
    net <- qgraph::EBICglasso(C, n = n_eff, gamma = gamma)
    return(net)
  }
    
  # Run NCT with conditional weight passing
  set.seed(1234)
  
  nct_result <- NCT(
    data1            = as.data.frame(npn_A),
    data2            = as.data.frame(npn_B),
    estimator        = myWeightedGlasso,
    estimatorArgs    = list(
      weights = weights,
      gamma   = gamma_value
    ),
    it               = n_boots_compare,
    test.edges       = TRUE,
    edges            = "all",
    p.adjust.methods = "holm", # Change here if you want a different correction method
    test.centrality  = TRUE,
    centrality       = "expectedInfluence",
    verbose          = TRUE
  )
  
  # Helper functions
  to_text <- function(x) {
    if (is.null(x)) return("NULL")
    paste(capture.output(print(x)), collapse = "\n")
  }
  
  create_summary_df <- function(nct_obj) {
    data.frame(
      Indicator = c(
        "Global strength difference value",
        "Global strength values of each network",
        "p-value of global strength difference",
        "Maximum edge weight difference value",
        "p-value of maximum edge weight difference"
      ),
      Value = sapply(list(
        nct_obj$glstrinv.real,
        nct_obj$glstrinv.sep,
        nct_obj$glstrinv.pval,
        nct_obj$nwinv.real,
        nct_obj$nwinv.pval
      ), to_text),
      stringsAsFactors = FALSE
    )
  }
  
  export_nct_to_excel <- function(nct_obj, save_path, file_name = "nct_summary_output_raw.xlsx") {
    wb <- createWorkbook()
    
    addWorksheet(wb, "Summary")
    summary_df <- create_summary_df(nct_obj)
    writeData(wb, sheet = "Summary", x = summary_df)
    
    # Edge weight differences
    addWorksheet(wb, "Edge_Diff_Value")
    writeData(wb, sheet = "Edge_Diff_Value", x = "Matrix of edge weight differences (einv.real)")
    writeData(wb, sheet = "Edge_Diff_Value", x = nct_obj$einv.real, startRow = 3, rowNames = TRUE)
    
    addWorksheet(wb, "Edge_Diff_PValue")
    writeData(wb, sheet = "Edge_Diff_PValue", x = "Matrix of edge weight difference p-values (einv.pvals)")
    writeData(wb, sheet = "Edge_Diff_PValue", x = nct_obj$einv.pvals, startRow = 3, rowNames = TRUE)
    
    # Centrality differences
    if (!is.null(nct_obj$diffcen.real)) {
      addWorksheet(wb, "Centrality_Diff_Value")
      diff_cen_df <- data.frame(
        Node       = row.names(nct_obj$diffcen.real),
        Difference = nct_obj$diffcen.real
      )
      writeData(wb, sheet = "Centrality_Diff_Value", x = "Centrality difference values (diffcen.real)")
      writeData(wb, sheet = "Centrality_Diff_Value", x = diff_cen_df, startRow = 3)
    }
    
    if (!is.null(nct_obj$diffcen.pval)) {
      addWorksheet(wb, "Centrality_Diff_PValue")
      diff_pval_df <- data.frame(
        Node  = row.names(nct_obj$diffcen.pval),
        P_val = nct_obj$diffcen.pval
      )
      writeData(wb, sheet = "Centrality_Diff_PValue", x = "Centrality difference p-values (diffcen.pval)")
      writeData(wb, sheet = "Centrality_Diff_PValue", x = diff_pval_df, startRow = 3)
    }
    
    file_path <- file.path(save_path, file_name)
    saveWorkbook(wb, file_path, overwrite = TRUE)
    cat("NCT detailed results exported to:", file_path, "\n")
  }
  
  # Save and visualize NCT results
  saveRDS(nct_result, file.path(save_path_cmp, "nct_result.rds"))
  cat("NCT result saved as RDS.\n")
  
  export_nct_to_excel(nct_result, save_path_cmp, "nct_summary_output_raw.xlsx")
  
  } else {
  cat("Skipping NCT analysis - only one network provided.\n")
}

# Generate comprehensive report
cat("Generating analysis report...\n")

# Determine analysis type
analysis_type <- if (!is.null(network_data_B)) "Dual Network Analysis with NCT" else "Single Network Analysis"

report_content <- paste0(
  "Network Analysis Report\n",
  "=====================================\n\n",
  "Analysis Type: ", analysis_type, "\n",
  "Analysis Time: ", Sys.time(), "\n",
  "Data File: ", data_file, "\n\n",
  
  # Network A information
  "Network A Overview:\n",
  "- Title: ", network_title_A, "\n",
  "- Sample Size: ", nrow(npn_A), "\n",
  "- Number of Nodes: ", ncol(npn_A), "\n",
  "- Using Weights: ", ifelse(has_weights && !is.null(weight_col), "Yes", "No"), "\n",
  "- Effective Sample Size: ", round(effective_n_A, 2), "\n",
  "- Mean Predictability (R²): ", mean_predictability_A, "\n\n"
)

# Add Network B information if exists
if (!is.null(network_data_B)) {
  report_content <- paste0(report_content,
    "Network B Overview:\n",
    "- Title: ", network_title_B, "\n",
    "- Sample Size: ", nrow(npn_B), "\n",
    "- Number of Nodes: ", ncol(npn_B), "\n",
    "- Using Weights: ", ifelse(has_weights && !is.null(weight_col), "Yes", "No"), "\n",
    "- Effective Sample Size: ", round(effective_n_B, 2), "\n",
    "- Mean Predictability (R²): ", mean_predictability_B, "\n\n"
  )
}

# Add common parameters and output information
report_content <- paste0(report_content,
  "Network Parameters:\n",
  "- EBIC gamma: ", gamma_value, "\n",
  "- Stability Analysis Bootstrap Iterations: ", n_boots_stability, "\n",
  "- Accuracy Analysis Bootstrap Iterations: ", n_boots_accuracy, "\n"
)

if (!is.null(network_data_B)) {
  report_content <- paste0(report_content,
    "- Network Comparison Bootstrap Iterations: ", n_boots_compare, "\n"
  )
}

report_content <- paste0(report_content,
  "\nOutput Files Description:\n",
  "- results/networks/: Network plots and legends\n",
  "- results/centrality/: Centrality analysis results\n",
  "- results/stability/: Stability analysis results\n", 
  "- results/accuracy/: Accuracy analysis results\n",
  "- results/predictability/: Node predictability results\n"
)

if (!is.null(network_data_B)) {
  report_content <- paste0(report_content,
    "- results/comparison/: Network comparison test results\n"
  )
}

# Create main results directory if it doesn't exist
if (!dir.exists(file.path(output_path))) {
  dir.create(file.path(output_path), recursive = TRUE)
}

writeLines(report_content, file.path(output_path, "analysis_report.txt"))

# Completion
cat("\n=====================================\n")
cat("Network analysis completed!\n")
cat("Analysis Type:", analysis_type, "\n")
if (!is.null(network_data_B)) {
  cat("Both networks analyzed and compared using NCT\n")
} else {
  cat("Single network analysis completed\n")
}
cat("All results have been saved to the", output_path, "folder\n")
cat("Please check", file.path(output_path, "analysis_report.txt"), "for detailed report\n")
cat("=====================================\n")
