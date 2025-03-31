# Script to process trading limits data against static reference data
# This script compares trading limits from CQG, TT, and Fidessa files against a static reference file
# and flags exceptions where limits exceed reference values

# Load required libraries
library(openxlsx)

# Error handling function to make script more robust
safe_as_numeric <- function(x) {
  # Convert values to numeric, treating non-numeric values as 0
  result <- suppressWarnings(as.numeric(as.character(x)))
  result[is.na(result)] <- 0
  return(result)
}

# Main processing function
process_trading_data <- function() {
  # Set working directory to the location of your files if needed
  # setwd("/path/to/your/files")
  
  cat("Starting data processing...\n")
  
  # Step 1: Load the static data file
  cat("Loading static reference data...\n")
  static_data <- try(read.csv("static_data.csv", stringsAsFactors = FALSE), silent = TRUE)
  
  if(inherits(static_data, "try-error")) {
    stop("Error loading static_data.csv. Please check if the file exists and is accessible.")
  }
  
  # Step 2: Process CQG data
  cat("Processing CQG data...\n")
  process_cqg_data(static_data)
  
  # Step 3: Process TT data
  cat("Processing TT data...\n")
  process_tt_data(static_data)
  
  # Step 4: Process Fidessa data
  cat("Processing Fidessa data...\n")
  process_fidessa_data(static_data)
  
  cat("Data processing completed successfully.\n")
}

# Function to process CQG data
process_cqg_data <- function(static_data) {
  # Load CQG data
  cqg_file <- try(loadWorkbook("CQG Data extract.xlsx"), silent = TRUE)
  
  if(inherits(cqg_file, "try-error")) {
    cat("Error loading CQG Data extract.xlsx. Skipping CQG processing.\n")
    return()
  }
  
  # Get sheet names
  sheet_names <- sheets(cqg_file)
  
  if(length(sheet_names) == 0) {
    cat("No sheets found in CQG Data extract.xlsx. Skipping CQG processing.\n")
    return()
  }
  
  # Read the first sheet (or specified sheet if needed)
  cqg_data <- try(read.xlsx(cqg_file, sheet = 1), silent = TRUE)
  
  if(inherits(cqg_data, "try-error") || ncol(cqg_data) < 5) {
    cat("Error reading CQG data or unexpected format. Skipping CQG processing.\n")
    return()
  }
  
  # Create output dataframe with required additional columns
  cqg_output <- cqg_data
  cqg_output$Check <- character(nrow(cqg_data))
  cqg_output$new_max <- character(nrow(cqg_data))
  cqg_output$new_net <- character(nrow(cqg_data))
  
  # Process each row
  for(i in 1:nrow(cqg_data)) {
    # Extract relevant information
    exchange <- cqg_data[i, "Exchange"]
    product <- cqg_data[i, "Product"]
    trade_size_limit <- safe_as_numeric(cqg_data[i, "Trade Size Limit"])
    contract_position_limit <- safe_as_numeric(cqg_data[i, "Contract Position Limit"])
    commodity_position_limit <- safe_as_numeric(cqg_data[i, "Commodity Position Limit"])
    type <- cqg_data[i, "Type"]
    
    # Check if this is an option (which can have limits up to 2x the reference)
    is_option <- FALSE
    if(!is.null(type) && !is.na(type)) {
      is_option <- grepl("Call option|Put option", type, ignore.case = TRUE)
    }
    
    # Find matching entry in static data
    match_idx <- which(static_data$`CQG Exchange` == exchange & 
                       static_data$`CQG Code` == product)
    
    if(length(match_idx) > 0) {
      # Get reference values
      ref_max <- static_data[match_idx[1], "Max"]
      ref_net <- static_data[match_idx[1], "Net"]
      
      # Apply multiplier for options (up to 2x)
      multiplier <- ifelse(is_option, 2, 1)
      
      # Compare limits
      max_exception <- trade_size_limit > (ref_max * multiplier)
      net_exception_contract <- contract_position_limit > (ref_net * multiplier)
      net_exception_commodity <- commodity_position_limit > (ref_net * multiplier)
      
      # Set Check value based on comparisons
      if(max_exception || net_exception_contract || net_exception_commodity) {
        cqg_output$Check[i] <- "EXCEPTION"
      } else {
        cqg_output$Check[i] <- "PASS"
      }
      
      # Set new_max and new_net values
      if(max_exception) {
        cqg_output$new_max[i] <- as.character(ref_max)
      } else {
        cqg_output$new_max[i] <- "NA"
      }
      
      if(net_exception_contract || net_exception_commodity) {
        cqg_output$new_net[i] <- as.character(ref_net)
      } else {
        cqg_output$new_net[i] <- "NA"
      }
    } else {
      # No match found in static data
      cqg_output$Check[i] <- "ILLIQUID PRODUCT"
      cqg_output$new_max[i] <- "NA"
      cqg_output$new_net[i] <- "NA"
    }
  }
  
  # Write output to file
  try(write.xlsx(cqg_output, "CQG_processed_output.xlsx"), silent = TRUE)
  cat("CQG data processed and saved to CQG_processed_output.xlsx\n")
}

# Function to process TT data
process_tt_data <- function(static_data) {
  # Load TT data
  tt_file <- try(loadWorkbook("TT data extract.xlsx"), silent = TRUE)
  
  if(inherits(tt_file, "try-error")) {
    cat("Error loading TT data extract.xlsx. Skipping TT processing.\n")
    return()
  }
  
  # Get sheet names
  sheet_names <- sheets(tt_file)
  
  if(length(sheet_names) == 0) {
    cat("No sheets found in TT data extract.xlsx. Skipping TT processing.\n")
    return()
  }
  
  # Read the first sheet (or specified sheet if needed)
  tt_data <- try(read.xlsx(tt_file, sheet = 1), silent = TRUE)
  
  if(inherits(tt_data, "try-error") || ncol(tt_data) < 7) {
    cat("Error reading TT data or unexpected format. Skipping TT processing.\n")
    return()
  }
  
  # Create output dataframe with required additional columns
  tt_output <- tt_data
  tt_output$Check <- character(nrow(tt_data))
  tt_output$new_max <- character(nrow(tt_data))
  tt_output$new_net <- character(nrow(tt_data))
  
  # Process each row
  for(i in 1:nrow(tt_data)) {
    # Extract relevant information
    exchange <- tt_data[i, "Exchange"]
    family <- tt_data[i, "Family"]
    max_order_quantity <- safe_as_numeric(tt_data[i, "Max order quantity"])
    max_position_net <- safe_as_numeric(tt_data[i, "Max position product (net)"])
    type <- tt_data[i, "Type"]
    
    # Check if this is an option or spread (which can have limits up to 2x the reference)
    is_option <- FALSE
    is_spread <- FALSE
    
    if(!is.null(type) && !is.na(type)) {
      is_option <- grepl("Option", type, ignore.case = TRUE)
      is_spread <- grepl("Spread|Option Strategy", type, ignore.case = TRUE)
    }
    
    # For spreads, use the spread-specific max order quantity
    if(is_spread) {
      max_order_quantity <- safe_as_numeric(tt_data[i, "Spreads:Max order quantity"])
    }
    
    # Find matching entry in static data
    match_idx <- which(static_data$`TT Exchange` == exchange & 
                       static_data$`TT Code` == family)
    
    if(length(match_idx) > 0) {
      # Get reference values
      ref_max <- static_data[match_idx[1], "Max"]
      ref_net <- static_data[match_idx[1], "Net"]
      
      # Apply multiplier for options and spreads (up to 2x)
      multiplier <- ifelse(is_option || is_spread, 2, 1)
      
      # Compare limits
      max_exception <- max_order_quantity > (ref_max * multiplier)
      net_exception <- max_position_net > (ref_net * multiplier)
      
      # Set Check value based on comparisons
      if(max_exception || net_exception) {
        tt_output$Check[i] <- "EXCEPTION"
      } else {
        tt_output$Check[i] <- "PASS"
      }
      
      # Set new_max and new_net values
      if(max_exception) {
        tt_output$new_max[i] <- as.character(ref_max)
      } else {
        tt_output$new_max[i] <- "NA"
      }
      
      if(net_exception) {
        tt_output$new_net[i] <- as.character(ref_net)
      } else {
        tt_output$new_net[i] <- "NA"
      }
    } else {
      # No match found in static data
      tt_output$Check[i] <- "ILLIQUID PRODUCT"
      tt_output$new_max[i] <- "NA"
      tt_output$new_net[i] <- "NA"
    }
  }
  
  # Write output to file
  try(write.xlsx(tt_output, "TT_processed_output.xlsx"), silent = TRUE)
  cat("TT data processed and saved to TT_processed_output.xlsx\n")
}

# Function to process Fidessa data
process_fidessa_data <- function(static_data) {
  # Load Fidessa data
  fidessa_file <- try(loadWorkbook("Fidessa data extract.xlsx"), silent = TRUE)
  
  if(inherits(fidessa_file, "try-error")) {
    cat("Error loading Fidessa data extract.xlsx. Skipping Fidessa processing.\n")
    return()
  }
  
  # Get sheet names
  sheet_names <- sheets(fidessa_file)
  
  if(length(sheet_names) == 0) {
    cat("No sheets found in Fidessa data extract.xlsx. Skipping Fidessa processing.\n")
    return()
  }
  
  # Read the first sheet (or specified sheet if needed)
  fidessa_data <- try(read.xlsx(fidessa_file, sheet = 1), silent = TRUE)
  
  if(inherits(fidessa_data, "try-error") || ncol(fidessa_data) < 7) {
    cat("Error reading Fidessa data or unexpected format. Skipping Fidessa processing.\n")
    return()
  }
  
  # Create output dataframe with required additional columns
  fidessa_output <- fidessa_data
  fidessa_output$Check <- character(nrow(fidessa_data))
  fidessa_output$new_max <- character(nrow(fidessa_data))
  fidessa_output$new_net <- character(nrow(fidessa_data))
  
  # Process each row
  for(i in 1:nrow(fidessa_data)) {
    # Extract relevant information
    exchange <- fidessa_data[i, "Exchange"]
    product <- fidessa_data[i, "Product"]
    max_ord_size <- safe_as_numeric(fidessa_data[i, "Max ord size"])
    max_pos_size <- safe_as_numeric(fidessa_data[i, "Max pos size"])
    max_spread_pos <- safe_as_numeric(fidessa_data[i, "Max spread pos"])
    asset_class <- fidessa_data[i, "Asset class"]
    
    # Check if this is an option or spread (which can have limits up to 2x the reference)
    is_option <- FALSE
    is_spread <- FALSE
    
    if(!is.null(asset_class) && !is.na(asset_class)) {
      is_option <- grepl("OptOut", asset_class, ignore.case = TRUE)
      is_spread <- grepl("FuStr|OpStr", asset_class, ignore.case = TRUE)
    }
    
    # For spreads, use the spread-specific position limit
    position_limit <- ifelse(is_spread, max_spread_pos, max_pos_size)
    
    # Find matching entry in static data
    match_idx <- which(static_data$`Fidessa Exchange` == exchange & 
                       static_data$`Fidessa Code` == product)
    
    if(length(match_idx) > 0) {
      # Get reference values
      ref_max <- static_data[match_idx[1], "Max"]
      ref_net <- static_data[match_idx[1], "Net"]
      
      # Apply multiplier for options and spreads (up to 2x)
      multiplier <- ifelse(is_option || is_spread, 2, 1)
      
      # Compare limits
      max_exception <- max_ord_size > (ref_max * multiplier)
      net_exception <- position_limit > (ref_net * multiplier)
      
      # Set Check value based on comparisons
      if(max_exception || net_exception) {
        fidessa_output$Check[i] <- "EXCEPTION"
      } else {
        fidessa_output$Check[i] <- "PASS"
      }
      
      # Set new_max and new_net values
      if(max_exception) {
        fidessa_output$new_max[i] <- as.character(ref_max)
      } else {
        fidessa_output$new_max[i] <- "NA"
      }
      
      if(net_exception) {
        fidessa_output$new_net[i] <- as.character(ref_net)
      } else {
        fidessa_output$new_net[i] <- "NA"
      }
    } else {
      # No match found in static data
      fidessa_output$Check[i] <- "ILLIQUID PRODUCT"
      fidessa_output$new_max[i] <- "NA"
      fidessa_output$new_net[i] <- "NA"
    }
  }
  
  # Write output to file
  try(write.xlsx(fidessa_output, "Fidessa_processed_output.xlsx"), silent = TRUE)
  cat("Fidessa data processed and saved to Fidessa_processed_output.xlsx\n")
}

# Execute the main function
process_trading_data()
