# ============================================================================
# Trading Limits Comparison Script
# 
# Purpose: This script compares trading limits from three different sources 
# (CQG, TT, and Fidessa) against reference values in a static data file.
# It identifies exceptions where limits exceed thresholds and generates 
# output files with additional columns showing validation results.
#
# Input Files:
# - static_data.csv: Reference data containing Max and Net limit values
# - CQG Data extract.xlsx: Trading limits from CQG platform
# - TT data extract.xls: Trading limits from TT platform
# - Fidessa data extract.xlsx: Trading limits from Fidessa platform
#
# Output Files:
# - CQG_output.csv: CQG data with added Check, new_max, and new_net columns
# - TT_output.csv: TT data with added Check, new_max, and new_net columns 
# - Fidessa_output.csv: Fidessa data with added Check, new_max, and new_net columns
# ============================================================================

# Load required libraries
library(readr)     # For reading and writing CSV files
library(dplyr)     # For data manipulation operations
library(stringr)   # For string manipulation functions
library(openxlsx)  # For reading Excel files without read_excel
library(tools)     # For file extension utilities

# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

#' Safe conversion of values to numeric type
#'
#' This function attempts to convert input values to numeric, handling various
#' non-numeric representations like empty strings, NA values, or text 
#' descriptions like "No Limit". When these are encountered, they are 
#' converted to 0 as per requirements.
#'
#' @param x The value to convert to numeric
#' @return A numeric value (0 if conversion failed or non-numeric input)
safe_as_numeric <- function(x) {
  # First handle various non-numeric values by converting them to "0"
  # - NA values
  # - Empty strings
  # - Text values containing "no limit" (case insensitive)
  x <- ifelse(is.na(x) | x == "" | str_detect(tolower(as.character(x)), "no limit"), "0", x)
  
  # Now try to convert to numeric
  # If conversion fails (e.g., for other text values), it will produce NA
  # which we replace with 0 as per requirements
  result <- suppressWarnings(as.numeric(as.character(x)))
  ifelse(is.na(result), 0, result)
}

#' Check if a value is numeric
#'
#' Helper function to determine if a value can be converted to a numeric type.
#' Used to validate inputs before numeric operations.
#'
#' @param x The value to check
#' @return Logical indicating if x can be converted to numeric
is_numeric_value <- function(x) {
  !is.na(suppressWarnings(as.numeric(as.character(x))))
}

#' Compare limit values against static reference values
#'
#' This function compares a limit value from a trading system against 
#' the reference value from static data. For spreads and options, limits 
#' can be up to 2x the static value before triggering an exception.
#'
#' @param limit_val The limit value from a trading system
#' @param static_val The reference value from static data
#' @param is_spread_or_option Logical indicating if the product is a spread or option
#' @return A list with 'check' status ("EXCEPTION" or "PASS") and 'new_val' 
#'         containing the static value if an exception was found
compare_limits <- function(limit_val, static_val, is_spread_or_option) {
  # Convert inputs to numeric safely
  limit_val <- safe_as_numeric(limit_val)
  static_val <- as.numeric(static_val)
  
  # For spreads and options, the limit can be up to 2x the static value
  # For regular products, the multiplier is 1x (no extra allowance)
  multiplier <- ifelse(is_spread_or_option, 2, 1)
  
  # Compare the limit value against the static value (with multiplier)
  # and return the appropriate result
  if (limit_val > (static_val * multiplier)) {
    # If limit exceeds the threshold, return an exception and the static value
    return(list(
      check = "EXCEPTION",
      new_val = static_val
    ))
  } else {
    # If limit is within threshold, return a pass and NA for new_val
    return(list(
      check = "PASS",
      new_val = NA
    ))
  }
}

#' Read Excel file safely
#'
#' This function reads an Excel file (.xlsx or .xls) using the openxlsx package
#' instead of readxl. It handles potential errors and returns NULL if reading fails.
#'
#' @param file_path Path to the Excel file
#' @param sheet Sheet name or index to read (defaults to 1)
#' @return A data frame containing the Excel data, or NULL if reading fails
read_excel_safe <- function(file_path, sheet = 1) {
  tryCatch({
    # Check file extension to determine how to read it
    ext <- tolower(file_ext(file_path))
    
    if (ext == "xlsx" || ext == "xls") {
      # Read using openxlsx
      data <- openxlsx::read.xlsx(file_path, sheet = sheet, detectDates = TRUE)
      
      # Replace empty strings with NA
      data[] <- lapply(data, function(x) {
        ifelse(x == "", NA, x)
      })
      
      return(data)
    } else {
      cat("Unsupported file format:", ext, "\n")
      return(NULL)
    }
  }, error = function(e) {
    cat("Error reading Excel file:", e$message, "\n")
    return(NULL)
  })
}

# =============================================================================
# MAIN PROCESSING FUNCTION
# =============================================================================

#' Main function to process all trading limits files
#'
#' This function orchestrates the entire process:
#' 1. Reads the static data reference file
#' 2. Processes each of the three trading system files (CQG, TT, Fidessa)
#' 3. Generates output files with comparison results
process_files <- function() {
  # =========================================================================
  # Read static data file
  # =========================================================================
  cat("Reading static_data.csv...\n")
  static_data <- tryCatch({
    # Explicitly specify column types to ensure proper data interpretation
    read_csv("static_data.csv", col_types = cols(
      `Fidessa Exchange` = col_character(),
      `CQG Exchange` = col_character(),
      `TT Exchange` = col_character(),
      `Fidessa Code` = col_character(),
      `CQG Code` = col_character(),
      `TT Code` = col_character(),
      `BBG Code` = col_character(),
      `Max` = col_integer(),
      `Net` = col_integer()
    ))
  }, error = function(e) {
    # Handle errors when reading the file
    cat("Error reading static_data.csv:", e$message, "\n")
    return(NULL)
  })
  
  # Exit function if static_data couldn't be loaded
  if (is.null(static_data)) {
    cat("Cannot proceed without static data. Please check the file and try again.\n")
    return()
  }
  
  # =========================================================================
  # Process CQG data
  # =========================================================================
  cat("Processing CQG data...\n")
  tryCatch({
    # Read CQG data from Excel file using our custom function
    cqg_data <- read_excel_safe("CQG Data extract.xlsx")
    
    # Exit if file couldn't be read
    if (is.null(cqg_data)) {
      cat("Skipping CQG processing due to file reading error.\n")
      return()
    }
    
    # Process CQG data
    # The rowwise() function is used to perform operations on each row independently
    cqg_output <- cqg_data %>%
      rowwise() %>%
      mutate(
        # Determine if the product is an option based on the Type column
        # Options are identified by "call option" or "put option" in the Type field
        is_option = str_detect(tolower(as.character(Type)), "call option|put option"),
        
        # Find the matching record in static_data by Exchange and Product code
        # This returns the entire matching row from static_data if found, otherwise NULL
        static_match = list({
          match_idx <- which(static_data$`CQG Exchange` == Exchange & 
                             static_data$`CQG Code` == Product)
          if (length(match_idx) > 0) {
            static_data[match_idx[1], ]
          } else {
            NULL
          }
        }),
        
        # Flag indicating whether a match was found in static_data
        has_match = !is.null(static_match),
        
        # Compare Trade Size Limit against Max from static_data
        # If no match found, set to "ILLIQUID PRODUCT"
        max_result = if (has_match) {
          compare_limits(`Trade Size Limit`, static_match$Max, is_option)
        } else {
          list(check = "ILLIQUID PRODUCT", new_val = NA)
        },
        
        # Compare Contract Position Limit against Net from static_data
        net_result_contract = if (has_match) {
          compare_limits(`Contract Position Limit`, static_match$Net, is_option)
        } else {
          list(check = "ILLIQUID PRODUCT", new_val = NA)
        },
        
        # Compare Commodity Position Limit against Net from static_data
        net_result_commodity = if (has_match) {
          compare_limits(`Commodity Position Limit`, static_match$Net, is_option)
        } else {
          list(check = "ILLIQUID PRODUCT", new_val = NA)
        },
        
        # Combine results to determine final Check status
        # If any comparison results in EXCEPTION, the overall Check is EXCEPTION
        Check = case_when(
          !has_match ~ "ILLIQUID PRODUCT",
          max_result$check == "EXCEPTION" ~ "EXCEPTION",
          net_result_contract$check == "EXCEPTION" ~ "EXCEPTION",
          net_result_commodity$check == "EXCEPTION" ~ "EXCEPTION",
          TRUE ~ "PASS"
        ),
        
        # Set new_max to static_data Max if max_result is EXCEPTION, otherwise NA
        new_max = if (has_match && max_result$check == "EXCEPTION") static_match$Max else NA,
        
        # Set new_net to static_data Net if either net_result is EXCEPTION, otherwise NA
        new_net = if (has_match && (net_result_contract$check == "EXCEPTION" || 
                                   net_result_commodity$check == "EXCEPTION")) 
                    static_match$Net else NA
      ) %>%
      # Remove intermediate calculation columns from final output
      select(-is_option, -static_match, -has_match, -max_result, 
             -net_result_contract, -net_result_commodity)
    
    # Write CQG output to CSV file
    write_csv(cqg_output, "CQG_output.csv")
    cat("CQG output written to CQG_output.csv\n")
    
  }, error = function(e) {
    # Handle errors during CQG processing
    cat("Error processing CQG data:", e$message, "\n")
  })
  
  # =========================================================================
  # Process TT data
  # =========================================================================
  cat("Processing TT data...\n")
  tryCatch({
    # Read TT data from Excel file using our custom function
    tt_data <- read_excel_safe("TT data extract.xls")
    
    # Exit if file couldn't be read
    if (is.null(tt_data)) {
      cat("Skipping TT processing due to file reading error.\n")
      return()
    }
    
    # Process TT data
    tt_output <- tt_data %>%
      rowwise() %>%
      mutate(
        # Determine if the product is an option based on the Type column
        # Options are identified by "option" in the Type field
        is_option = str_detect(tolower(as.character(Type)), "option"),
        
        # Determine if the product is a spread based on the Type column
        # Spreads are identified by "spread" or "option strategy" in the Type field
        is_spread = str_detect(tolower(as.character(Type)), "spread|option strategy"),
        
        # Flag for either spread or option (both get 2x allowance)
        is_spread_or_option = is_option | is_spread,
        
        # Find the matching record in static_data by Exchange and Family code
        static_match = list({
          match_idx <- which(static_data$`TT Exchange` == Exchange & 
                             static_data$`TT Code` == Family)
          if (length(match_idx) > 0) {
            static_data[match_idx[1], ]
          } else {
            NULL
          }
        }),
        
        # Flag indicating whether a match was found in static_data
        has_match = !is.null(static_match),
        
        # Determine which max order quantity to use based on whether it's a spread
        # For spreads, use Spreads:Max order quantity
        # For others, use Max order quantity
        max_order_qty = if (is_spread) {
          safe_as_numeric(`Spreads:Max order quantity`)
        } else {
          safe_as_numeric(`Max order quantity`)
        },
        
        # Compare max order quantity against Max from static_data
        max_result = if (has_match) {
          compare_limits(max_order_qty, static_match$Max, is_spread_or_option)
        } else {
          list(check = "ILLIQUID PRODUCT", new_val = NA)
        },
        
        # Compare Max position product (net) against Net from static_data
        net_result = if (has_match) {
          compare_limits(`Max position product (net)`, static_match$Net, is_spread_or_option)
        } else {
          list(check = "ILLIQUID PRODUCT", new_val = NA)
        },
        
        # Combine results to determine final Check status
        Check = case_when(
          !has_match ~ "ILLIQUID PRODUCT",
          max_result$check == "EXCEPTION" ~ "EXCEPTION",
          net_result$check == "EXCEPTION" ~ "EXCEPTION",
          TRUE ~ "PASS"
        ),
        
        # Set new_max to static_data Max if max_result is EXCEPTION, otherwise NA
        new_max = if (has_match && max_result$check == "EXCEPTION") static_match$Max else NA,
        
        # Set new_net to static_data Net if net_result is EXCEPTION, otherwise NA
        new_net = if (has_match && net_result$check == "EXCEPTION") static_match$Net else NA
      ) %>%
      # Remove intermediate calculation columns from final output
      select(-is_option, -is_spread, -is_spread_or_option, -static_match, 
             -has_match, -max_order_qty, -max_result, -net_result)
    
    # Write TT output to CSV file
    write_csv(tt_output, "TT_output.csv")
    cat("TT output written to TT_output.csv\n")
    
  }, error = function(e) {
    # Handle errors during TT processing
    cat("Error processing TT data:", e$message, "\n")
  })
  
  # =========================================================================
  # Process Fidessa data
  # =========================================================================
  cat("Processing Fidessa data...\n")
  tryCatch({
    # Read Fidessa data from Excel file using our custom function
    fidessa_data <- read_excel_safe("Fidessa data extract.xlsx")
    
    # Exit if file couldn't be read
    if (is.null(fidessa_data)) {
      cat("Skipping Fidessa processing due to file reading error.\n")
      return()
    }
    
    # Process Fidessa data
    fidessa_output <- fidessa_data %>%
      rowwise() %>%
      mutate(
        # Determine if the product is an option based on the Asset class column
        # Options are identified by "OptOut" in the Asset class field
        is_option = !is.na(`Asset class`) && str_detect(`Asset class`, "OptOut"),
        
        # Determine if the product is a spread based on the Asset class column
        # Spreads are identified by "FuStr" or "OpStr" in the Asset class field
        is_spread = !is.na(`Asset class`) && str_detect(`Asset class`, "FuStr|OpStr"),
        
        # Flag for either spread or option (both get 2x allowance)
        is_spread_or_option = is_option | is_spread,
        
        # Find the matching record in static_data by Exchange and Product code
        static_match = list({
          match_idx <- which(static_data$`Fidessa Exchange` == Exchange & 
                             static_data$`Fidessa Code` == Product)
          if (length(match_idx) > 0) {
            static_data[match_idx[1], ]
          } else {
            NULL
          }
        }),
        
        # Flag indicating whether a match was found in static_data
        has_match = !is.null(static_match),
        
        # Compare Max ord size against Max from static_data
        max_result = if (has_match) {
          compare_limits(`Max ord size`, static_match$Max, is_spread_or_option)
        } else {
          list(check = "ILLIQUID PRODUCT", new_val = NA)
        },
        
        # Determine which position size to check based on whether it's a spread
        # For spreads, use Max spread pos
        # For others, use Max pos size
        net_size_to_check = if (is_spread) {
          `Max spread pos`
        } else {
          `Max pos size`
        },
        
        # Compare the appropriate position size against Net from static_data
        net_result = if (has_match) {
          compare_limits(net_size_to_check, static_match$Net, is_spread_or_option)
        } else {
          list(check = "ILLIQUID PRODUCT", new_val = NA)
        },
        
        # Combine results to determine final Check status
        Check = case_when(
          !has_match ~ "ILLIQUID PRODUCT",
          max_result$check == "EXCEPTION" ~ "EXCEPTION",
          net_result$check == "EXCEPTION" ~ "EXCEPTION",
          TRUE ~ "PASS"
        ),
        
        # Set new_max to static_data Max if max_result is EXCEPTION, otherwise NA
        new_max = if (has_match && max_result$check == "EXCEPTION") static_match$Max else NA,
        
        # Set new_net to static_data Net if net_result is EXCEPTION, otherwise NA
        new_net = if (has_match && net_result$check == "EXCEPTION") static_match$Net else NA
      ) %>%
      # Remove intermediate calculation columns from final output
      select(-is_option, -is_spread, -is_spread_or_option, -static_match, 
             -has_match, -net_size_to_check, -max_result, -net_result)
    
    # Write Fidessa output to CSV file
    write_csv(fidessa_output, "Fidessa_output.csv")
    cat("Fidessa output written to Fidessa_output.csv\n")
    
  }, error = function(e) {
    # Handle errors during Fidessa processing
    cat("Error processing Fidessa data:", e$message, "\n")
  })
  
  cat("Processing complete.\n")
}

# =============================================================================
# SCRIPT EXECUTION
# =============================================================================

# Run the main processing function
# This will read all input files, perform the comparisons, and generate output files
process_files()
