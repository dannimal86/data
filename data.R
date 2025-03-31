# Load necessary library
library(openxlsx)

# -----------------------------------------------------------------------------
# --- 1. Data Loading and Preparation ----------------------------------------
# -----------------------------------------------------------------------------

# Load static data
static_data <- read.csv("static_data.csv", stringsAsFactors = FALSE)

# Function to safely convert to numeric (treating non-numeric as 0)
safe_as_numeric <- function(x) {
  # Replace non-numeric values with NA, then convert to numeric, finally replace NA with 0
  x <- as.numeric(ifelse(is.na(as.numeric(x)), NA, x))
  x[is.na(x)] <- 0
  return(x)
}

# -----------------------------------------------------------------------------
# --- 2. CQG Data Processing -------------------------------------------------
# -----------------------------------------------------------------------------

process_cqg_data <- function() {

  # Load CQG data
  cqg_data <- openxlsx::read.xlsx("CQG Data extract.xlsx")

  # Pre-allocate columns for results
  cqg_data$Check <- character(nrow(cqg_data))
  cqg_data$new_max <- character(nrow(cqg_data))
  cqg_data$new_net <- character(nrow(cqg_data))

  # Iterate through each row of CQG data
  for (i in 1:nrow(cqg_data)) {

    # Extract relevant values from the current row
    exchange <- as.character(cqg_data$Exchange[i])
    product <- as.character(cqg_data$Product[i])
    trade_size_limit <- safe_as_numeric(cqg_data$`Trade Size Limit`[i])
    contract_position_limit <- safe_as_numeric(cqg_data$`Contract Position Limit`[i])
    commodity_position_limit <- safe_as_numeric(cqg_data$`Commodity Position Limit`[i])
    type <- as.character(cqg_data$Type[i])

    # Determine if it's an option
    is_option <- grepl("Call option|Put option", type)

    # Find matching row in static_data
    matching_row <- static_data[static_data$`CQG Exchange` == exchange & static_data$`CQG Code` == product, ]

    # Handle cases where no matching entry is found
    if (nrow(matching_row) == 0) {
      cqg_data$Check[i] <- "ILLIQUID PRODUCT"
      cqg_data$new_max[i] <- "NA"
      cqg_data$new_net[i] <- "NA"
      next # Skip to the next iteration
    }

    # Extract Max and Net from static_data
    max_value <- matching_row$Max
    net_value <- matching_row$Net

    # Adjust limits for options
    max_multiplier <- ifelse(is_option, 2, 1)
    net_multiplier <- ifelse(is_option, 2, 1)

    # --- Compare Trade Size Limit against Max ---
    if (trade_size_limit > (max_value * max_multiplier)) {
      cqg_data$Check[i] <- "EXCEPTION"
      cqg_data$new_max[i] <- max_value
    } else {
      cqg_data$new_max[i] <- "NA"
      if (is.na(cqg_data$Check[i])) { #Only populate if not already an illiquid product
        cqg_data$Check[i] <- "PASS"
      }
    }

    # --- Compare Contract Position Limit and Commodity Position Limit against Net ---
    net_exception <- FALSE # Flag to track if either limit exceeds Net

    if (contract_position_limit > (net_value * net_multiplier)) {
      net_exception <- TRUE
    }

    if (commodity_position_limit > (net_value * net_multiplier)) {
      net_exception <- TRUE
    }

    if (net_exception) {
      cqg_data$Check[i] <- "EXCEPTION"
      cqg_data$new_net[i] <- net_value
    } else {
      cqg_data$new_net[i] <- "NA"
      if (is.na(cqg_data$Check[i])) { #Only populate if not already an illiquid product
        cqg_data$Check[i] <- "PASS"
      }
    }
  }

  # Write the output to a new Excel file
  output_file <- "CQG_Data_extract_processed.xlsx"
  openxlsx::write.xlsx(cqg_data, file = output_file, sheetName = "Sheet1",
                       col.names = TRUE, row.names = FALSE, append = FALSE)

  cat("CQG data processing complete. Output file:", output_file, "\n")
}

# -----------------------------------------------------------------------------
# --- 3. TT Data Processing --------------------------------------------------
# -----------------------------------------------------------------------------

process_tt_data <- function() {
  # Load TT data
  tt_data <- openxlsx::read.xlsx("TT data extract.xlsx")

  # Pre-allocate columns for results
  tt_data$Check <- character(nrow(tt_data))
  tt_data$new_max <- character(nrow(tt_data))
  tt_data$new_net <- character(nrow(tt_data))

  # Iterate through each row of TT data
  for (i in 1:nrow(tt_data)) {
    # Extract relevant values from the current row
    exchange <- as.character(tt_data$Exchange[i])
    family <- as.character(tt_data$Family[i])
    max_order_quantity <- safe_as_numeric(tt_data$`Max order quantity`[i])
    max_position_product_net <- safe_as_numeric(tt_data$`Max position product (net)`[i])
    type <- as.character(tt_data$Type[i])
    spreads_max_order_quantity <- safe_as_numeric(tt_data$`Spreads:Max order quantity`[i])

    # Determine if it's an option or spread
    is_option <- grepl("Option", type)
    is_spread <- grepl("Spread|Option Strategy", type)

    # Find matching row in static_data
    matching_row <- static_data[static_data$`TT Exchange` == exchange & static_data$`TT Code` == family, ]

    # Handle cases where no matching entry is found
    if (nrow(matching_row) == 0) {
      tt_data$Check[i] <- "ILLIQUID PRODUCT"
      tt_data$new_max[i] <- "NA"
      tt_data$new_net[i] <- "NA"
      next # Skip to the next iteration
    }

    # Extract Max and Net from static_data
    max_value <- matching_row$Max
    net_value <- matching_row$Net

    # Adjust limits for options and spreads
    max_multiplier <- ifelse(is_option || is_spread, 2, 1)
    net_multiplier <- ifelse(is_option || is_spread, 2, 1)

    # --- Compare Max Order Quantity against Max ---
    limit_to_compare <- ifelse(is_spread, spreads_max_order_quantity, max_order_quantity) # Use spread limit if it's a spread

    if (limit_to_compare > (max_value * max_multiplier)) {
      tt_data$Check[i] <- "EXCEPTION"
      tt_data$new_max[i] <- max_value
    } else {
      tt_data$new_max[i] <- "NA"
      if (is.na(tt_data$Check[i])) { #Only populate if not already an illiquid product
        tt_data$Check[i] <- "PASS"
      }
    }

    # --- Compare Max Position Product (net) against Net ---
    if (max_position_product_net > (net_value * net_multiplier)) {
      tt_data$Check[i] <- "EXCEPTION"
      tt_data$new_net[i] <- net_value
    } else {
      tt_data$new_net[i] <- "NA"
      if (is.na(tt_data$Check[i])) { #Only populate if not already an illiquid product
        tt_data$Check[i] <- "PASS"
      }
    }
  }

  # Write the output to a new Excel file
  output_file <- "TT_data_extract_processed.xlsx"
  openxlsx::write.xlsx(tt_data, file = output_file, sheetName = "Sheet1",
                       col.names = TRUE, row.names = FALSE, append = FALSE)

  cat("TT data processing complete. Output file:", output_file, "\n")
}

# -----------------------------------------------------------------------------
# --- 4. Fidessa Data Processing ---------------------------------------------
# -----------------------------------------------------------------------------

process_fidessa_data <- function() {
  # Load Fidessa data
  fidessa_data <- openxlsx::read.xlsx("Fidessa data extract.xlsx")

  # Pre-allocate columns for results
  fidessa_data$Check <- character(nrow(fidessa_data))
  fidessa_data$new_max <- character(nrow(fidessa_data))
  fidessa_data$new_net <- character(nrow(fidessa_data))

  # Iterate through each row of Fidessa data
  for (i in 1:nrow(fidessa_data)) {
    # Extract relevant values from the current row
    exchange <- as.character(fidessa_data$Exchange[i])
    product <- as.character(fidessa_data$Product[i])
    max_ord_size <- safe_as_numeric(fidessa_data$`Max ord size`[i])
    max_pos_size <- safe_as_numeric(fidessa_data$`Max pos size`[i])
    max_spread_pos <- safe_as_numeric(fidessa_data$`Max spread pos`[i])
    asset_class <- as.character(fidessa_data$`Asset class`[i])

    # Determine if it's an option or spread
    is_option <- grepl("OptOut", asset_class)
    is_spread <- grepl("FuStr|OpStr", asset_class)

    # Find matching row in static_data
    matching_row <- static_data[static_data$`Fidessa Exchange` == exchange & static_data$`Fidessa Code` == product, ]

    # Handle cases where no matching entry is found
    if (nrow(matching_row) == 0) {
      fidessa_data$Check[i] <- "ILLIQUID PRODUCT"
      fidessa_data$new_max[i] <- "NA"
      fidessa_data$new_net[i] <- "NA"
      next # Skip to the next iteration
    }

    # Extract Max and Net from static_data
    max_value <- matching_row$Max
    net_value <- matching_row$Net

    # Adjust limits for options and spreads
    max_multiplier <- ifelse(is_option || is_spread, 2, 1)
    net_multiplier <- ifelse(is_option || is_spread, 2, 1)

    # --- Compare Max Order Size against Max ---
    if (max_ord_size > (max_value * max_multiplier)) {
      fidessa_data$Check[i] <- "EXCEPTION"
      fidessa_data$new_max[i] <- max_value
    } else {
      fidessa_data$new_max[i] <- "NA"
       if (is.na(fidessa_data$Check[i])) { #Only populate if not already an illiquid product
        fidessa_data$Check[i] <- "PASS"
      }
    }

    # --- Compare Max Position Size (or Max Spread Pos for Spreads) against Net ---
    limit_to_compare <- ifelse(is_spread, max_spread_pos, max_pos_size) # Use spread limit if it's a spread

    if (limit_to_compare > (net_value * net_multiplier)) {
      fidessa_data$Check[i] <- "EXCEPTION"
      fidessa_data$new_net[i] <- net_value
    } else {
      fidessa_data$new_net[i] <- "NA"
       if (is.na(fidessa_data$Check[i])) { #Only populate if not already an illiquid product
        fidessa_data$Check[i] <- "PASS"
      }
    }
  }

  # Write the output to a new Excel file
  output_file <- "Fidessa_data_extract_processed.xlsx"
  openxlsx::write.xlsx(fidessa_data, file = output_file, sheetName = "Sheet1",
                       col.names = TRUE, row.names = FALSE, append = FALSE)

  cat("Fidessa data processing complete. Output file:", output_file, "\n")
}

# -----------------------------------------------------------------------------
# --- 5. Main Execution ------------------------------------------------------
# -----------------------------------------------------------------------------

# Process each data source
process_cqg_data()
process_tt_data()
process_fidessa_data()

cat("All data processing complete.\n")
