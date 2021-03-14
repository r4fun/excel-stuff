library(dplyr)
library(openxlsx)
library(glue)

# Function starts here

data_with_subtotal <- function(workbook, data, sheet.name){

  # Detect excel column name
  excel_col_num <- function(n){
    index <- NULL
    
    while(n > 1){
      if(n %% 26 %in% 0){
        index <- paste0("Z", index) 
      } else {
        index <- paste0(LETTERS[n %% 26], index)
      }
      n <- (n-1)/26
    }
    return(index)
  }
  
  # Excel aggregate function list
  function_table <-
    data.frame(
      "Num" = c(101:111),
      "Function" = c("AVERGE",
                     "COUNT",
                     "COUNTA",
                     "MAX",
                     "MIN",
                     "PRODUCT",
                     "STDEV",
                     "STDEVP",
                     "SUM",
                     "VAR",
                     "VARP"
      )
    )
  
  # Set up function list
  addWorksheet(wb, "function.list", visible = FALSE)
  writeDataTable(wb, "function.list", function_table)
  addWorksheet(wb, sheet.name)
  table.name <- "data"
  writeDataTable(wb, sheet.name, data, tableName = table.name, stack = TRUE)
  
  # Detect last row and last column
  last.col <- ncol(data) + 1
  last.row <- nrow(data) + 2
  
  for (i in 1:ncol(data)) {
    writeFormula(wb,
                 sheet.name,
                 glue("=SUBTOTAL(function.list!$C2,data[{names(data)[i]}])"),
                 i,
                 last.row)
  }
  
  # Default function
  writeData(wb, sheet.name, "SUM", last.col, last.row)
  
  # Add drop-downs
  dataValidation(
    wb,
    sheet.name,
    col = last.col,
    rows = last.row,
    type = "list",
    value =
      glue("'function.list'!$B$2:$B${nrow(function_table) + 1}")
  )
  writeData(wb, "function.list", "Detected Function", 3, 1)
  writeFormula(
    wb,
    "function.list",
    glue("=INDEX(function.list!$A$2:$A{nrow(function_table) + 1}, ", 
    "MATCH(data!${excel_col_num(last.col)}${last.row}, ",
    "function.list!$B$2:$B{nrow(function_table) + 1}, 1))"),
    3,
    2
  )
  setColWidths(wb, sheet.name, cols = 1:ncol(data), widths = "auto")
}

#-------------------------------------------------------------------------------
# Example
#-------------------------------------------------------------------------------

wb <- createWorkbook()
set.seed(1024)

# Function allows one to add data table to existing workbook.
data_with_subtotal(
  workbook = wb,
  data = sample_n(iris, 20,  replace = FALSE),
  sheet.name = "data"
)

# Save workbook
saveWorkbook(wb, "output/subtotal irist sample.xlsx", overwrite = TRUE)
