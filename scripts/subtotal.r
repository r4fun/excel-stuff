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
    
  if(attr(wb, "Class") %>% is.null){
    # Number of sheet with subtotals
    sheet.counts <- 1
    # Set up function list
    addWorksheet(wb, "function.list", visible = FALSE)
    writeDataTable(wb, "function.list", function_table)
  } else {
    # Number of sheet with subtotals
    sheet.counts <- as.integer((attr(wb, "Class"))[2]) + 1
  }
  
  addWorksheet(wb, sheet.name)
  table.name <- glue("data{sheet.counts}")
  writeDataTable(wb, sheet.name, data, tableName = table.name, stack = TRUE)
  
  # Detect last row and last column
  last.col <- ncol(data) + 1
  last.row <- nrow(data) + 2
  
  for (i in 1:ncol(data)) {
    writeFormula(wb,
                 sheet.name,
                 glue("=SUBTOTAL(function.list!$C{sheet.counts + 1},",
                      "{table.name}[{names(data)[i]}])"),
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
    "MATCH({sheet.name}!${excel_col_num(last.col)}${last.row}, ",
    "function.list!$B$2:$B{nrow(function_table) + 1}, 1))"),
    3,
    sheet.counts + 1
  )
  setColWidths(wb, sheet.name, cols = 1:ncol(data), widths = 20)
  
  #Set attribute
  wb <<-
    wb %>%
    structure(., "Class" = c("subtotal", sheet.counts))
}
