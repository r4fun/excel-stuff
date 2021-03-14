
#-------------------------------------------------------------------------------
# Example
#-------------------------------------------------------------------------------
source("scripts/subtotal.r")
wb <- createWorkbook()
set.seed(1024)

# Function adds data tables with subtotals into existing workbooks

data_with_subtotal(
  workbook = wb,
  data = sample_n(iris, 20,  replace = FALSE),
  sheet.name = "data"
)
data_with_subtotal(
  workbook = wb,
  data = sample_n(mtcars, 20,  replace = FALSE),
  sheet.name = "data2"
)

# Save Example
saveWorkbook(wb, "output/subtotal output example.xlsx", overwrite = TRUE)
