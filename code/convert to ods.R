
install.packages("devtools")
devtools::install_github("department-for-transport/odsconvertr")
library(odsconvertr)
convert_to_ods("demo.xlsx")
