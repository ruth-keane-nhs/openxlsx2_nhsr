
library(openxlsx2)
library(dplyr)
library(readr)
library(readxl)
library(tidyr)
library(stringr)
library(lubridate)

title_lookup <- read_csv("src/create_report/title_lookup.csv")
date_up_to<- format(dmy("06.10.2024"),"w-e %d %B %y") 
template_path<- "src/create_report/template_excel.xlsx"
dq_data_long<- read_csv("data/dq_issues.csv")
pigeon_counts <-read_csv("data/pigeon_counts.csv")
squirrel_counts <-read_csv("data/squirrel_counts.csv")
vole_counts <-read_csv("data/vole_counts.csv")

car_park <- read_csv("data/car_park_waiting_list.csv")

variables<-list(names_vec=c("Pigeons",
                            "Squirrels",
                            "Voles"),
                dq_data = NA)

data<- list("data_pigeons" = pigeon_counts, 
            "data_squirrels" = squirrel_counts,
            "data_voles" =  vole_counts)

data_list_counts <-format_data(data, variables)

variables<-list(names_vec=c(""),
                dq_data = dq_data_long)

data<- list("car_park" = car_park)

data_list_car_park <-format_data(data, variables)

report_sheets <- list()

report_sheets$animal_counts$data <- data_list_counts
report_sheets$animal_counts$sheet_params$tab_name <- "animal_counts"
report_sheets$car_park$data <- data_list_car_park
report_sheets$car_park$sheet_params$tab_name <- "car_park"