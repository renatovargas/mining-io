# Input-Output Frameworks
# Renato Vargas

# In this file, we setup an input-analysis framework
# to regionalize with disaggregated labor data.

# Preamble
library(tidyverse)
library(haven)
library(readxl)

# Clean environment
rm(list = ls())

# Datasets

## Excel files
io_files <- c(
    #"data/BRA/BRA_WIOT2014.xlsx",
    "data/CHL/CHI_MIP_2022_Short_Framework.xlsx",
    "data/PER/PER_IO_Framework.xlsx",
    "data/ZAF/Final_Input_Output_tables_for_South_Africa_2014.xlsx"
    )

## Excel sheets
io_sheets <- c(
    #"IOT",                             # BRA
    "1",                               # CHL
    "MIP_Basico_2007_Corrientes_101_", # PER
    "Input - output table, 2014"       # ZAF
)

## Z Ranges
z_ranges <- c(
  # "E8:AY54",   # BRA
  "D12:O23",   # CHL
  "N12:DJ112", # PER
  "C4:AZ53"    # ZAF
)

## FD Ranges

fd_ranges <- c(
  # "AZ8:BD54",   # BRA
  "Q12:V23",    # CHL
  "DN12:DT112", # PER
  "BB4:BG53"    # ZAF
)



## Number of Countries
no_countries <- length(io_sheets)
iso3 <- tolower(c(
  # "BRA", 
  "CHL", 
  "PER", 
  "ZAF"
))

## Regional employment
employment_chl <- readRDS(file = "outputs/CHL/employment_chl.RDS")
employment_per <- readRDS(file = "outputs/PER/employment_per.RDS")
employment_zaf <- readRDS(file = "outputs/ZAF/employment_zaf.RDS")


for (i in 1:length(io_files)) {

# Basic Input-Output Components

## Interindustry transactions (Z)

assign(
  paste0("Z_", iso3[i]),
  as.matrix(
    read_excel(
      path  = io_files[i],
      sheet = io_sheets[i],
      range = z_ranges[i],
      col_names = F
    )
  )
)

## Final Demand by Components (FD)

assign(
  paste0("FD_", iso3[i]),
  as.matrix(
    read_excel(
      path  = io_files[i],
      sheet = io_sheets[i],
      range = fd_ranges[i],
      col_names = F
    )
  )
)


# Input-Output Model

# Output
assign(
  paste0("x_", iso3[i]),
  rowSums(get(paste0("Z_",  iso3[i]))) +
  rowSums(get(paste0("FD_", iso3[i])))
)

# Final Demand
assign(
  paste0("f_", iso3[i]),
  rowSums(get(paste0("FD_", iso3[i])))
)

# x hat
assign(
  paste0("x_hat_", iso3[i]),
  diag(get(paste0("x_",  iso3[i]))) 
)

# x hat inverse
assign(
  paste0("x_hat_inv_", iso3[i]),
  solve(get(paste0("x_hat_",  iso3[i]))) 
)

# Technical Coefficients Matrix

assign(
  paste0("A_", iso3[i]),
  get(paste0("Z_",  iso3[i])) %*%
    get(paste0("x_hat_inv_", iso3[i]))
)

# Identity matrix
assign(
  paste0("I_", iso3[i]),
  diag(dim(get(paste0("A_",  iso3[i])))[1]) 
)

# I - A
assign(
  paste0("IminusA_", iso3[i]),
  get(paste0("I_", iso3[i])) - 
    get(paste0("A_", iso3[i]))
)

# Leontief Matrix
assign(
  paste0("L_", iso3[i]),
  solve(get(paste0("IminusA_", iso3[i]))) 
)

}
  