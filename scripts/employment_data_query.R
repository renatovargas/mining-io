# Preamble
library(tidyverse)
library(haven)
library(readxl)

# South Africa
# Datasets
lookup_table_zaf <- read_excel(
  "data/ZAF/ZAF_Employment.xlsx",
  sheet = "Sheet1") |>
  select(1,4,5)
  
qlfs_2014_4 <- read_sav("data/ZAF/QLFS 2014_4 Worker v1.0 SPSS.sav")

# Employment
# No provinces
test <- qlfs_2014_4 |> 
  group_by(Q43INDUSTRY, as_factor(Q43INDUSTRY)) |> 
  summarize(
    total = sum(Weight)
  )

# write.table(test, file = pipe("xclip -selection clipboard"), sep = "\t", row.names = FALSE)

# With provinces

test2 <- qlfs_2014_4 |> 
  filter(! is.na(Q43INDUSTRY),
         ! Q43INDUSTRY %in% c(10,20,30)) |> 
  left_join(
    lookup_table,
    join_by(Q43INDUSTRY == `Industry Code`)
  ) |> 
  group_by(Province, as_factor(Province), `IO Code`, `IO Activity`) |> 
  summarize(
    total = sum(Weight)
  ) |> 
  pivot_wider(
    id_cols = c(Province,`as_factor(Province)`),
    names_from = c(`IO Code`, `IO Activity`),
    values_from = total,
    names_sep = "_",
    names_sort = T,
    values_fill = 0
  )

# write.table(test2, file = pipe("xclip -selection clipboard"), sep = "\t", row.names = FALSE)

# Chile
lookup_table_chl <- read_excel(
  "data/CHL/Empleo_Rama_Actividad_Anual.xlsx",
  sheet = "Equivalencia") |> 
  select(! `Employees (thousands)`)

employment_chl <- read_excel(
  "data/CHL/Empleo_Rama_Actividad_Anual.xlsx",
  sheet = "Regional_Employment") |>
  mutate(region_code = str_extract(Región, "(?<=\\()[A-Z]{2}(?=\\))")) |> 
  select(! `Población ocupada (total)`) |> 
  pivot_longer(
    cols = c(3:24),
    names_to = "Actividad"
  ) |> 
  left_join(
    lookup_table_chl,
    join_by(Actividad == Activity)
  ) |> 
  filter(`Año` == 2022) |> 
  group_by(region_code,`Región`, `IO Code`, `IO Code String`,`IO Name`) |> 
  summarize(
    total = sum(value * 1000, na.rm = T)  # thousands of people
  ) |> 
  ungroup() |> 
  pivot_wider(
    id_cols = c(region_code,`Región`),
    names_from = c(`IO Code String`,`IO Name`),
    values_from = total,
    names_sep = "_",
    names_sort = T,
    values_fill = 0
  )

# write.table(employment_chl, file = pipe("xclip -selection clipboard"), sep = "\t", row.names = FALSE)


