# --- Packages ---
library(readxl)
rm(list = ls())

# --- Config ---
dir_xls <- "data/BRA/Conta_da_Producao_2010_2022_xls"
files <- sprintf("%s/Tabela%d.xls", dir_xls, 1:33)
files <- files[file.exists(files)] # in case a file is missing

# Fixed data column names (we ignore in-sheet headers entirely)
cn <- c(
  "year",
  "valor_ano_anterior_milhao_rs",
  "indice_volume",
  "valor_a_precos_ano_anterior_milhao_rs",
  "indice_preco",
  "valor_a_preco_corrente_milhao_rs"
)

# Three blocks per sheet: data ranges + their transaction label cells
blocks <- list(
  list(data_rng = "A7:F19", label_cell = "A3"), # Output
  list(data_rng = "A27:F38", label_cell = "A22"), # Intermediate consumption
  list(data_rng = "A46:F57", label_cell = "A41") # Value added
)

# --- Runner ---
out <- list()
k <- 0L

for (f in files) {
  sh <- excel_sheets(f)
  sh <- sh[grepl("^Tabela", sh, ignore.case = TRUE)] # ignore "Sumário"
  if (!length(sh)) {
    next
  }

  for (s in sh) {
    # Fixed cells per sheet
    region <- read_excel(
      f,
      sheet = s,
      range = "A4",
      col_names = FALSE,
      col_types = "text"
    )[[1]]
    activity <- read_excel(
      f,
      sheet = s,
      range = "A5",
      col_names = FALSE,
      col_types = "text"
    )[[1]]

    # Three blocks
    for (b in blocks) {
      dat <- read_excel(
        f,
        sheet = s,
        range = b$data_rng,
        col_names = FALSE,
        col_types = "numeric",
        na = c("", "-", "…")
      )

      # Force exactly 6 data cols by position; drop all-NA rows
      if (ncol(dat) < 6 || nrow(dat) == 0) {
        next
      }
      dat <- dat[, 1:6, drop = FALSE]
      keep <- apply(dat, 1, function(x) any(!is.na(x)))
      dat <- dat[keep, , drop = FALSE]
      if (!nrow(dat)) {
        next
      }

      colnames(dat) <- cn

      trans <- read_excel(
        f,
        sheet = s,
        range = b$label_cell,
        col_names = FALSE,
        col_types = "text"
      )[[1]]

      # Append the 3 tags (now 9 columns total)
      dat$transaction <- trans
      dat$region_name <- region
      dat$economic_activity <- activity

      k <- k + 1L
      out[[k]] <- dat
    }
  }
}

# Bind everything (exactly 9 columns)
bra_prod <- if (length(out)) {
  do.call(rbind, out)
} else {
  setNames(
    data.frame(matrix(nrow = 0, ncol = 9)),
    c(cn, "transaction", "region_name", "economic_activity")
  )
}

# ---- Example: keep only 2014 values (optional) ----
# bra_2014 <- subset(bra_prod, year == 2014,
#                    select = c(region_name, economic_activity, transaction,
#                               valor_a_preco_corrente_milhao_rs))
