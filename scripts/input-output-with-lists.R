library(tidyverse)
library(readxl)

rm(list = ls())

# Files, sheets, ranges
io_files <- c(
  "data/CHL/CHI_MIP_2022_Short_Framework.xlsx",
  "data/PER/PER_IO_Framework.xlsx",
  "data/ZAF/Final_Input_Output_tables_for_South_Africa_2014.xlsx"
)
io_sheets <- c("1", "MIP_Basico_2007_Corrientes_101_", "Input - output table, 2014")
z_ranges  <- c("D12:O23", "N12:DJ112", "C4:AZ53")
fd_ranges <- c("Q12:V23", "DN12:DT112", "BB4:BG53")

iso3 <- tolower(c("CHL","PER","ZAF"))

## Regional employment
employment_chl <- readRDS(file = "outputs/CHL/employment_chl.RDS")
employment_per <- readRDS(file = "outputs/PER/employment_per.RDS")
employment_zaf <- readRDS(file = "outputs/ZAF/employment_zaf.RDS")

# Preallocate named lists
Z          <- setNames(vector("list", length(iso3)), iso3)
FD         <- setNames(vector("list", length(iso3)), iso3)
x          <- setNames(vector("list", length(iso3)), iso3)
f          <- setNames(vector("list", length(iso3)), iso3)
x_hat      <- setNames(vector("list", length(iso3)), iso3)
x_hat_inv  <- setNames(vector("list", length(iso3)), iso3)
A          <- setNames(vector("list", length(iso3)), iso3)
I_mat      <- setNames(vector("list", length(iso3)), iso3)
IminusA    <- setNames(vector("list", length(iso3)), iso3)
L          <- setNames(vector("list", length(iso3)), iso3)

for (i in seq_along(io_files)) {
  cc <- iso3[i]

  # Read matrices
  Z[[cc]] <- as.matrix(read_excel(io_files[i], sheet = io_sheets[i],
                                  range = z_ranges[i], col_names = FALSE))
  FD[[cc]] <- as.matrix(read_excel(io_files[i], sheet = io_sheets[i],
                                   range = fd_ranges[i], col_names = FALSE))

  # Core IO pieces
  x[[cc]]         <- rowSums(Z[[cc]],  na.rm = TRUE) + rowSums(FD[[cc]], na.rm = TRUE)
  f[[cc]]         <- rowSums(FD[[cc]], na.rm = TRUE)
  x_hat[[cc]]     <- diag(x[[cc]])
  x_hat_inv[[cc]] <- solve(x_hat[[cc]])
  A[[cc]]         <- Z[[cc]] %*% x_hat_inv[[cc]]
  I_mat[[cc]]     <- diag(nrow(A[[cc]]))
  IminusA[[cc]]   <- I_mat[[cc]] - A[[cc]]
  L[[cc]]         <- solve(IminusA[[cc]])
}

# Example access:
# L[["chl"]], x[["per"]], A[["zaf"]]

## Assume you already have lists: Z, FD, iso3 (from the refactor)
## And these employment data frames already loaded:
## employment_chl, employment_per, employment_zaf
## Each has: region_code, region_name, then sector columns that match Z’s sector order

Emp <- list(
  chl = employment_chl,
  per = employment_per,
  zaf = employment_zaf
)

# Helper to ensure sector column order in Emp matches Z row order
# If your Z matrices don’t have rownames yet, set them once here using the employment column names.
for (cc in names(Z)) {
  if (is.null(rownames(Z[[cc]]))) {
    # Take sector names from employment df (excluding first 2 columns)
    sec_names <- names(Emp[[cc]])[-(1:2)]
    stopifnot(nrow(Z[[cc]]) == length(sec_names))  # sanity check
    rownames(Z[[cc]]) <- sec_names
    colnames(Z[[cc]]) <- sec_names
    rownames(FD[[cc]]) <- sec_names
    # FD can be multiple components in columns; we leave its column names as-is
  }
}

build_Z_mrio <- function(Z_cc, Emp_cc) {
  # 1) Align sector names and get dimensions
  sec_names <- rownames(Z_cc)         # length n
  stopifnot(!is.null(sec_names), ncol(Z_cc) == length(sec_names))
  Emp_mat <- as.matrix(Emp_cc[, -(1:2)])
  Emp_mat <- Emp_mat[, sec_names, drop = FALSE]
  R <- nrow(Emp_mat); n <- length(sec_names)
  rcodes <- Emp_cc$region_code

  # 2) Regional production shares by sector: w[r,s] = emp[r,s] / sum_r emp[r,s]
  sec_tot <- colSums(Emp_mat, na.rm = TRUE)
  share   <- sweep(Emp_mat, 2, ifelse(sec_tot == 0, 1, sec_tot), "/")
  share[, sec_tot == 0] <- 0  # guard zero sectors

  # 3) Stack W_r diagonals into a tall operator: W_stack = rbind_r diag(share[r,])
  W_list  <- lapply(seq_len(R), function(r) diag(as.numeric(share[r, ]), nrow = n, ncol = n))
  W_stack <- do.call(rbind, W_list)  # (R*n x n)

  # 4) Replicate buyer columns for every destination region:
  #     K = (1_R^T ⊗ I_n)  -> (n x R*n)
  K <- kronecker(matrix(1, nrow = 1, ncol = R), diag(n))

  # 5) Build MRIO: Z_mrio = W_stack %*% Z_cc %*% K   -> (R*n x R*n)
  Z_mrio <- W_stack %*% Z_cc %*% K

  # 6) Nice row/col names: "RCODE|SECTOR"
  rownames(Z_mrio) <- as.vector(t(outer(rcodes, sec_names, paste, sep = "|")))
  colnames(Z_mrio) <- rownames(Z_mrio)

  Z_mrio
}


# Example for Chile:
Z_mrio_chl <- build_Z_mrio(Z[["chl"]], employment_chl)
dim(Z_mrio_chl)  # 192 192