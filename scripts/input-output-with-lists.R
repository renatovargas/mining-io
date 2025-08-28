# ============ Setup ============
library(readxl)
library(openxlsx)  
library(haven)

rm(list = ls())

# ---- Inputs: files, sheets, ranges ----
io_files <- c(
  "data/CHL/CHI_MIP_2022_Short_Framework.xlsx",
  "data/PER/PER_IO_Framework.xlsx",
  "data/ZAF/Final_Input_Output_tables_for_South_Africa_2014.xlsx"
)
io_sheets <- c("1", "MIP_Basico_2007_Corrientes_101_", "Input - output table, 2014")
z_ranges  <- c("D12:O23", "N12:DJ112", "C4:AZ53")
fd_ranges <- c("Q12:V23", "DN12:DT112", "BB4:BG53")
iso3 <- tolower(c("CHL","PER","ZAF"))

# ---- Employment inputs (region_code, region_name, then sector cols) ----
employment_chl <- readRDS("outputs/CHL/employment_chl.RDS")
employment_per <- readRDS("outputs/PER/employment_per.RDS")
employment_zaf <- readRDS("outputs/ZAF/employment_zaf.RDS")

Emp <- list(chl = employment_chl, per = employment_per, zaf = employment_zaf)

# ---- Containers ----
Z  <- setNames(vector("list", length(iso3)), iso3)
FD <- setNames(vector("list", length(iso3)), iso3)
f  <- setNames(vector("list", length(iso3)), iso3)

# ---- Read national Z and FD; compute f ----
for (i in seq_along(io_files)) {
  cc <- iso3[i]
  Z[[cc]]  <- as.matrix(read_excel(io_files[i], sheet = io_sheets[i], range = z_ranges[i], col_names = FALSE))
  FD[[cc]] <- as.matrix(read_excel(io_files[i], sheet = io_sheets[i], range = fd_ranges[i], col_names = FALSE))
  f[[cc]]  <- rowSums(FD[[cc]], na.rm = TRUE)
}

# ---- Name alignment: use employment sector names as canonical order ----
for (cc in names(Z)) {
  sec_names <- names(Emp[[cc]])[-(1:2)]
  stopifnot(nrow(Z[[cc]]) == length(sec_names),
            ncol(Z[[cc]]) == length(sec_names),
            nrow(FD[[cc]]) == length(sec_names))
  rownames(Z[[cc]])  <- sec_names
  colnames(Z[[cc]])  <- sec_names
  rownames(FD[[cc]]) <- sec_names
  names(f[[cc]])     <- sec_names
}

# ---- Helpers ----
as_df <- function(M, first_col = "name") {
  if (is.null(dim(M))) {
    rn <- names(M); if (is.null(rn)) rn <- seq_along(M)
    out <- data.frame(rowname = rn, value = as.numeric(M), check.names = FALSE, row.names = NULL)
    names(out)[1] <- first_col
    out
  } else {
    rn <- rownames(M); if (is.null(rn)) rn <- seq_len(nrow(M))
    out <- data.frame(rowname = rn, M, check.names = FALSE, row.names = NULL)
    names(out)[1] <- first_col
    out
  }
}

build_Zf_mrio <- function(Z_cc, f_cc, Emp_cc) {
  rcodes <- Emp_cc$region_code
  if (inherits(rcodes, "haven_labelled")) rcodes <- as.character(as_factor(rcodes))
  if (!is.character(rcodes)) rcodes <- as.character(rcodes)

  sec_df <- Emp_cc[, -(1:2), drop = FALSE]
  for (j in seq_along(sec_df)) {
    x <- sec_df[[j]]
    if (inherits(x, "haven_labelled")) x <- as.numeric(x)
    sec_df[[j]] <- as.numeric(x)
  }

  sec_names <- rownames(Z_cc)
  stopifnot(!is.null(sec_names), ncol(Z_cc) == length(sec_names), length(f_cc) == length(sec_names))

  Emp_mat <- as.matrix(sec_df)
  Emp_mat <- Emp_mat[, sec_names, drop = FALSE]

  R <- nrow(Emp_mat); n <- length(sec_names)
  sec_tot <- colSums(Emp_mat, na.rm = TRUE)
  share   <- sweep(Emp_mat, 2, ifelse(sec_tot == 0, 1, sec_tot), "/")
  share[, sec_tot == 0] <- 0

  W_list  <- lapply(seq_len(R), function(r) diag(as.numeric(share[r, ]), nrow = n, ncol = n))
  W_stack <- do.call(rbind, W_list)                                  # (R*n x n)

  # Column allocator weighted by the same shares for each sector k
  K_shares <- matrix(0, nrow = n, ncol = R * n)
  for (k in seq_len(n)) {
    cols_k <- (seq_len(R) - 1L) * n + k
    K_shares[k, cols_k] <- share[, k]
  }

  Z_mrio <- W_stack %*% Z_cc %*% K_shares                           # (R*n x R*n)
  f_mrio <- W_stack %*% matrix(f_cc, ncol = 1)                      # (R*n x 1)

  rs_names <- as.vector(t(outer(rcodes, sec_names, paste, sep = "|")))
  rownames(Z_mrio) <- rs_names
  colnames(Z_mrio) <- rs_names
  rownames(f_mrio) <- rs_names
  colnames(f_mrio) <- "f"

  list(Z_mrio = Z_mrio, f_mrio = f_mrio, Emp_mat = Emp_mat)
}

compute_L_mrio <- function(Z_mrio, f_mrio) {
  x <- rowSums(Z_mrio, na.rm = TRUE) + as.numeric(f_mrio[, 1])
  x_inv <- diag(ifelse(x == 0, 0, 1 / x), nrow = length(x))
  dimnames(x_inv) <- list(colnames(Z_mrio), colnames(Z_mrio))
  A <- Z_mrio %*% x_inv
  dimnames(A) <- list(rownames(Z_mrio), colnames(Z_mrio))
  I <- diag(nrow(A)); dimnames(I) <- list(rownames(A), colnames(A))
  L <- tryCatch(solve(I - A),
                error = function(e) {
                  M <- matrix(NA_real_, nrow(A), ncol(A))
                  dimnames(M) <- list(rownames(A), colnames(A))
                  M
                })
  list(L = L, x = x)
}

# ---- Build MRIO, employment coefficients, and write Excel with What-If ----
for (cc in names(Z)) {
  mr   <- build_Zf_mrio(Z[[cc]], f[[cc]], Emp[[cc]])
  Zc   <- mr$Z_mrio
  fc   <- mr$f_mrio
  outL <- compute_L_mrio(Zc, fc)

  Lc <- outL$L
  xc <- outL$x

  # Employment coefficients e = E_rs / x_rs
  # Flatten Emp_mat to region-major, sector-minor to align with row order in MRIO
  E_rs <- as.numeric(t(mr$Emp_mat))
  names(E_rs) <- rownames(Zc)
  e_coeff <- ifelse(xc == 0, 0, E_rs / xc)
  names(e_coeff) <- rownames(Zc)

  # Workbook
  wb <- createWorkbook()

  addWorksheet(wb, "Z_mrio"); writeData(wb, "Z_mrio", as_df(Zc, "region|sector"))
  addWorksheet(wb, "f_mrio"); writeData(wb, "f_mrio", as_df(fc, "region|sector"))
  addWorksheet(wb, "L_mrio"); writeData(wb, "L_mrio", as_df(Lc, "region|sector"))
  addWorksheet(wb, "Emp_coeff"); writeData(wb, "Emp_coeff", as_df(e_coeff, "region|sector"))

  # Calibration 
  old_x  <- rowSums(Zc, na.rm = TRUE) + as.numeric(fc[,1])
  new_x  <- as.numeric(Lc %*% fc)
  resid  <- new_x - old_x
  relerr <- ifelse(old_x == 0, NA_real_, resid / old_x)
  addWorksheet(wb, "Calibration")
  writeData(wb, "Calibration",
            data.frame(`region|sector` = rownames(Zc),
                       old_x = old_x, new_x = new_x, resid = resid, rel_error = relerr,
                       check.names = FALSE))

  
