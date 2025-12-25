# =========================================================
# scripts/01_load_and_validate_data.R
# STEP 1: Load & validate all data sources
# =========================================================

suppressPackageStartupMessages({
  library(tidyverse)
  library(readxl)
  library(sf)
  library(janitor)
  library(stringr)
  library(rprojroot)
})

ROOT <- rprojroot::find_root(rprojroot::is_rstudio_project)
DATA_DIR <- file.path(ROOT, "data")
OUT_DIR  <- file.path(ROOT, "outputs")
SHP_DIR  <- file.path(DATA_DIR, "shapefiles")

dir.create(OUT_DIR, showWarnings = FALSE)

# ---- Unzip shapefiles if needed ----
zip_path <- file.path(DATA_DIR, "shapefiles.zip")
if (file.exists(zip_path) && !dir.exists(SHP_DIR)) {
  unzip(zip_path, exdir = SHP_DIR)
}
stopifnot(dir.exists(SHP_DIR))

# ---- Load datasets ----
maternal <- read_excel(file.path(DATA_DIR, "Maternal deaths.xlsx")) %>%
  clean_names()

measles <- read_excel(file.path(DATA_DIR, "measles_lab.xlsx")) %>%
  clean_names()

weekly <- readr::read_csv(file.path(DATA_DIR, "weekly_data_by_province.csv"),
                          show_col_types = FALSE) %>%
  clean_names()

# ---- Load shapefile (auto-detect first .shp) ----
shp_candidates <- list.files(SHP_DIR, pattern = "\\.shp$", full.names = TRUE, ignore.case = TRUE, recursive = TRUE)
stopifnot(length(shp_candidates) >= 1)

zambia_sf <- sf::st_read(shp_candidates[1], quiet = TRUE) %>%
  clean_names()

# ---- Save to outputs for downstream scripts (optional but clean) ----
saveRDS(
  list(maternal = maternal, measles = measles, weekly = weekly, zambia_sf = zambia_sf),
  file = file.path(OUT_DIR, "step1_loaded_data.rds")
)

message("✅ STEP 1 complete: data loaded; shapefiles extracted (if needed).")
