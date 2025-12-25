# ============================================================
# scripts/03_plots_and_maps.R
# FULL STABLE VERSION
# - Measles: stacked bar using CODED IgM Results
# - Maternal deaths:
#     * Causes chart: WEEK 34 % by "Short name"
#     * Map: cumulative 2024 by province (exclude TOTAL row)
# ============================================================

suppressPackageStartupMessages({
  library(rprojroot)
  library(readr)
  library(dplyr)
  library(stringr)
  library(ggplot2)
  library(readxl)
  library(janitor)
  library(sf)
  library(lubridate)
  library(tidyr)
  library(scales)
})

# ----------------------------
# Paths
# ----------------------------
ROOT <- rprojroot::find_root(rprojroot::is_rstudio_project)
DAT  <- file.path(ROOT, "data")
OUT  <- file.path(ROOT, "outputs")
SHP  <- file.path(DAT, "shapefiles")
dir.create(OUT, showWarnings = FALSE)

zip_path <- file.path(DAT, "shapefiles.zip")
if (!dir.exists(SHP) && file.exists(zip_path)) {
  dir.create(SHP, recursive = TRUE, showWarnings = FALSE)
  unzip(zip_path, exdir = SHP)
}

# ----------------------------
# Helpers
# ----------------------------
find_shp <- function(folder) {
  shp <- list.files(folder, pattern = "\\.shp$", recursive = TRUE, full.names = TRUE)
  if (length(shp) == 0) stop("No shapefile found in data/shapefiles.")
  shp[1]
}

norm_txt <- function(x) {
  x |> as.character() |> str_squish() |> str_to_upper()
}

pick_col <- function(df_names, patterns) {
  for (p in patterns) {
    hit <- df_names[str_detect(df_names, p)]
    if (length(hit) > 0) return(hit[1])
  }
  NA_character_
}

detect_province_col_by_values <- function(df) {
  provinces <- c(
    "CENTRAL","COPPERBELT","EASTERN","LUAPULA","LUSAKA",
    "MUCHINGA","NORTH WESTERN","NORTHERN","SOUTHERN","WESTERN",
    "NORTH-WESTERN","NORTHWESTERN"
  )
  nms <- names(df)
  
  scores <- sapply(nms, function(col) {
    vals <- df[[col]]
    if (is.null(vals)) return(0)
    v <- norm_txt(vals)
    v <- v[!is.na(v) & v != ""]
    if (length(v) == 0) return(0)
    sum(sapply(provinces, function(p) any(str_detect(v, fixed(p)))))
  })
  
  best <- nms[which.max(scores)]
  if (max(scores) == 0) NA_character_ else best
}

# ----------------------------
# Load data
# ----------------------------
weekly_path  <- file.path(DAT, "weekly_data_by_province.csv")
measles_path <- file.path(DAT, "measles_lab.xlsx")
md_path      <- file.path(DAT, "Maternal deaths.xlsx")

stopifnot(file.exists(weekly_path), file.exists(measles_path), file.exists(md_path))

weekly_raw  <- read_csv(weekly_path, show_col_types = FALSE) |> clean_names()
measles_raw <- read_excel(measles_path) |> clean_names()

md_summary  <- read_excel(md_path, sheet = "MD summary")  |> clean_names()
md_linelist <- read_excel(md_path, sheet = "MD linelist") |> clean_names()

# ============================================================
# A) MEASLES – STACKED BAR CHART (CODED IgM Results)
# Codes:
# 1 Positive, 2 Negative, 3 Indeterminate, 4 Not done, 5 Pending, 9 Unknown
# ============================================================

prov_col <- if ("province_of_residence" %in% names(measles_raw)) {
  "province_of_residence"
} else stop("Measles: province_of_residence column not found.")

igm_col <- if ("ig_m_results" %in% names(measles_raw)) {
  "ig_m_results"
} else stop("Measles: ig_m_results column not found.")

measles_plot_df <- measles_raw %>%
  transmute(
    province = .data[[prov_col]] %>% as.character() %>% str_squish(),
    igm_code = suppressWarnings(as.integer(.data[[igm_col]]))
  ) %>%
  mutate(
    igm_cat = case_when(
      igm_code == 1 ~ "Positive",
      igm_code == 2 ~ "Negative",
      igm_code == 3 ~ "Indeterminate",
      igm_code == 4 ~ "Not done",
      igm_code == 5 ~ "Pending",
      igm_code == 9 ~ "Unknown",
      TRUE ~ "Unknown"
    ),
    igm_cat = factor(
      igm_cat,
      levels = c("Indeterminate", "Negative", "Not done", "Pending", "Positive", "Unknown")
    )
  ) %>%
  filter(!is.na(province), province != "") %>%
  count(province, igm_cat, name = "n")

# If the example bulletin doesn’t show Unknown, uncomment:
# measles_plot_df <- measles_plot_df %>% filter(igm_cat != "Unknown")

p_measles <- ggplot(measles_plot_df, aes(x = province, y = n, fill = igm_cat)) +
  geom_col(width = 0.78) +
  labs(x = NULL, y = "Number of tests", fill = NULL) +
  theme_minimal(base_size = 11) +
  theme(
    legend.position = "right",
    axis.text.x = element_text(angle = 45, hjust = 1)
  )

ggsave(
  filename = file.path(OUT, "measles_lab_by_province.png"),
  plot = p_measles,
  width = 7.2, height = 3.4, dpi = 300
)

# ============================================================
# B) MATERNAL DEATHS – CAUSES CHART (WEEK 34) USING % (Short name)
# Requirement:
# - Y axis: Short name (cause)
# - X axis: % of deaths (frequency / total * 100)
# ============================================================

# Detect short name column
short_col <- if ("short_name" %in% names(md_linelist)) {
  "short_name"
} else {
  pick_col(names(md_linelist), c("short.*name"))
}

if (is.na(short_col)) {
  stop("MD linelist: could not detect 'Short name' column. Available columns:\n",
       paste(names(md_linelist), collapse = ", "))
}

# Detect date column for week
date_col <- pick_col(names(md_linelist), c("date_of_notification", "notif"))
if (is.na(date_col)) {
  stop("MD linelist: could not detect notification date column. Available columns:\n",
       paste(names(md_linelist), collapse = ", "))
}

md_week34 <- md_linelist %>%
  mutate(
    notif_date = suppressWarnings(as.Date(.data[[date_col]])),
    epi_week   = isoweek(notif_date),
    cause      = .data[[short_col]] %>% as.character() %>% str_squish()
  ) %>%
  filter(!is.na(epi_week), epi_week == 34, !is.na(cause), cause != "")

if (nrow(md_week34) == 0) {
  stop("MD linelist: no records found for Week 34 after filtering. Check date parsing / week definition.")
}

md_cause_pct <- md_week34 %>%
  count(cause, name = "n") %>%
  mutate(pct = (n / sum(n)) * 100) %>%
  arrange(pct)  # arrange for nice bottom-to-top ordering in coord_flip

p_md_cause <- ggplot(md_cause_pct, aes(x = pct, y = reorder(cause, pct))) +
  geom_col(width = 0.75) +
  labs(x = "% of maternal deaths", y = NULL) +
  scale_x_continuous(labels = function(x) paste0(x, "%")) +
  theme_minimal(base_size = 11)

ggsave(
  filename = file.path(OUT, "maternal_deaths_causes_week34.png"),
  plot = p_md_cause,
  width = 3.6, height = 2.8, dpi = 300
)

# ============================================================
# C) MATERNAL DEATHS – CUMULATIVE MAP (2024) FROM MD SUMMARY
# Requirement:
# - Use Province + maternal Death cumulative
# - Exclude TOTAL row
# ============================================================

prov_sum_col <- detect_province_col_by_values(md_summary)
if (is.na(prov_sum_col)) {
  stop(
    "MD summary: could not detect province column. Available columns:\n",
    paste(names(md_summary), collapse = ", ")
  )
}

cum_col <- if ("maternal_death_cumulative" %in% names(md_summary)) {
  "maternal_death_cumulative"
} else {
  pick_col(names(md_summary), c("maternal.*death.*cumulative", "cumulative", "cum"))
}

if (is.na(cum_col)) {
  stop(
    "MD summary: could not detect maternal death cumulative column. Available columns:\n",
    paste(names(md_summary), collapse = ", ")
  )
}

# Exclude TOTAL row (and any blank province rows)
md_map_df <- md_summary %>%
  mutate(
    prov_raw = .data[[prov_sum_col]] %>% as.character() %>% str_squish(),
    prov_key = norm_txt(prov_raw),
    cum_md   = suppressWarnings(as.numeric(.data[[cum_col]]))
  ) %>%
  filter(!is.na(prov_key), prov_key != "") %>%
  filter(!str_detect(prov_key, "TOTAL")) %>%     # remove totals row
  group_by(prov_key) %>%
  summarise(cum_md = sum(cum_md, na.rm = TRUE), .groups = "drop")

# Load shapefile + join
shp_file <- find_shp(SHP)
zmb_sf <- st_read(shp_file, quiet = TRUE) |> clean_names()

sf_prov_col <- pick_col(names(zmb_sf), c("^province$", "provname", "prov"))
if (is.na(sf_prov_col)) {
  stop(
    "Shapefile: could not detect province column. Available columns:\n",
    paste(names(zmb_sf), collapse = ", ")
  )
}

zmb_sf2 <- zmb_sf %>%
  mutate(prov_key = norm_txt(.data[[sf_prov_col]])) %>%
  left_join(md_map_df, by = "prov_key") %>%
  mutate(cum_md = replace_na(cum_md, 0))

p_md_map <- ggplot(zmb_sf2) +
  geom_sf(aes(fill = cum_md), linewidth = 0.2) +
  theme_void(base_size = 11) +
  labs(fill = "Maternal deaths\n(cumulative 2024)") +
  theme(legend.position = "right")

ggsave(
  filename = file.path(OUT, "maternal_deaths_cumulative_map.png"),
  plot = p_md_map,
  width = 3.6, height = 2.8, dpi = 300
)

message("✅ STEP 3 complete: measles plot + maternal deaths chart/map saved in outputs/")
