# ============================================================
# scripts/02_week34_tables_and_text.R
# FULL STABLE VERSION (NO orgUnit_name dependency)
# Creates:
#  - outputs/priority_table_week34.csv
#  - outputs/week34_narratives.txt  (exactly 7 lines: 4 + 3)
#  - outputs/measles_lab_indicators.txt
#  - outputs/maternal_deaths_cumulative_total.txt
# ============================================================

suppressPackageStartupMessages({
  library(rprojroot)
  library(readr)
  library(dplyr)
  library(stringr)
  library(readxl)
  library(janitor)
  library(tidyr)
})

message("---- STEP 2: tables + text generation starting ----")

# ----------------------------
# Paths
# ----------------------------
ROOT <- rprojroot::find_root(rprojroot::is_rstudio_project)
DAT  <- file.path(ROOT, "data")
OUT  <- file.path(ROOT, "outputs")
dir.create(OUT, showWarnings = FALSE)

weekly_path  <- file.path(DAT, "weekly_data_by_province.csv")
measles_path <- file.path(DAT, "measles_lab.xlsx")
md_path      <- file.path(DAT, "Maternal deaths.xlsx")

if (!file.exists(weekly_path))  stop("Missing file: ", weekly_path)
if (!file.exists(measles_path)) stop("Missing file: ", measles_path)
if (!file.exists(md_path))      stop("Missing file: ", md_path)

message("✅ Files found in data/")

# ----------------------------
# Helpers
# ----------------------------
clean_lines <- function(x) {
  x <- as.character(x)
  x <- str_replace_all(x, "\u00A0", " ")
  x <- str_squish(x)
  x <- x[x != "" & !is.na(x)]
  x <- str_replace(x, "^•\\s*", "")
  x <- str_replace(x, "^[-*]\\s*", "")
  x
}

fmt_int <- function(x) {
  x <- suppressWarnings(as.integer(round(as.numeric(x), 0)))
  x[is.na(x)] <- 0L
  formatC(x, format = "d", big.mark = ",")
}

# ----------------------------
# A) Priority table from weekly_data_by_province.csv
# ----------------------------
message("---- A) Loading weekly_data_by_province.csv ----")
weekly_raw <- read_csv(weekly_path, show_col_types = FALSE) |> clean_names()

if (!("period" %in% names(weekly_raw))) {
  stop("weekly_data_by_province.csv: after clean_names(), column 'period' not found. Available: ",
       paste(names(weekly_raw), collapse = ", "))
}

cn <- names(weekly_raw)

confirmed_cols <- cn[str_detect(cn, "confirmed$")]
suspected_cols <- cn[str_detect(cn, "suspected$")]
tested_cols    <- cn[str_detect(cn, "sent_to_lab$")]

if (length(confirmed_cols) + length(suspected_cols) + length(tested_cols) == 0) {
  stop(
    "Could not detect any disease columns ending with confirmed/suspected/sent_to_lab.\n",
    "Check the column names in weekly_data_by_province.csv.\n",
    "First 30 columns are:\n- ",
    paste(head(cn, 30), collapse = "\n- ")
  )
}

extract_disease <- function(colname) {
  colname %>%
    str_replace("confirmed$", "") %>%
    str_replace("suspected$", "") %>%
    str_replace("sent_to_lab$", "") %>%
    str_replace("_+$", "") %>%
    str_replace_all("_", " ") %>%
    str_squish()
}

diseases <- sort(unique(c(
  vapply(confirmed_cols, extract_disease, character(1)),
  vapply(suspected_cols, extract_disease, character(1)),
  vapply(tested_cols, extract_disease, character(1))
)))
if (length(diseases) == 0) stop("No diseases extracted from weekly columns.")

disease_display <- function(x) {
  x %>%
    str_replace_all("\\bafp\\b", "AFP") %>%
    str_replace_all("\\bhiv\\b", "HIV") %>%
    str_to_title()
}

find_metric_col <- function(d, metric = c("confirmed","suspected","tested")) {
  metric <- match.arg(metric)
  d2 <- d %>% str_to_lower() %>% str_replace_all(" ", "_")
  
  if (metric == "confirmed") {
    hit <- confirmed_cols[str_detect(confirmed_cols, paste0("^", d2, ".*confirmed$"))]
    if (length(hit) == 0) hit <- confirmed_cols[str_detect(confirmed_cols, fixed(d2))]
    return(if (length(hit) == 0) NA_character_ else hit[1])
  }
  
  if (metric == "suspected") {
    hit <- suspected_cols[str_detect(suspected_cols, paste0("^", d2, ".*suspected$"))]
    if (length(hit) == 0) hit <- suspected_cols[str_detect(suspected_cols, fixed(d2))]
    return(if (length(hit) == 0) NA_character_ else hit[1])
  }
  
  hit <- tested_cols[str_detect(tested_cols, paste0("^", d2, ".*sent_to_lab$"))]
  if (length(hit) == 0) hit <- tested_cols[str_detect(tested_cols, fixed(d2))]
  return(if (length(hit) == 0) NA_character_ else hit[1])
}

# totals by period (sum over provinces etc.)
weekly_totals <- weekly_raw %>%
  group_by(period) %>%
  summarise(across(where(is.numeric), ~ sum(.x, na.rm = TRUE)), .groups = "drop")

# Week 34 label detection
week34_label <- "2024W34"
if (!(week34_label %in% weekly_totals$period)) {
  w34 <- weekly_totals$period[str_detect(weekly_totals$period, "W34$")]
  if (length(w34) == 0) {
    stop("Could not find Week 34 in period column. Example periods: ",
         paste(head(weekly_totals$period, 10), collapse = ", "))
  }
  week34_label <- w34[1]
}
message("✅ Using Week 34 period label: ", week34_label)

week34_row <- weekly_totals %>% filter(period == week34_label)

# Cumulative (week 1–34)
cum_rows <- weekly_totals %>%
  filter(str_detect(period, "^2024W")) %>%
  mutate(week_num = as.integer(str_extract(period, "(?<=W)\\d+"))) %>%
  filter(!is.na(week_num), week_num >= 1, week_num <= 34)

cum_sum <- cum_rows %>% summarise(across(where(is.numeric), ~ sum(.x, na.rm = TRUE)))

priority_table <- lapply(diseases, function(d) {
  col_c <- find_metric_col(d, "confirmed")
  col_s <- find_metric_col(d, "suspected")
  col_t <- find_metric_col(d, "tested")
  
  w34_c <- if (!is.na(col_c) && nrow(week34_row) == 1) as.numeric(week34_row[[col_c]]) else 0
  w34_s <- if (!is.na(col_s) && nrow(week34_row) == 1) as.numeric(week34_row[[col_s]]) else 0
  w34_t <- if (!is.na(col_t) && nrow(week34_row) == 1) as.numeric(week34_row[[col_t]]) else 0
  
  cum_c <- if (!is.na(col_c)) as.numeric(cum_sum[[col_c]]) else 0
  cum_s <- if (!is.na(col_s)) as.numeric(cum_sum[[col_s]]) else 0
  cum_t <- if (!is.na(col_t)) as.numeric(cum_sum[[col_t]]) else 0
  
  tibble(
    `Disease/Event/Condition` = disease_display(d),
    `Week 34 Confirmed`       = w34_c,
    `Week 34 Suspected`       = w34_s,
    `Week 34 Tested`          = w34_t,
    `Cumulative Confirmed`    = cum_c,
    `Cumulative Suspected`    = cum_s,
    `Cumulative Tested`       = cum_t
  )
}) %>% bind_rows() %>%
  arrange(`Disease/Event/Condition`)

write_csv(priority_table, file.path(OUT, "priority_table_week34.csv"))
message("✅ Wrote outputs/priority_table_week34.csv (rows = ", nrow(priority_table), ")")

# ----------------------------
# B) Narratives (EXACT 7 lines: 4 + 3)
# ----------------------------
message("---- B) Creating week34_narratives.txt ----")

# NOTE: You can change these disease names to match your priority_table names exactly
notifiable_list <- c("Cholera", "Anthrax", "Plague", "Monkeypox")
other_list      <- c("Measles", "Malaria", "Diarrhoea Non Bloody")

make_line <- function(disease_name) {
  row <- priority_table %>% filter(`Disease/Event/Condition` == disease_name)
  if (nrow(row) == 0) {
    return(paste0(disease_name, ": No cases reported in week 34 (or name not found in dataset)."))
  }
  
  paste0(
    disease_name, ": Week 34 reported ",
    fmt_int(row$`Week 34 Suspected`), " suspected, ",
    fmt_int(row$`Week 34 Tested`), " tested, and ",
    fmt_int(row$`Week 34 Confirmed`), " confirmed cases. ",
    "Cumulative (Week 1–34): ",
    fmt_int(row$`Cumulative Suspected`), " suspected, ",
    fmt_int(row$`Cumulative Tested`), " tested, ",
    fmt_int(row$`Cumulative Confirmed`), " confirmed."
  )
}

week34_narratives <- clean_lines(c(
  vapply(notifiable_list, make_line, character(1)),
  vapply(other_list, make_line, character(1))
))

# Guarantee exactly 7 lines (no blanks, no extras)
week34_narratives <- week34_narratives[seq_len(min(7, length(week34_narratives)))]
while (length(week34_narratives) < 7) {
  week34_narratives <- c(week34_narratives, "No additional narrative generated for this slot (check disease naming).")
}

writeLines(week34_narratives, file.path(OUT, "week34_narratives.txt"))
message("✅ Wrote outputs/week34_narratives.txt (lines = ", length(week34_narratives), ")")

# ----------------------------
# C) Measles lab indicators (coded IgM)
# ----------------------------
message("---- C) Creating measles_lab_indicators.txt ----")

