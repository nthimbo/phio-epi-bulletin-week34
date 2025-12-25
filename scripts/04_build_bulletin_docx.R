# ============================================================
# scripts/04_build_bulletin_docx.R
# Build DOCX bulletin with fixed aesthetics + sizing
# Update: thinner green/orange bars + uses stacked measles chart
# ============================================================

suppressPackageStartupMessages({
  library(rprojroot)
  library(readr)
  library(dplyr)
  library(stringr)
  library(officer)
})

# ----------------------------
# Paths
# ----------------------------
ROOT <- rprojroot::find_root(rprojroot::is_rstudio_project)
OUT  <- file.path(ROOT, "outputs")
REP  <- file.path(ROOT, "report")
TMP  <- file.path(OUT, "tmp_docx_assets")
dir.create(REP, showWarnings = FALSE, recursive = TRUE)
dir.create(TMP, showWarnings = FALSE, recursive = TRUE)

TEMPLATE <- file.path(REP, "template.docx")
if (!file.exists(TEMPLATE)) stop("❌ Missing report/template.docx")

priority_csv <- file.path(OUT, "priority_table_week34.csv")
narr_txt     <- file.path(OUT, "week34_narratives.txt")
measles_txt  <- file.path(OUT, "measles_lab_indicators.txt")

img_measles_path  <- file.path(OUT, "measles_lab_by_province.png")   # generated in script 03
img_md_cause_path <- file.path(OUT, "maternal_deaths_causes_week34.png")
img_md_map_path   <- file.path(OUT, "maternal_deaths_cumulative_map.png")
md_total_txt      <- file.path(OUT, "maternal_deaths_cumulative_total.txt")

needed <- c(priority_csv, narr_txt, measles_txt,
            img_measles_path, img_md_cause_path, img_md_map_path, md_total_txt)

missing <- needed[!file.exists(needed)]
if (length(missing) > 0) stop("❌ Missing required files:\n", paste(missing, collapse = "\n"))

priority_table <- read_csv(priority_csv, show_col_types = FALSE)
narratives     <- readLines(narr_txt, warn = FALSE)
measles_ind    <- readLines(measles_txt, warn = FALSE)
md_total_line  <- readLines(md_total_txt, warn = FALSE)

img_measles  <- normalizePath(img_measles_path,  mustWork = TRUE)
img_md_cause <- normalizePath(img_md_cause_path, mustWork = TRUE)
img_md_map   <- normalizePath(img_md_map_path,   mustWork = TRUE)

# ----------------------------
# Colors (exact)
# ----------------------------
COL_DARK_GREEN  <- "#125C2C"
COL_LIGHT_GREEN <- "#92D050"
COL_ORANGE      <- "#F79546"
COL_RED         <- "#C00000"

# ----------------------------
# US Letter sizing
# IMPORTANT: your template.docx must be US Letter
# ----------------------------
USABLE_W <- 7.4   # if margins are Normal, use 6.6 instead

# ✅ thinner bars than before (closer to example)
H_BANNER <- 0.48
H_GREEN  <- 0.34
H_ORANGE <- 0.30

# content heights (adjust if Word spills to extra page)
H_PRIORITY       <- 2.35
H_MEASLES_PANEL  <- 3.20
H_MATERNAL_PANEL <- 3.10

# ----------------------------
# Helpers
# ----------------------------
safe_get <- function(x, i, fallback = "Fill in manually.") {
  if (length(x) >= i && !is.na(x[i]) && nzchar(x[i])) x[i] else fallback
}

extract_number <- function(x) {
  n <- str_extract(paste(x, collapse = " "), "\\d+")
  ifelse(is.na(n), "XXX", n)
}

# Real red bullets (fallback safely if officer version doesn't support list.style)
add_red_bullet <- function(doc, text, color = COL_RED) {
  out <- tryCatch({
    body_add_fpar(
      doc,
      fpar(
        ftext(text, prop = fp_text(color = color, font.size = 11)),
        fp_p = fp_par(list.style = "List Bullet")
      )
    )
  }, error = function(e) {
    body_add_par(doc, text, style = "List Bullet")
  })
  out
}

# Make a colored bar PNG (stable for exact colours)
PAGE_W_IN <- 7.5
DPI <- 300
BANNER_W_PX <- round(PAGE_W_IN * DPI)

make_banner_png <- function(file, bg, left, center, right) {
  png(file, width = BANNER_W_PX, height = 180, res = DPI)
  par(mar = c(0, 0, 0, 0), xaxs = "i", yaxs = "i")
  plot.new()
  rect(0, 0, 1, 1, col = bg, border = NA)
  
  text(0.02, 0.5, left,   col = "white", adj = c(0, 0.5), cex = 1.1, font = 2)
  text(0.50, 0.5, center, col = "white", adj = c(0.5, 0.5), cex = 1.3, font = 2)
  text(0.98, 0.5, right,  col = "white", adj = c(1, 0.5), cex = 1.1, font = 2)
  
  dev.off()
}




# Priority table as PNG (keeps exact orange header + orange first column)
make_priority_table_png <- function(df, file, orange = COL_ORANGE) {
  
  df2 <- df[, c(
    "Disease/Event/Condition",
    "Week 34 Suspected", "Week 34 Tested", "Week 34 Confirmed",
    "Cumulative Suspected", "Cumulative Tested", "Cumulative Confirmed"
  )]
  
  colnames(df2) <- c("Disease/Event/Condition",
                     "Suspected", "Tested", "Confirmed",
                     "Suspected", "Tested", "Confirmed")
  
  nrows <- nrow(df2)
  ncols <- ncol(df2)
  
  png(file, width = 2400, height = 820, res = 300)
  par(mar = c(0,0,0,0))
  plot.new()
  plot.window(xlim = c(0,1), ylim = c(0,1))
  
  x0 <- 0.03; x1 <- 0.97
  y0 <- 0.06; y1 <- 0.94
  
  w_total <- x1 - x0
  w_disease <- 0.36 * w_total
  w_other   <- (w_total - w_disease) / (ncols - 1)
  col_w <- c(w_disease, rep(w_other, ncols - 1))
  
  h_total <- y1 - y0
  h_hdr1  <- 0.15 * h_total
  h_hdr2  <- 0.15 * h_total
  h_body  <- (h_total - h_hdr1 - h_hdr2) / nrows
  
  xs <- numeric(ncols + 1)
  xs[1] <- x0
  for (j in 1:ncols) xs[j+1] <- xs[j] + col_w[j]
  
  y_top <- y1
  y_hdr1_bottom <- y_top - h_hdr1
  y_hdr2_bottom <- y_hdr1_bottom - h_hdr2
  
  # Header row 1
  rect(xs[1], y_hdr1_bottom, xs[2], y_top, col = orange, border = "white", lwd = 2)
  rect(xs[2], y_hdr1_bottom, xs[5], y_top, col = orange, border = "white", lwd = 2)
  text((xs[2] + xs[5])/2, (y_top + y_hdr1_bottom)/2, "Week 34",
       col = "white", font = 2, cex = 1.0)
  rect(xs[5], y_hdr1_bottom, xs[8], y_top, col = orange, border = "white", lwd = 2)
  text((xs[5] + xs[8])/2, (y_top + y_hdr1_bottom)/2, "Week 1 to 34, Cumulative Total",
       col = "white", font = 2, cex = 0.95)
  
  # Header row 2
  for (j in 1:ncols) {
    rect(xs[j], y_hdr2_bottom, xs[j+1], y_hdr1_bottom, col = orange, border = "white", lwd = 2)
    text((xs[j] + xs[j+1])/2, (y_hdr1_bottom + y_hdr2_bottom)/2, colnames(df2)[j],
         col = "white", font = 2, cex = 0.90)
  }
  
  # Body
  y_curr_top <- y_hdr2_bottom
  for (i in 1:nrows) {
    y_next <- y_curr_top - h_body
    for (j in 1:ncols) {
      cell_col <- if (j == 1) orange else "white"
      rect(xs[j], y_next, xs[j+1], y_curr_top, col = cell_col, border = "gray85", lwd = 1.5)
      val <- as.character(df2[i, j])
      tcol <- if (j == 1) "white" else "black"
      xpos <- if (j == 1) xs[j] + 0.01 else (xs[j] + xs[j+1])/2
      adjx <- if (j == 1) 0 else 0.5
      text(xpos, (y_curr_top + y_next)/2, val, col = tcol, cex = 0.85, adj = adjx)
    }
    y_curr_top <- y_next
  }
  
  dev.off()
  file
}

# Combine two PNGs side-by-side (requires png package)
combine_side_by_side <- function(left_png, right_png, out_png, gap_px = 30) {
  if (!requireNamespace("png", quietly = TRUE)) {
    install.packages("png", repos = "https://cloud.r-project.org")
  }
  library(png)
  
  L <- png::readPNG(left_png)
  R <- png::readPNG(right_png)
  
  to_rgba <- function(img) {
    d <- dim(img)
    if (length(d) == 2) {
      rgba <- array(1, dim = c(d[1], d[2], 4))
      rgba[,,1] <- img; rgba[,,2] <- img; rgba[,,3] <- img; rgba[,,4] <- 1
      return(rgba)
    }
    if (length(d) == 3 && d[3] == 3) {
      rgba <- array(1, dim = c(d[1], d[2], 4))
      rgba[,,1:3] <- img
      rgba[,,4] <- 1
      return(rgba)
    }
    if (length(d) == 3 && d[3] == 4) return(img)
    stop("Unsupported image format.")
  }
  
  L <- to_rgba(L)
  R <- to_rgba(R)
  
  h <- max(dim(L)[1], dim(R)[1])
  w <- dim(L)[2] + gap_px + dim(R)[2]
  canvas <- array(1, dim = c(h, w, 4))
  
  canvas[1:dim(L)[1], 1:dim(L)[2], ] <- L
  start_x <- dim(L)[2] + gap_px + 1
  end_x   <- start_x + dim(R)[2] - 1
  canvas[1:dim(R)[1], start_x:end_x, ] <- R
  
  png(out_png, width = w, height = h, res = 300)
  par(mar = c(0,0,0,0))
  plot.new()
  rasterImage(canvas, 0, 0, 1, 1)
  dev.off()
  
  out_png
}

make_bullets_png <- function(file, bullets, w = 1500, h = 980, wrap = 68) {
  png(file, width = w, height = h, res = 300)
  par(mar = c(0,0,0,0))
  plot.new()
  plot.window(xlim = c(0,1), ylim = c(0,1))
  
  y <- 0.95
  for (b in bullets) {
    lines <- strwrap(b, width = wrap)
    for (ln in lines) {
      text(0.05, y, paste0("\u2022 ", ln), adj = 0, cex = 0.92, col = "black")
      y <- y - 0.06
    }
    y <- y - 0.02
  }
  dev.off()
  file
}

# ============================================================
# Build PNG assets
# ============================================================

banner_png <- file.path(TMP, "banner_page.png")
make_banner_png(
  banner_png,
  bg = COL_DARK_GREEN,
  left = "Week 34",
  center = "Epidemiological Bulletin",
  right = "19th - 25th, Aug, 2024"
)





summary_png <- make_bar_png(file.path(TMP, "summary.png"), COL_LIGHT_GREEN,
                            center="Summary", text_col="black", w=2400, h=130, cex=1.10)

priority_bar_png <- make_bar_png(file.path(TMP, "priority_bar.png"), COL_LIGHT_GREEN,
                                 center="Summary Report Priority Diseases, Conditions and Events",
                                 text_col="black", w=2400, h=135, cex=0.98)

vpd_png <- make_bar_png(file.path(TMP, "vpd.png"), COL_LIGHT_GREEN,
                        center="Summary of VPD Surveillance Indicators",
                        text_col="black", w=2400, h=135, cex=1.00)

measles_orange_png <- make_bar_png(file.path(TMP, "measles_orange.png"), COL_ORANGE,
                                   center="Measles Laboratory Test Results by Province",
                                   text_col="white", w=2400, h=125, cex=1.00)

md_green_png <- make_bar_png(file.path(TMP, "md_green.png"), COL_LIGHT_GREEN,
                             center="Maternal Deaths", text_col="black", w=2400, h=135, cex=1.05)

md_orange_png <- make_bar_png(file.path(TMP, "md_orange.png"), COL_ORANGE,
                              left="Causes of maternal death (Week 34, n=15)",
                              right="Cumulative distribution of maternal deaths (2024) by province",
                              text_col="white", w=2400, h=125, cex=0.90)

priority_png <- make_priority_table_png(priority_table, file.path(TMP, "priority_table.png"))

# Measles panel = bullets (left) + stacked chart (right)
measles_total_sus <- priority_table %>%
  filter(`Disease/Event/Condition` == "Measles") %>%
  pull(`Cumulative Suspected`)
if (length(measles_total_sus) == 0 || is.na(measles_total_sus[1])) measles_total_sus <- "XXX"

measles_bullets <- c(
  paste0("The country has recorded a total of ", measles_total_sus[1], " suspected measles cases in 2024."),
  safe_get(measles_ind, 1, "Fill in measles indicator line 1."),
  safe_get(measles_ind, 2, "Fill in measles indicator line 2."),
  "The country’s non-measles febrile rash rate now stands at XXX per 100,000."
)

measles_left_png <- make_bullets_png(file.path(TMP, "measles_left.png"), measles_bullets)
measles_panel_png <- file.path(TMP, "measles_panel.png")
combine_side_by_side(measles_left_png, img_measles, measles_panel_png, gap_px = 30)

# Maternal panel = cause chart + map
maternal_panel_png <- file.path(TMP, "maternal_panel.png")
combine_side_by_side(img_md_cause, img_md_map, maternal_panel_png, gap_px = 30)

md_total_num <- extract_number(md_total_line)

# ============================================================
# Build DOCX
# ============================================================

doc <- read_docx(TEMPLATE)

# PAGE 1
doc <- body_add_img(doc, src = banner_png, width = 7.5, height = 0.6, style = "Normal")

doc <- body_add_img(doc, src = summary_png, width = USABLE_W, height = H_GREEN)

doc <- body_add_par(doc, "Key Highlights for Decision-Makers", style = "heading 3")
doc <- add_red_bullet(doc, "Fill in manually (2–3 bullet points).")
doc <- add_red_bullet(doc, "Fill in manually (2–3 bullet points).")
doc <- add_red_bullet(doc, "Fill in manually (2–3 bullet points).")

doc <- body_add_par(doc, "Immediately Notifiable Diseases and Events", style = "heading 3")
for (i in 1:4) doc <- body_add_par(doc, safe_get(narratives, i), style = "List Bullet")

doc <- body_add_par(doc, "Other Diseases and Events", style = "heading 3")
for (i in 6:8) doc <- body_add_par(doc, safe_get(narratives, i), style = "List Bullet")

doc <- body_add_img(doc, src = priority_bar_png, width = USABLE_W, height = H_GREEN)
doc <- body_add_img(doc, src = priority_png, width = USABLE_W, height = H_PRIORITY)

doc <- body_add_break(doc)

# PAGE 2
doc <- body_add_img(doc, src = vpd_png, width = USABLE_W, height = H_GREEN)
doc <- body_add_img(doc, src = measles_orange_png, width = USABLE_W, height = H_ORANGE)
doc <- body_add_img(doc, src = measles_panel_png, width = USABLE_W, height = H_MEASLES_PANEL)

doc <- body_add_img(doc, src = md_green_png, width = USABLE_W, height = H_GREEN)
doc <- body_add_img(doc, src = md_orange_png, width = USABLE_W, height = H_ORANGE)
doc <- body_add_img(doc, src = maternal_panel_png, width = USABLE_W, height = H_MATERNAL_PANEL)

doc <- body_add_par(doc, "The bar chart on the left summarizes the causes of deaths of 15 maternal deaths recorded in week 34.", style = "List Bullet")
doc <- body_add_par(doc, "Hypertensive disorder, Non-obstetric complications, and Obstetric haemorrhage continue to be the leading causes of maternal deaths this year.", style = "List Bullet")
doc <- body_add_par(doc, paste0("Cumulatively, in 2024, ", md_total_num, " maternal deaths have been recorded across the country, as depicted on the map."), style = "List Bullet")
doc <- body_add_par(doc, "Provinces with darker shades indicate those with a higher number of reported maternal deaths.", style = "List Bullet")

OUTPUT <- file.path(REP, "bulletin.docx")
print(doc, target = OUTPUT)

message("✅ Bulletin generated: ", OUTPUT)
message("ℹ If the green bars are still thicker than the example, reduce H_GREEN to 0.30.")
message("ℹ If you get an extra page, reduce H_MEASLES_PANEL and H_MATERNAL_PANEL by 0.10 each.")
