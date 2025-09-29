# Installazione pacchetti 
# install.packages("openxlsx")
# install.packages("devtools")
# devtools::install_github("omegahat/RDCOMClient")

# =============== Librerie =================
library(dplyr)
library(tidyr)
library(openxlsx)
library(RDCOMClient)

# DATA MODIFICABILE
data_report <- "28 settembre 2026"

# 1) Import e pulizia
VENDUTO_SETTORI_PV <- read.xlsx("Input_file") # SELEZIONARE PERCORSO DEL FILE DI INPUT PRESENTE NEL REPOSITORY

colonne <- c("kgroup_warehouse", "kg1_description", "kg2_description", 
             "kquantity_dec", "kqstock_dec", "ksale_dec", "krevenue_dec")

VENDUTO_SETTORI_PV <- VENDUTO_SETTORI_PV %>%  
  select(all_of(colonne)) %>% 
  slice(-1) %>% 
  rename(
    Punto_vendita = kgroup_warehouse,
    Settore1 = kg1_description,
    Settore2 = kg2_description, 
    Venduti = kquantity_dec,
    Giacenza = kqstock_dec,
    Ricavo = ksale_dec,
    Margine = krevenue_dec
  ) %>% 
  mutate(across(c(Venduti, Giacenza, Ricavo, Margine), as.numeric)) %>% 
  filter(
    Venduti > 0,
    Settore1 != "GENERICO (1)"
  )

# 2) Pivot: ogni PV diventa colonne Venduti/Giacenza/Ricavo/Margine
wide <- VENDUTO_SETTORI_PV %>%
  select(Settore1, Settore2, Punto_vendita, Venduti, Giacenza, Ricavo, Margine) %>%
  pivot_wider(
    names_from  = Punto_vendita,
    values_from = c(Venduti, Giacenza, Ricavo, Margine),
    names_glue  = "{Punto_vendita}_{.value}",
    values_fill = 0
  )

# 3) Calcolo MLVE per ciascun PV
pv_prefix <- sub("_Ricavo$", "", grep("_Ricavo$", names(wide), value = TRUE))
for (pv in pv_prefix) {
  ric <- paste0(pv, "_Ricavo")
  mar <- paste0(pv, "_Margine")
  mlv <- paste0(pv, "_MLVE")
  wide[[mlv]] <- ifelse(wide[[ric]] > 0, wide[[mar]] / wide[[ric]], NA_real_)
}

# 4) Calcolo totali riga per riga
wide <- wide %>%
  mutate(
    Tot_Venduti  = rowSums(across(ends_with("_Venduti")),  na.rm = TRUE),
    Tot_Giacenza = rowSums(across(ends_with("_Giacenza")), na.rm = TRUE),
    Tot_Ricavo   = rowSums(across(ends_with("_Ricavo")),   na.rm = TRUE),
    Tot_Margine  = rowSums(across(ends_with("_Margine")),  na.rm = TRUE),
    Tot_MLVE     = ifelse(Tot_Ricavo > 0, Tot_Margine / Tot_Ricavo, NA_real_)
  )

# 5) Funzione per aggiungere subtotali con MLVE corretti
aggiungi_subtotali <- function(df, pv_prefix) {
  df %>%
    group_split(Settore1) %>%
    lapply(function(blocco) {
      settore <- unique(blocco$Settore1)
      subtot <- blocco %>% summarise(across(where(is.numeric), sum, na.rm = TRUE))
      # Ricalcolo MLVE per ogni PV
      for (pv in pv_prefix) {
        ric <- paste0(pv, "_Ricavo")
        mar <- paste0(pv, "_Margine")
        mlv <- paste0(pv, "_MLVE")
        if (ric %in% names(subtot) && mar %in% names(subtot)) {
          subtot[[mlv]] <- ifelse(subtot[[ric]] > 0, subtot[[mar]] / subtot[[ric]], NA_real_)
        }
      }
      subtot$Tot_MLVE <- ifelse(subtot$Tot_Ricavo > 0, subtot$Tot_Margine / subtot$Tot_Ricavo, NA_real_)
      subtot$Settore1 <- paste("Subtotale", settore)
      subtot$Settore2 <- ""
      bind_rows(blocco, subtot)
    }) %>%
    bind_rows()
}

wide_subtot <- aggiungi_subtotali(wide, pv_prefix)

# 6) Riduzione colonne dopo i subtotali
wide_reduced <- wide_subtot %>%
  select(
    Settore1, Settore2,
    Tot_Venduti, Tot_Giacenza, Tot_Ricavo, Tot_Margine, Tot_MLVE,
    !!!setNames(
      object = unlist(lapply(pv_prefix, function(pv) c(paste0(pv, "_Venduti"), paste0(pv, "_MLVE")))),
      nm     = NULL
    )
  )

# 7) Riga totale complessivo (ricalcolando MLVE per ogni PV)
totale_complessivo <- wide %>%
  summarise(across(where(is.numeric), sum, na.rm = TRUE))

for (pv in pv_prefix) {
  ric <- paste0(pv, "_Ricavo")
  mar <- paste0(pv, "_Margine")
  mlv <- paste0(pv, "_MLVE")
  if (ric %in% names(totale_complessivo) && mar %in% names(totale_complessivo)) {
    totale_complessivo[[mlv]] <- ifelse(totale_complessivo[[ric]] > 0,
                                        totale_complessivo[[mar]] / totale_complessivo[[ric]],
                                        NA_real_)
  }
}

totale_complessivo$Tot_MLVE <- ifelse(totale_complessivo$Tot_Ricavo > 0,
                                      totale_complessivo$Tot_Margine / totale_complessivo$Tot_Ricavo,
                                      NA_real_)

totale_complessivo$Settore1 <- "Totale complessivo"
totale_complessivo$Settore2 <- ""

totale_complessivo <- totale_complessivo %>%
  select(
    Settore1, Settore2,
    Tot_Venduti, Tot_Giacenza, Tot_Ricavo, Tot_Margine, Tot_MLVE,
    !!!setNames(
      object = unlist(lapply(pv_prefix, function(pv) c(paste0(pv, "_Venduti"), paste0(pv, "_MLVE")))),
      nm     = NULL
    )
  )

# 8) Report finale
report_finale <- bind_rows(wide_reduced, totale_complessivo)
# 9) Esportazione su template
titolo_file <- paste0("VENDUTO PER SETTORE PUNTI VENDITA ", data_report)

template_path <- "Template_file.xlsx" # SELEZIONARE PERCORSO DEL TEMPLATE PRESENTE NEL REPOSITORY
desktop_path <- file.path(Sys.getenv("USERPROFILE"), "Desktop")
excel_out <- file.path(desktop_path, paste0(titolo_file, ".xlsx"))
pdf_out   <- file.path(desktop_path, paste0(titolo_file, ".pdf"))

wb <- loadWorkbook(template_path)

# Scrivo tabella da A7 senza intestazioni
writeData(wb, "Foglio3", report_finale, startRow = 7, startCol = 1, 
          rowNames = FALSE, colNames = FALSE)

# Scrivo data in A4
writeData(wb, "Foglio3", data_report, startRow = 4, startCol = 1, colNames = FALSE)

# ================= STILI ===================
stile_base_left <- createStyle(fontName="Calibri", fontSize=12,
                               border="TopBottomLeftRight", borderColour="black",
                               halign="left")
stile_base_center <- createStyle(fontName="Calibri", fontSize=12,
                                 border="TopBottomLeftRight", borderColour="black",
                                 halign="center")

stile_subtot <- createStyle(fontName="Calibri", fontSize=12,
                            border="TopBottomLeftRight", borderColour="black",
                            fgFill="#D9D9D9", textDecoration="bold")

stile_totale <- createStyle(fontName="Calibri", fontSize=12,
                            border="TopBottomLeftRight", borderColour="black",
                            fgFill="#FFC000", textDecoration="bold")

stile_int <- createStyle(numFmt="#,##0", fontName="Calibri", fontSize=12,
                         border="TopBottomLeftRight", borderColour="black",
                         halign="center")
stile_valuta <- createStyle(numFmt="#,##0.00 [$€-it-IT]", fontName="Calibri", fontSize=12,
                            border="TopBottomLeftRight", borderColour="black",
                            halign="center")
stile_percent <- createStyle(numFmt="0.00%", fontName="Calibri", fontSize=12,
                             border="TopBottomLeftRight", borderColour="black",
                             halign="center")

# ==================== APPLICO STILI ====================
nrighe <- nrow(report_finale)
ncolonne <- ncol(report_finale)

# Prime due colonne a sinistra
addStyle(wb, "Foglio3", style=stile_base_left,
         rows=7:(7+nrighe-1), cols=1:2, gridExpand=TRUE, stack=TRUE)

# Tutte le altre centrate
addStyle(wb, "Foglio3", style=stile_base_center,
         rows=7:(7+nrighe-1), cols=3:ncolonne, gridExpand=TRUE, stack=TRUE)

# Individuo colonne
venduti_cols <- grep("Venduti|Giacenza", names(report_finale))
ricavo_cols  <- grep("Ricavo|Margine", names(report_finale))
mlve_cols    <- grep("MLVE", names(report_finale))

# Applico formattazione numerica
if (length(venduti_cols) > 0) {
  addStyle(wb, "Foglio3", style=stile_int,
           rows=7:(7+nrighe-1), cols=venduti_cols, gridExpand=TRUE, stack=TRUE)
}
if (length(ricavo_cols) > 0) {
  addStyle(wb, "Foglio3", style=stile_valuta,
           rows=7:(7+nrighe-1), cols=ricavo_cols, gridExpand=TRUE, stack=TRUE)
}
if (length(mlve_cols) > 0) {
  addStyle(wb, "Foglio3", style=stile_percent,
           rows=7:(7+nrighe-1), cols=mlve_cols, gridExpand=TRUE, stack=TRUE)
}

# Subtotali
righe_subtot <- which(grepl("^Subtotale", report_finale$Settore1)) + 6
if (length(righe_subtot) > 0) {
  addStyle(wb, "Foglio3", style=stile_subtot,
           rows=righe_subtot, cols=1:ncolonne, gridExpand=TRUE, stack=TRUE)
}

# Totale complessivo
riga_totale <- which(report_finale$Settore1 == "Totale complessivo") + 6
if (length(riga_totale) > 0) {
  addStyle(wb, "Foglio3", style=stile_totale,
           rows=riga_totale, cols=1:ncolonne, gridExpand=TRUE, stack=TRUE)
}

# ==================== SALVATAGGIO ====================
saveWorkbook(wb, excel_out, overwrite = TRUE)

# Esporto anche in PDF sul Desktop
library(RDCOMClient)
excel_app <- COMCreate("Excel.Application")
excel_app[["Visible"]] <- FALSE
wb_com <- excel_app[["Workbooks"]]$Open(excel_out)
wb_com$ExportAsFixedFormat(0, pdf_out)  # 0 = xlTypePDF
wb_com$Close(FALSE)
excel_app$Quit()
