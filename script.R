# ++++++++++++++++++++++++++++++ ------------------------------------------
# Automatic validation with Excel and R -----------------------------------
# ++++++++++++++++++++++++++++++ ------------------------------------------

# Librerías utilizadas ----------------------------------------------------

##  Instalar y cargar paquetes
rm(list = ls())

if (!require("readr")) install.packages("readr") ; library(readr)   
if (!require("readxl")) install.packages("readxl") ; library(readxl)   
if (!require("haven")) install.packages("haven") ; library(haven)    
if (!require("tidyr")) install.packages("tidyr") ; library(tidyr)   
if (!require("openxlsx")) install.packages("openxlsx") ; library(openxlsx) 
if (!require("stringr")) install.packages("stringr") ; library(stringr)  
if (!require("rlang")) install.packages("rlang") ; library(rlang)
if (!require("purrr")) install.packages("purrr") ; library(purrr) 
if (!require("RCurl")) install.packages("RCurl") ; library(RCurl) 

# Ingest data and rules -----------------------------------------------------
cwd <- '/Users/icastillob/Downloads/Validación automática'
setwd(cwd)

df <- read.csv("https://raw.githubusercontent.com/ywchiu/riii/master/data/house-prices.csv")
rules <- read_excel('rules.xlsx')

# Edit data with wrong data
df[3, "Bedrooms"] <- 0
df[3, "Bathrooms"] <- 0
df[4, "Price"] <- 0
df[5, "SqFt"] <- -1


F_val <-  function(df, rules) {
  # ID lists
  val <- list()
  ids <- sort(unique(rules$ID))
  
  for (i in ids) {
    # Components
    x.tab <-  filter(rules, ID == i)
    r.val <-  str_trim(unlist(str_split(x.tab$VALIDATION[1], ",")), "both")
    r.ids <-  paste0("ID",x.tab$ID[1])
    
    # Evaluation of rules
    rule  <-  data_frame("Home" = df$Home , 
                         !!r.ids := ifelse(eval(parse(text = r.val), df) == FALSE, 0, 1))
    
    # Save results in a list
    val[[as.character(i)]] <- rule
  }
  return(val)  
}


# Validation automatic
Validation <- F_val(df = df, rules = rules)
ids <- sort(unique(rules$ID))

# Validaciones
for (i in ids) {
  if(i == 1){
    df_Validation <- Validation[[1]]}
  else if(i != 1){
    df_Validation <- left_join(df_Validation, Validation[[i]], by = c("Home"))
  }
}


# Final results ----------------------------------------------------------

i <- grepl("ID", names(df_Validation), fixed = TRUE)

df_Validation <- df_Validation %>% 
  mutate(ID_TOT = rowSums(.[i], na.rm = TRUE))


# Export ----------------------------------------------------------------

wb <- createWorkbook("VALIDATION_AUTOMATIC")
addWorksheet(wb, "Val_Automatic", gridLines = FALSE)

hd <- createStyle(fgFill = "#08519c", halign = "CENTER", textDecoration = "Bold", fontColour = "white")
writeData(wb,"Val_Automatic", df_Validation, startRow = 1, startCol = 1, headerStyle = hd, withFilter = TRUE)

red <- createStyle(fontColour = "red")
conditionalFormatting(wb, "Val_Automatic", cols = 5:42, rows = 2:5000, rule = ">0", style = red)

saveWorkbook(wb, "Results/VALIDATION_AUTOMATIC.xlsx", overwrite = TRUE)


