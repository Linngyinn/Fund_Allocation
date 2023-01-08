library(readxl)
library(lubridate)
library(dplyr)
library(pivottabler)

raw_data <- read_excel("Z:/Sizing Capital/out2_15jul_7pm.xlsx",sheet="result")

colnames(raw_data)[colnames(raw_data) == "Trade Name"] <- "Trade_Name"
raw_data<-subset(raw_data, Trade_Name!="Total")
raw_data <- subset(raw_data, select = -c(total))

# AUM <- readline(prompt="Enter AUM Size (Boston, K240act, K2UCITS, KLF, MCMM, MonksHill, RVMaster, UCITS): ")
AUM <- "262 34 51 234 500 54 219 185"

if (grepl(",", AUM)){
  AUM <- gsub(", ", " ", AUM)
}

initial_col_len <- length(colnames(raw_data))
mapping <- data.frame(c(colnames(raw_data[, c(4:initial_col_len)])), c(strsplit(AUM, " ")))
colnames(mapping) <- c('Fund','Allocation')
mapping$Allocation <- as.numeric(as.character(mapping$Allocation))

raw_data[is.na(raw_data)] = 0

AUM <- strsplit(AUM, split = " ")[[1]]
AUM <-  as.numeric(AUM)

raw_data$Boston_tobe = AUM[1]/AUM[5] * raw_data$MCMM
raw_data$K240act_tobe = AUM[2]/AUM[5] * raw_data$MCMM
raw_data$K2UCITS_tobe = AUM[3]/AUM[5] * raw_data$MCMM
raw_data$KLF_tobe = AUM[4]/AUM[5] * raw_data$MCMM
raw_data$MCMM_tobe = AUM[5]/AUM[5] * raw_data$MCMM
raw_data$MonksHill_tobe = AUM[6]/AUM[5] * raw_data$MCMM
raw_data$RVMaster_tobe = AUM[7]/AUM[5] * raw_data$MCMM
raw_data$UCITS_tobe = AUM[8]/AUM[5] * raw_data$MCMM

raw_data$Boston_error = raw_data$Boston - raw_data$Boston_tobe
raw_data$K240act_error = raw_data$K240act - raw_data$K240act_tobe
raw_data$K2UCITS_error = raw_data$K2UCITS - raw_data$K2UCITS_tobe
raw_data$KLF_error = raw_data$KLF - raw_data$KLF_tobe
raw_data$MCMM_error = raw_data$MCMM - raw_data$MCMM_tobe
raw_data$MonksHill_error = raw_data$MonksHill - raw_data$MonksHill_tobe
raw_data$RVMaster_error = raw_data$RVMaster - raw_data$RVMaster_tobe
raw_data$UCITS_error = raw_data$UCITS - raw_data$UCITS_tobe

raw_data <- raw_data %>% select(-(Boston_tobe:UCITS_tobe))
raw_data <- data.frame(append(raw_data, list(Sep = ""), after =  initial_col_len))

# for any fund > 150 aum, if error < 1k, error = 0

for (fund in mapping$Fund[mapping$Allocation >= 150]){
  col = paste0(fund, "_error")
  raw_data[c(col)][(raw_data[c(col)] < 1000) & (raw_data[c(col)] > -1000)] = 0
}

# for any fund < 150 aum, if error * 20bp *10000 /aum * 1million < 10, error = 0

for (fund in mapping$Fund[mapping$Allocation < 150]){
  col = paste0(fund, "_error")
  raw_data[c(col)][(raw_data[c(col)] * 20 / mapping$Allocation[mapping$Fund == fund] * 1000000/ 10000 < 10) & 
                     (raw_data[c(col)] * 20 / mapping$Allocation[mapping$Fund == fund] * 1000000/ 10000 > -10)] = 0
}

library(openxlsx)
wb = createWorkbook()
negStyle <- createStyle(fontColour = "#FF0000")
posStyle <- createStyle(fontColour = "#000000")
options("openxlsx.numFmt" = "#,##0")
addWorksheet(wb,"result",gridLines = TRUE)

writeData(wb,"result",raw_data,withFilter = TRUE)
conditionalFormatting(wb, "result", cols=4:length(colnames(raw_data)),rows = (2:nrow(raw_data)+1), rule="<0", style = negStyle)
conditionalFormatting(wb,"result", cols=4:length(colnames(raw_data)),rows = (2:nrow(raw_data)+1), rule=">=0", style = posStyle)

saveWorkbook(wb,"Z:/Fund Allocation/To_Be_Allocated.xlsx",overwrite = TRUE)
