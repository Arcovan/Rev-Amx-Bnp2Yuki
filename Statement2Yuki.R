# This program converts statements CSV from AMEX, BNP,Revolut or Julius Bar to a format
# to import into accounting system (Yuki)
# For AMEX:
#   Download transactions via https://www.americanexpress.com/
#   Choose format : CSV and include transaction details
# For JuliusBaer: 
#   Download transactions; (selection does not work); import csv in G-sheets; delete old lines and save as CSV
# For Revolut: Go to browser version en download excel per currency, so seperate file per currency
# Second last Edit date: 8-may-2023 / Files is managed in Github
# added Saxo bank via XLSx
# added Mastercard BNP transaction recognistion
# added Julius USD account
# added recognize voided transaction in Revolut
# Last edit: 24-sep-2024 
# TODO: dir(dir_csv, pattern = "*.csv") process all file in directory
rm(list = ls())  # clear environment
cat("\014")      # clear console (CLS)
cat("Supported banks: Revolut, Amex, BNP and Julius Baer\n")
# ==== SET Environment ====
options(OutDec = ".")     #set decimal point to "."
options(scipen = 999)     # avoid scientific notation for large amounts
# Installeer het readxl-pakket indien nodig
# install.packages("readxl")
# Laad het readxl-pakket
library(readxl)
# ===== Define Functions --------------------------------------------------
CheckDocType <- function(x) {
  # x = ifile
  doc_types <- list(                # fields used to check which file format is presented
    "REV" = c("Product"),            # Revolut Bank
    "AMX" = c("Kaartlid"),           # American Express
    "BNP" = c("Uitvoeringsdatum"),   # BNP Paribas (Belgiu)
    "JUB" = c("Booking")             # Julius Baer
  )
  doc_nfields <- list(    # nr of fields in Header
    "REV" = 10,           # Revolut Bank
    "AMX" = 13,           # American Express sometimes 13 (category )
    "BNP" = 13,           # BNP Paribas (Belgiu)
    "JUB" = 6            # Julius Baer
  )
  # Read the first line of the file as the header
  # x<-ifile
  header <- tryCatch(readLines(x, n = 1), error = function(e) NULL)
  if (is.null(header)) {stop("Failed to read header from file: ", x)}
  DocType<-"UNKNOWN"
  header<-readLines(x,skip=0, n = 1)
  NrOfColumns<-length(strsplit(header,"[;,]")[[1]])           # splits header in velden door seperator , of ;
  for (type in names(doc_types)) {
    if (grepl(doc_types[type], header)) {
      DocType <- type
      break
    }
  }
  if (DocType =="UNKNOWN") { 
    message("Did not understand this header:\n", header, "\n")
    stop("Unknown format: Relevant Keyword not found in header",call. = FALSE)
  }
  if  (NrOfColumns!=doc_nfields[DocType]) {
    message("Header from ", DocType, " has incorrect number of columns:",NrOfColumns," should be ",doc_nfields[DocType],"\n","Found HEADER:\"",header,"\"")
    stop("Incorrect number of columns",call. = FALSE)    
  }
  message("INPUT format: ", DocType, " based on \"",doc_types[DocType], "\" in header and Nr Columns:", NrOfColumns)
  return(DocType)
} # check document type based on keywords in header
CreateYukiDF <-function(x) {
  # x = nr of records to be created
  YukiDF <- data.frame(IBAN = character(length = x),
                       Valuta = character(length = x),
                       Afschrift = integer(length = x),
                       Datum = character(length = x),
                       Rentedatum = character(length = x),
                       Tegenrekening = character(length = x),
                       Naam_tegenrekening = character(length = x),
                       Omschrijving = character(length = x),
                       Bedrag = numeric(length = x))
  return(YukiDF) 
}  # data frame which will be used for export to CSV
CreateFeeDF <-function(x) {
  # x = nr of records to be created
  FeeDF <- data.frame(IBAN = character(length = x),
                       Valuta = character(length = x),
                       Afschrift = integer(length = x),
                       Datum = character(length = x),
                       Rentedatum = character(length = x),
                       Tegenrekening = character(length = x),
                       Naam_tegenrekening = character(length = x),
                       Omschrijving = character(length = x),
                       Bedrag = numeric(length = x))
  return(FeeDF) # return data frame which will be used for export
}   # Data frame to split fee from lines

# ==== Read and Check import file ----------------------------------------
ifile <- file.choose()    #Select Import file Stop is incorrect
if (ifile == "") {stop("Empty File name [ifile]\n", call. = FALSE)}
if (file.info(ifile)$size <= 90) {
  stop("File is empty: ", ifile, " ", file.info(ifile)$size, " bytes.", call. = FALSE)
} # Smaller than 90 bytes
FileTypeXls<-grepl(".xls",basename(ifile),ignore.case = TRUE)  # Grepl = grep logical
FileTypeCSV<-grepl(".csv",basename(ifile),ignore.case = TRUE)
if (!(FileTypeCSV | FileTypeXls )) {    
  stop("Please choose file with extension 'csv' or 'xls'.\n", call. = FALSE)
}
setwd(dirname(ifile))     #set working directory to input directory where file is
if (FileTypeXls) {
  # Lees het XLSX-bestand in een data frame
  DType <- "SAX"
  ofile <- sub("\\.xlsx", "-YukiR\\.csv", ifile, ignore.case = TRUE) # output file \\ to avoid regex
  ofile <- sub("transactions", DType, ofile, ignore.case = TRUE) # output file \\ to avoid regex
}
if (FileTypeCSV) {
  DType<-CheckDocType(ifile) # if UNKNOWN then STOP
  ofile <- sub("\\.csv", "-YukiR\\.csv", ifile, ignore.case = TRUE) # output file \\ to avoid regex
}
if (DType=="BNP") {
  ofile <- sub("CSV_", "BNP_", ofile, ignore.case = TRUE) # Adjust name to identify Bank in Name 
}
message("Input file : ", basename(ifile), "\n", 
        "Output file: ", basename(ofile)) # display file name and output file with full dir name
message("Output file to directory: ", getwd())
# ==== Process Input file and create output DATAFRAME ====
if (DType =="AMX") {
  AmexRaw <-read.csv(ifile, header = TRUE ,sep = "," , dec = ",", stringsAsFactors = FALSE)   #Reads field as factors or as characters
  if (ncol(AmexRaw)!=13) {stop("Aantal kolommen is: ", ncol(AmexRaw), "; Inclusief transactiedetails vergeten aan te vinken.",call. = FALSE)}
  NROF_Rawrecords <- nrow(AmexRaw)
  AmexRaw$Omschrijving <-gsub("\\s+", " ", AmexRaw$Omschrijving) # Replace all instances of double or more spaces with a single space
  AmexRaw$Vermeld.op.uw.rekeningoverzicht.als <-
    gsub("\\s+", " ", AmexRaw$Vermeld.op.uw.rekeningoverzicht.als) #replace all instances of double or more spaces with a single space
  View(AmexRaw)
  # Create empty data frame
  YukiDF<-CreateYukiDF(NROF_Rawrecords)
  YukiDF$IBAN<-"3753.8822.3759.008"     # nb Hard coded should be changed for other Cards
  YukiDF$Valuta<-"EUR"
  YukiDF$Datum<-format(as.Date(AmexRaw[,"Datum"],"%m/%d/%Y"),format ="%d-%m-%Y")
  YukiDF$Rentedatum<-YukiDF$Datum
  AfschriftNR<-substr(AmexRaw$Datum[which(AmexRaw$Omschrijving=="HARTELIJK BEDANKT VOOR UW BETALING")],1,2)
  if (length(AfschriftNR) == 0) { stop("Betaling niet gevonden",call. = FALSE) }
  YukiDF$Afschrift<-paste(substr(YukiDF$Datum,7,11),AfschriftNR, sep="") #Year + nr
  YukiDF$Naam_tegenrekening<-gsub(",",".",gsub("^([^ ]* [^ ]*) .*$", "\\1", AmexRaw$Omschrijving)) # Copy all until 2nd SPACE, "paste(strsplit(s," ")[[1]][1:2],collapse = " ")" kan ook
  YukiDF$Omschrijving<-paste(substr(AmexRaw$Kaartlid,1,2),":",gsub(",",".",AmexRaw$Omschrijving))  # Add 1st 2 chars of member to Description and swap , for .
  YukiDF$Omschrijving[which(AmexRaw$Aanvullende.informatie!="")]<-paste(YukiDF$Omschrijving[which(AmexRaw$Aanvullende.informatie!="")],"#",gsub(",",".",AmexRaw$Aanvullende.informatie[which(AmexRaw$Aanvullende.informatie!="")]))
  YukiDF$Bedrag<-AmexRaw$Bedrag*-1
  YukiDF$Naam_tegenrekening[which(AmexRaw$Omschrijving=="HARTELIJK BEDANKT VOOR UW BETALING")]<-"American Express"
  # ==== Additional Records created for Commission charged by AMEX ====
  # Disadvatage is original amount on statement cannot be found
  FeeRecords<-length(grep("Commission Amount:",AmexRaw$Aanvullende.informatie)) #number of records containing a Fee (transaction cost)
  if (FeeRecords>0) {
    FeeDF<-CreateFeeDF(FeeRecords)
    FeeDF$IBAN<-YukiDF$IBAN[1]
    FeeDF$Valuta <-YukiDF$Valuta[1]
    FeeDF$Naam_tegenrekening<-"American Express"
    FeeDF$Datum<-YukiDF$Datum[grep("Commission Amount:",AmexRaw$Aanvullende.informatie)]
    FeeDF$Rentedatum<-FeeDF$Datum
    FeeDF$Afschrift<-YukiDF$Afschrift[grep("Commission Amount:",AmexRaw$Aanvullende.informatie)] #number of records containing a Fee (transaction cost)
    FeeDF$Omschrijving<-paste("[FEE:]",YukiDF$Omschrijving[grep("Commission Amount:",AmexRaw$Aanvullende.informatie)])
    FeeList<-unlist(strsplit(FeeDF$Omschrijving," "))
    FeeDF$Bedrag<--as.numeric(FeeList[which(FeeList=="Commission")+2])
    YukiDF$Bedrag[grep("Commission Amount:",AmexRaw$Aanvullende.informatie)]<-YukiDF$Bedrag[grep("Commission Amount:",AmexRaw$Aanvullende.informatie)]-FeeDF$Bedrag
    View(FeeDF)
    message("Fee Split in seperate Records: ",FeeRecords, " Fee Amount: ",sum(FeeDF$Bedrag))
    YukiDF<-rbind(YukiDF,FeeDF)
   }
}
if (DType =="BNP") {
  BNPRaw<-read.csv(ifile, header= TRUE, sep = ";", quote = "", dec = ",",stringsAsFactors = FALSE)
  NROF_Rawrecords <- nrow(BNPRaw)
  BNPRaw$Details <- gsub("\\s+", " ", BNPRaw$Details) # replace all instances of double or more spaces with a single space
  View(BNPRaw)
  # Create empty data frame
  YukiDF <- CreateYukiDF(NROF_Rawrecords)
  YukiDF$IBAN<-BNPRaw$Rekeningnummer
  YukiDF$Valuta<-BNPRaw$Valuta.rekening
  YukiDF$Datum<-format(as.Date(BNPRaw$Uitvoeringsdatum, format = "%d/%m/%Y"), "%d-%m-%Y")
  #YukiDF$Datum<-gsub("/","-",BNPRaw$Uitvoeringsdatum)  # replaced by the line above less efficient more concise
  YukiDF$Rentedatum<-gsub("/","-",BNPRaw$Valutadatum)
  YukiDF$Afschrift<-paste(substr(BNPRaw$Uitvoeringsdatum ,7,11),substr(BNPRaw$Uitvoeringsdatum ,4,5),sep="")
  YukiDF$Naam_tegenrekening<-"" 
  YukiDF$Naam_tegenrekening[grep("BIC",BNPRaw$Details)]<-BNPRaw$Naam.van.de.tegenpartij[grep("BIC",BNPRaw$Details)]
  YukiDF$Omschrijving<-gsub("BANKREFERENTIE","Ref",gsub("VALUTADATUM ","ValDat",BNPRaw$Details))
  YukiDF$Omschrijving[which(BNPRaw$Mededeling!="")]<-paste(YukiDF$Omschrijving[which(BNPRaw$Mededeling!="")], "/", BNPRaw$Mededeling[which(BNPRaw$Mededeling!="")])
  YukiDF$Bedrag<-BNPRaw$Bedrag
  YukiDF$Tegenrekening<-""
  YukiDF$Tegenrekening<-BNPRaw$Tegenpartij   #recently added since field seems properly filled
  # YukiDF$Tegenrekening[grep("BIC",BNPRaw$Details)]<-BNPRaw$Tegenpartij[grep("BIC",BNPRaw$Details)]
  # split string in losse woorden en selecteer 9e + 10e woord als tegenpartij voor alle BETALING MET DEBETKAART
  kaartbetaling_rows <- BNPRaw$Type.verrichting == "Kaartbetaling"  # logical indexing for rows containing "Kaartbetaling"
    YukiDF$Naam_tegenrekening[kaartbetaling_rows] <- paste(
    sapply(strsplit(YukiDF$Omschrijving[kaartbetaling_rows], " "), '[', 9),
    sapply(strsplit(YukiDF$Omschrijving[kaartbetaling_rows], " "), '[', 10)
  )
  # Add name to Cash withdrawal ---------------------------------------------
  YukiDF$Naam_tegenrekening[which(BNPRaw$Type.verrichting=="Geldopname met kaart")]<-
    paste(sapply(strsplit(YukiDF$Omschrijving[which(BNPRaw$Type.verrichting=="Geldopname met kaart")]," "),'[',13))
  YukiDF$Naam_tegenrekening[which(YukiDF$Naam_tegenrekening=="8706")]<-"Karin van Leeuwen"
  YukiDF$Naam_tegenrekening[which(YukiDF$Naam_tegenrekening=="7729")]<-"Arco van Nieuwland"
  # Add name to Direct Debet (Domiciliering) --------------------------------
  YukiDF$Naam_tegenrekening[which(BNPRaw$Type.verrichting=="Domiciliëring")]<-
    BNPRaw$Naam.van.de.tegenpartij[which(BNPRaw$Type.verrichting=="Domiciliëring")]
  YukiDF$Tegenrekening[which(BNPRaw$Type.verrichting=="Domiciliëring")]<-
    BNPRaw$Tegenpartij[which(BNPRaw$Type.verrichting=="Domiciliëring")]
  
  YukiDF$Naam_tegenrekening[which(BNPRaw$Type.verrichting=="Kredietkaartbetaling")]<-"MASTERCARD BNP"
  
  # Add name to Card transaction --------
  YukiDF$Omschrijving<-gsub("BETALING MET DEBETKAART NUMMER 4871 04XX XXXX 8706","BETALING MET KAART Nr 4871 xx 8706 (Karin)",YukiDF$Omschrijving)
  YukiDF$Omschrijving<-gsub("BETALING MET DEBETKAART NUMMER 4871 04XX XXXX 7729","BETALING MET KAART Nr 4871 xx 7729 (Arco)",YukiDF$Omschrijving)
  YukiDF$Omschrijving<-gsub("BETALING MET DEBETKAART NUMMER 4871 04XX XXXX 9449","BETALING MET KAART Nr 4871 xx 9449 (Luis)",YukiDF$Omschrijving)
  
  }
if (DType =="REV") {
  REVRaw<-read.csv(ifile, header= TRUE, sep = c(",",";"), quote = "", dec = ".",stringsAsFactors = FALSE)
  NROF_Rawrecords <- nrow(REVRaw)
  REVRaw$Details <- gsub("\\s+", " ", REVRaw$Description) # Replace all instances of double or more spaces with a single space
  curreny <-REVRaw$Currency[1]
  ofile <- sub("account-statement", paste("REV-", curreny, sep = "") , ofile, ignore.case = TRUE) # Adjust name to identify Bank in Name and curr
  ofile <- sub("_en-gb", "",ofile, ignore.case = TRUE)
  cat("Output file: ", basename(ofile), "\n") # display file name and output file with full dir name
  View(REVRaw)
  aggregate.data.frame(REVRaw$Amount, list(REVRaw$Type) ,sum)
  # Create empty data frame
  YukiDF <- CreateYukiDF(NROF_Rawrecords) # create DF for output with same nr of rows
  YukiDF$IBAN<-paste("LT773250030128232439",REVRaw$Currency,sep = "") #hardcoded IBAN since not in FILE
  YukiDF$Valuta<-REVRaw$Currency
  YukiDF$Rentedatum<-format(as.Date(REVRaw$Started.Date, "%Y-%m-%d"),format ="%d-%m-%Y")
  YukiDF$Datum<-format(as.Date(REVRaw$Completed.Date, "%Y-%m-%d"),format ="%d-%m-%Y")
  if (length(which(nchar(REVRaw$Completed.Date)==0))>0) {
    YukiDF$Datum[which(nchar(REVRaw$Completed.Date)==0)] <- YukiDF$Rentedatum[which(nchar(REVRaw$Completed.Date)==0)]
    cat("Date corrected:", YukiDF$Rentedatum[which(nchar(REVRaw$Completed.Date)==0)],"\n" )
  }
  YukiDF$Afschrift<-paste(substr(YukiDF$Datum ,7,11),substr(YukiDF$Datum ,4,5),sep="")
  YukiDF$Omschrijving<-paste(REVRaw$Type,":",REVRaw$Description)
  YukiDF$Bedrag<-0
  YukiDF$Bedrag<-REVRaw$Amount
  YukiDF$Bedrag[which(REVRaw$State!="COMPLETED")]<-0
  YukiDF$Omschrijving[which(REVRaw$Fee!=0)]<-paste(YukiDF$Omschrijving[which(REVRaw$Fee!=0)],"FEE:",REVRaw$Fee[which(REVRaw$Fee!=0)])
  YukiDF$Naam_tegenrekening<-""
  YukiDF$Naam_tegenrekening[which(REVRaw$Type=="EXCHANGE")]<-"Arco van Nieuwland"
  YukiDF$Naam_tegenrekening[which(REVRaw$Type=="CASHBACK")]<-"Revolut"
  YukiDF$Naam_tegenrekening[which(REVRaw$Type=="CARD_PAYMENT")]<-REVRaw$Description[which(REVRaw$Type=="CARD_PAYMENT")]
  YukiDF$Naam_tegenrekening[which(REVRaw$Type=="TRANSFER")]<-REVRaw$Description[which(REVRaw$Type=="TRANSFER")]
  YukiDF$Naam_tegenrekening[which(substr(YukiDF$Naam_tegenrekening,1,3)=="To ")]<-paste(substr(YukiDF$Naam_tegenrekening[which(substr(YukiDF$Naam_tegenrekening,1,3)=="To ")],4,25))
  YukiDF$Naam_tegenrekening[which(REVRaw$Type=="TOPUP")]<-substr(REVRaw$Description[which(REVRaw$Type=="TOPUP")],14,35)
  FeeRecords<-sum(REVRaw$Fee != 0 & REVRaw$State == "COMPLETED") #number of records containing a Fee (transaction cost)
  
  if (FeeRecords>0) {
    FeeDF<-CreateFeeDF(FeeRecords)
    FeeDF$IBAN<-paste("LT773250030128232439",REVRaw$Currency[1],sep = "")
    FeeDF$Valuta <-REVRaw$Currency[1]
    FeeDF$Naam_tegenrekening<-"Revolut"
    FeeDF$Rentedatum<-format(as.Date(REVRaw$Completed.Date[REVRaw$Fee!=0 & REVRaw$State=="COMPLETED" ],"%Y-%m-%d"),format ="%d-%m-%Y")
    FeeDF$Datum<-format(as.Date(REVRaw$Started.Date[REVRaw$Fee!=0 & REVRaw$State=="COMPLETED"],"%Y-%m-%d"),format ="%d-%m-%Y")
    FeeDF$Bedrag<--REVRaw$Fee[REVRaw$Fee!=0 & REVRaw$State=="COMPLETED"]
    FeeDF$Afschrift<-YukiDF$Afschrift[REVRaw$Fee!=0 & REVRaw$State=="COMPLETED"]
    FeeDF$Omschrijving<-paste("[FEE:]",REVRaw$Description[REVRaw$Fee!=0 & REVRaw$State=="COMPLETED"])
    View(FeeDF)
    message("Fee Split in seperate Records: ",FeeRecords, " Fee Amount:  ",sum(FeeDF$Bedrag))
    YukiDF<-rbind(YukiDF,FeeDF)
    } #Split fee to separate records
} 
if (DType =="JUB") {
  JUBRaw <-read.csv(ifile, header= TRUE, sep = ",",  dec = ".", stringsAsFactors = FALSE, quote = "\"")
  NROF_Rawrecords <- nrow(JUBRaw)
  JUBRaw$Booking.text <- gsub("\\s+", " ", JUBRaw$Booking.text) # Replace all instances of double or more spaces with a single space
  Currency<-substr(colnames(JUBRaw)[6],9,12)
  ofile <- sub("Statement", paste0(DType,"-",Currency), ofile, ignore.case = TRUE) # output file \\ to avoid regex
  colnames(JUBRaw)[6]<-"Balance"
  # after following operation RAW is not RAW anymore ;)
  # prepare amounts JB delivers funny format with quotes as thousand seperator
  JUBRaw$Credit[which(JUBRaw$Credit=="")]<-sub("","0",JUBRaw$Credit[which(JUBRaw$Credit=="")],ignore.case = TRUE) # so values can be converted to numeric
  JUBRaw$Debit[which(JUBRaw$Debit=="")]<-sub("","0",JUBRaw$Debit[which(JUBRaw$Debit=="")],ignore.case = TRUE)
  JUBRaw$Debit<-as.numeric(gsub("\'","", JUBRaw$Debit, ignore.case = TRUE))
  JUBRaw$Credit<-as.numeric(gsub("\'","", JUBRaw$Credit, ignore.case = TRUE))
  JUBRaw$Balance<-as.numeric(gsub("\'","", JUBRaw$Balance, ignore.case = TRUE))

  View(JUBRaw)
  # Create empty data frame
  YukiDF <- CreateYukiDF(NROF_Rawrecords) # Create empty data frame
  switch (Currency,
          "CHF" = YukiDF$IBAN<-"CH1108515052916502001",
          "EUR" = YukiDF$IBAN<-"CH8108515052916502002",
          "USD" = YukiDF$IBAN<-"CH2708515052916502004",
          " " = stop("Unknown currency: Relevant Keyword not found in header",call. = FALSE)
  )
  YukiDF$Valuta<-Currency
  YukiDF$Datum<-gsub("\\.","-",JUBRaw$Booking)
  YukiDF$Rentedatum<-gsub("\\.","-",JUBRaw$Value.date)
  YukiDF$Afschrift<-paste(substr(YukiDF$Datum ,7,11),substr(YukiDF$Datum ,4,5),sep="")
  YukiDF$Naam_tegenrekening<-"" 
  YukiDF$Omschrijving<-trimws(JUBRaw$Booking.text) # strip quote space at beginning and space quote at the end
  YukiDF$Naam_tegenrekening <- ifelse(substr(YukiDF$Omschrijving, 9, 18) == "WITHDRAWAL",
                                      paste(substr(YukiDF$Omschrijving, 43, 60)),
                                      YukiDF$Naam_tegenrekening)
  YukiDF$Naam_tegenrekening[grep("ALL-INCLUSIVE FEE",YukiDF$Omschrijving)]<-"Julius Baer"
  YukiDF$Naam_tegenrekening[grep("Interest",YukiDF$Omschrijving, ignore.case = TRUE)]<-"Julius Baer"
  
  YukiDF$Naam_tegenrekening[grep("PAYMENT TO",YukiDF$Omschrijving)]<-sub("PAYMENT TO BPS\\d{12}", "", YukiDF$Omschrijving[grep("PAYMENT TO",YukiDF$Omschrijving)])
  #YukiDF$Naam_tegenrekening[grep("PAYMENT TO",YukiDF$Omschrijving)]<-
  #paste(sapply(strsplit(YukiDF$Omschrijving[grep("PAYMENT TO",YukiDF$Omschrijving)]," "),'[',4),
  #      sapply(strsplit(YukiDF$Omschrijving[grep("PAYMENT TO",YukiDF$Omschrijving)]," "),'[',5))
  
  YukiDF$Naam_tegenrekening[grep("PAYMENT FROM",YukiDF$Omschrijving)]<-sub("PAYMENT FROM BPS\\d{12}", "", YukiDF$Omschrijving[grep("PAYMENT FROM",YukiDF$Omschrijving)])
  #YukiDF$Naam_tegenrekening[grep("PAYMENT FROM",YukiDF$Omschrijving)]<-
  #  paste(sapply(strsplit(YukiDF$Omschrijving[grep("PAYMENT FROM",YukiDF$Omschrijving)]," "),'[',4),
  #        sapply(strsplit(YukiDF$Omschrijving[grep("PAYMENT FROM",YukiDF$Omschrijving)]," "),'[',5),
  #        sapply(strsplit(YukiDF$Omschrijving[grep("PAYMENT FROM",YukiDF$Omschrijving)]," "),'[',6))
  YukiDF$Bedrag<-JUBRaw$Credit-JUBRaw$Debit
}
if (DType =="SAX") {
  SaxoRaw<-read_excel(ifile)
  View(SaxoRaw)
  SameHeader <- length (grep("Boekingsbedrag|Valutadatum|Transactiedatum|Instrument|Acties|Instrumentsymbool|Instrument ISIN|Instrumentvaluta",names(SaxoRaw)))==8
  stopifnot(SameHeader)
  
  NROF_Rawrecords <- nrow(SaxoRaw)
  
  YukiDF <- CreateYukiDF(NROF_Rawrecords)  # Create empty data frame
  SaxoCur <- SaxoRaw$Instrumentvaluta[grep("1", SaxoRaw$Omrekeningskoers)][1]
  if (is.na(SaxoCur)) {
    stop("no currency found [line300]")
    }
  switch (SaxoCur,
          "EUR" = YukiDF$IBAN <- "NL50BICK0253616360",      # Saxo works with customer ID 13782606 and Account : 69900/3616360EUR and IBAN
          "USD" = YukiDF$IBAN <- "NL90BICK3616360424",
          )
  YukiDF$Valuta<- SaxoCur
  YukiDF$Afschrift <- format(as.Date(SaxoRaw$Transactiedatum, "%d-%b-%Y"),format ="%Y%m")     # Afschrift (statement) equals yyyymm
  YukiDF$Rentedatum<- format(as.Date(SaxoRaw$Valutadatum, "%d-%b-%Y"),format ="%d-%m-%Y")
  YukiDF$Datum<- format(as.Date(SaxoRaw$Transactiedatum, "%d-%b-%Y"),format ="%d-%m-%Y")
  YukiDF$Tegenrekening<-""
  
  YukiDF$Naam_tegenrekening <- ifelse(is.na(SaxoRaw$Instrument),"",SaxoRaw$Instrument)
  SaxoRecords <- grep("Rente|Service fee",SaxoRaw$Acties)
  YukiDF$Naam_tegenrekening[SaxoRecords] <- "Saxo Bank"
  
  YukiDF$Omschrijving <- 
    ifelse(is.na(SaxoRaw$Instrumentsymbool),
           SaxoRaw$Acties,
           paste(SaxoRaw$Acties, "ISIN:", SaxoRaw$`Instrument ISIN`, "Sym:", SaxoRaw$Instrumentsymbool, SaxoRaw$Instrumentvaluta)
    )
  DividendRecords<-grep("Dividend",YukiDF$Omschrijving)
  YukiDF$Omschrijving[DividendRecords] <- paste(YukiDF$Omschrijving[DividendRecords], 
                                               "Ingehouden dividendbelasting 15%:", 
                                               as.character(round(SaxoRaw$Boekingsbedrag[DividendRecords] /0.85 * 0.15, 2))
                                               )
  YukiDF$Omschrijving[grep("Service fee",YukiDF$Omschrijving)]<- paste(YukiDF$Omschrijving[grep("Service fee",YukiDF$Omschrijving)], "Green 0,08% * Portefeuille /12")

  YukiDF$Bedrag<- SaxoRaw$Boekingsbedrag # before 2022 SaxoRaw$Aantal
}
# ==== Post processing and Write output file ====
if (DType != "UNKNOWN") {
  YukiDF$Naam_tegenrekening<-gsub("\\s+", " ", YukiDF$Naam_tegenrekening)
  YukiDF$Omschrijving<-gsub("\\s+", " ", YukiDF$Omschrijving)
  YukiDF$Omschrijving<-gsub("\n","##",YukiDF$Omschrijving)  # Replace line feed in description for 2 hash tags
  View(YukiDF)
  Smry<-aggregate.data.frame(YukiDF$Bedrag, list(YukiDF$IBAN, YukiDF$Naam_tegenrekening) ,sum)
  Smry <- Smry[order(Smry$Group.1, decreasing = FALSE), ]
  colnames(Smry) <- c("IBAN", "Account Name", "Total")  # Replace with your desired names
  
  View(Smry)
  if (exists("FeeDF") ) {message("Totale kosten:",DType,":",sum(FeeDF$Bedrag))}
  aggregate.data.frame(YukiDF$Bedrag, list(substr(YukiDF$Omschrijving,1,2)) ,sum)
  if (any(is.na(YukiDF))) {
    cat("Incorrect field found. CSV will not be saved.")
    stop("One of the value in dataframe YukiDF has value: NA")
  } # Stop
  ColumnNames <-
    c(
      "IBAN",
      "Valuta",
      "Afschrift",
      "Datum",
      "Rentedatum",
      "Tegenrekening",
      "Naam tegenrekening",
      "Omschrijving",
      "Bedrag"
    )
  write.table(
    YukiDF,
    file = ofile,   #Output file define at start
    quote = FALSE,
    sep = ";",
    dec = ".",
    row.names = FALSE,
    col.names = ColumnNames  # Vanwege de underscore in de header die geen spatie kan zijn
  ) 
  message("File Created from: ", DType, " Nof Raw Records: ", NROF_Rawrecords," Total Records created:",nrow(YukiDF), " Amount: ",formatC(sum(YukiDF$Bedrag) , format="f", big.mark = ",",digits=2))
  cat("Highest amount: ",formatC(min(YukiDF$Bedrag) , format="f", big.mark = ",",digits=2),
      "Supplier:", YukiDF$Naam_tegenrekening[which.min(YukiDF$Bedrag)])
} #Recognised document so write output file
