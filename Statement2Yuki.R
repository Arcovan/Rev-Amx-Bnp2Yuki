# This program converts statements CSV from AMEX or BNP to a format to import in
# accounting system 
# For AMEX:
#   Download transactions via https://www.americanexpress.com/
#   Choose format : CSV and include transaction details
# 26-nov-2022

# ===== Define Functions --------------------------------------------------
CheckDocType <- function(x) {
  DocType<-"UNKNOWN"
  header<-readLines(x,skip=0, n = 1)
  NrOfColumns<-length(strsplit(header,"[;,]")[[1]])           # splits header in velden door seperator , of ;
  if ( length(grep("Product",header))!= 0) { DocType<-"REV"}  # soms zit er een extra kolom bij categorie
  if ( length(grep("Kaartlid",header))!= 0) { DocType<-"AMX"} # soms zit er een extra kolom bij categorie
  if ( length(grep("Uitvoeringsdatum",header)) != 0) { DocType<-"BNP"} # door de ; wordt deze ingelezen als 1 kolom
  if ( length(grep("Booking",header))!= 0) { DocType<-"JUB"} 
  message("Nr Columns RAW source: ", DocType ,"->", NrOfColumns)
  if (DocType =="UNKNOWN") { 
    message("Did not understand this header:\n", header, "\n")
    }
  return(DocType)
} # check document type based on keywords in header
CreateYukiDF <-function(x) {
  # x = nr of records to be created
  YukiDF <- data.frame(IBAN = character(length = x),
                       Valuta = character(length = x),
                       Afschrift = numeric(length = x),
                       Datum = character(length = x),
                       Rentedatum = character(length = x),
                       Tegenrekening = character(length = x),
                       Naam_tegenrekening = character(length = x),
                       Omschrijving = character(length = x),
                       Bedrag = integer(length = x))
  return(YukiDF) 
}  # data frame which will be used for export to CSV
CreateFeeDF <-function(x) {
  # x = nr of records to be created
  FeeDF <- data.frame(IBAN = character(length = x),
                       Valuta = character(length = x),
                       Afschrift = numeric(length = x),
                       Datum = character(length = x),
                       Rentedatum = character(length = x),
                       Tegenrekening = character(length = x),
                       Naam_tegenrekening = character(length = x),
                       Omschrijving = character(length = x),
                       Bedrag = integer(length = x))
  return(FeeDF) # return data frame which will be used for export
}   # Dataframe to split fee from lines

# ==== Read and Check import file ----------------------------------------
ifile <- file.choose()    #Select Import file Stop is incorrect
if (!(grepl(".csv",basename(ifile),ignore.case = TRUE))) {     # Grepl = grep logical
  stop("Please choose file with extension 'csv'.\n", call. = FALSE)
}
if (ifile == "") {
  stop("Empty File name [ifile]\n", call. = FALSE)
}
ofile <- sub(".csv", "-YukiR.csv", ifile, ignore.case = TRUE) # output file
setwd(dirname(ifile))     #set working directory to input directory where file is
message("Input file: ", basename(ifile), "\nOutput file: ", basename(ofile)) # display file name and output file with full dir name
message("Output file to directory: ", getwd())

# ==== SET Environment ====
options(OutDec = ".")     #set decimal point to "."

# Test Switch functie
DType<-CheckDocType(ifile)
switch (DType,
        "AMX" = message("Document: AMEX based on 'Kaartlid' in header"),
        "BNP" = message("Document: BNP-Fortis based on 'Uitvoeringsdatum' in header"),
        "REV" = message("Document: Revolut based on 'Product' in header"),
        "UNKNOWN" = stop("Unknown format: Relevant Keyword not found in header",call. = FALSE)
)
# ==== Process Input file and create output DATAFRAME ====
if (DType =="AMX") {
  AmexRaw <-read.csv(ifile, header = TRUE ,sep = "," , dec = ",", stringsAsFactors = FALSE)   #Reads field as factors or as characters
  NROF_Rawrecords <- nrow(AmexRaw)
  message("Aantal ingelezen records: ", NROF_Rawrecords)
  AmexRaw[, "Omschrijving"] <-
    gsub("(?<=[\\s])\\s*|^\\s+|\\s+$", "", AmexRaw[, "Omschrijving"], perl = TRUE) #strip all double spaces
  AmexRaw[, "Vermeld.op.uw.rekeningoverzicht.als"] <-
    gsub("(?<=[\\s])\\s*|^\\s+|\\s+$", "", AmexRaw[, "Vermeld.op.uw.rekeningoverzicht.als"], perl = TRUE) #strip all double spaces
  View(AmexRaw)
  # Create empty data frame
  YukiDF<-CreateYukiDF(NROF_Rawrecords)
  AfschriftNR<-substr(AmexRaw$Datum[which(AmexRaw$Omschrijving=="HARTELIJK BEDANKT VOOR UW BETALING")],1,2)
  if (length(AfschriftNR) == 0) { stop("Betaling niet gevonden",call. = FALSE) }
  # unlist(gregexpr(" ",YukiDF$Naam_tegenrekening))    # positie van Spaties
  YukiDF$IBAN<-"3753.8822.3759.008"     # nb Hard coded should be changed for other Cards
  YukiDF$Valuta<-"EUR"
  YukiDF$Datum<-format(as.Date(AmexRaw[,"Datum"],"%m/%d/%Y"),format ="%d-%m-%Y")
  YukiDF$Rentedatum<-YukiDF$Datum
  YukiDF$Afschrift<-paste(substr(YukiDF$Datum,7,11),AfschriftNR, sep="")
  YukiDF$Naam_tegenrekening<-gsub(",",".",gsub("^([^ ]* [^ ]*) .*$", "\\1", AmexRaw$Omschrijving)) # Copy all until 2nd SPACE, "paste(strsplit(s," ")[[1]][1:2],collapse = " ")" kan ook
  YukiDF$Omschrijving<-paste(substr(AmexRaw$Kaartlid,1,2),":",gsub(",",".",AmexRaw$Omschrijving))  # Add 1st 2 chars of member to Description and swap , for .
  YukiDF$Omschrijving[which(AmexRaw$Aanvullende.informatie!="")]<-paste(YukiDF$Omschrijving[which(AmexRaw$Aanvullende.informatie!="")],"#",gsub(",",".",AmexRaw$Aanvullende.informatie[which(AmexRaw$Aanvullende.informatie!="")]))
  YukiDF$Omschrijving<- gsub("\n","##",YukiDF$Omschrijving)  # Replace line feed in description for 2 hash tags
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
  #BNPRaw<-read.table(ifile, header= TRUE, sep = ";", quote = "", dec = ",",stringsAsFactors = FALSE)
  BNPRaw<-read.csv(ifile, header= TRUE, sep = ";", quote = "", dec = ",",stringsAsFactors = FALSE)
  NROF_Rawrecords <- nrow(BNPRaw)
  BNPRaw$Details <- gsub("(?<=[\\s])\\s*|^\\s+|\\s+$", "", BNPRaw$Details, perl = TRUE) #strip all double spaces
  View(BNPRaw)
  # Create empty data frame
  YukiDF <- CreateYukiDF(NROF_Rawrecords)
  YukiDF$IBAN<-BNPRaw$Rekeningnummer
  YukiDF$Valuta<-"EUR"
  YukiDF$Datum<-gsub("/","-",BNPRaw$Uitvoeringsdatum)
  YukiDF$Rentedatum<-gsub("/","-",BNPRaw$Valutadatum)
  #YukiDF$Afschrift<-gsub("-","",BNPRaw$Volgnummer)
  YukiDF$Afschrift<-paste(substr(BNPRaw$Uitvoeringsdatum ,7,11),substr(BNPRaw$Uitvoeringsdatum ,4,5),sep="")
  # substr(BNPRaw$Uitvoeringsdatum ,4,5)
  # substr(BNPRaw$Uitvoeringsdatum ,7,11)
  YukiDF$Naam_tegenrekening<-"" 
  YukiDF$Naam_tegenrekening[grep("BIC",BNPRaw$Details)]<-BNPRaw$Naam.van.de.tegenpartij[grep("BIC",BNPRaw$Details)]
  YukiDF$Omschrijving<-gsub("BANKREFERENTIE","Ref",gsub("VALUTADATUM ","ValDat",BNPRaw$Details))
  YukiDF$Naam_tegenrekening[length(BNPRaw$Mededeling)]
  YukiDF$Omschrijving[which(BNPRaw$Mededeling!="")]<-paste(YukiDF$Omschrijving[which(BNPRaw$Mededeling!="")], "/", BNPRaw$Mededeling[which(BNPRaw$Mededeling!="")])
  YukiDF$Bedrag<-BNPRaw$Bedrag
  YukiDF$Tegenrekening<-""
  YukiDF$Tegenrekening[grep("BIC",BNPRaw$Details)]<-BNPRaw$Tegenpartij[grep("BIC",BNPRaw$Details)]
  # split string in losse woorden en selecteer 9e + 10e woord als tegenpartij voor alle BETALING MET DEBETKAART
  YukiDF$Naam_tegenrekening[which(BNPRaw$Type.verrichting=="Kaartbetaling")]<-
    paste(sapply(strsplit(YukiDF$Omschrijving[which(BNPRaw$Type.verrichting=="Kaartbetaling")]," "),'[',9),
          sapply(strsplit(YukiDF$Omschrijving[which(BNPRaw$Type.verrichting=="Kaartbetaling")]," "),'[',10))
# Add name to Cash withdrawal ---------------------------------------------
  YukiDF$Naam_tegenrekening[which(BNPRaw$Type.verrichting=="Geldopname met kaart")]<-
    paste(sapply(strsplit(YukiDF$Omschrijving[which(BNPRaw$Type.verrichting=="Geldopname met kaart")]," "),'[',11))
  YukiDF$Naam_tegenrekening[which(YukiDF$Naam_tegenrekening=="4871")]<-"Karin van Leeuwen"
  }
if (DType =="REV") {
  #BNPRaw<-read.table(ifile, header= TRUE, sep = ";", quote = "", dec = ",",stringsAsFactors = FALSE)
  REVRaw<-read.csv(ifile, header= TRUE, sep = c(",",";"), quote = "", dec = ".",stringsAsFactors = FALSE)
  REVRaw$Details <- gsub("(?<=[\\s])\\s*|^\\s+|\\s+$", "", REVRaw$Description, perl = TRUE) #strip all double spaces
  View(REVRaw)
  NROF_Rawrecords <- nrow(REVRaw)
  aggregate.data.frame(REVRaw$Amount, list(REVRaw$Type) ,sum)
  # Create empty data frame
  YukiDF <- CreateYukiDF(NROF_Rawrecords) # create DF for output with same nr of rows
  YukiDF$IBAN<-paste("LT773250030128232439",REVRaw$Currency,sep = "") #hardcoded IBAN since not in FILE
  YukiDF$Valuta<-REVRaw$Currency
  YukiDF$Datum<-  format(as.Date(REVRaw$Completed.Date, "%Y-%m-%d"),format ="%d-%m-%Y")
  YukiDF$Rentedatum<-format(as.Date(REVRaw$Started.Date, "%Y-%m-%d"),format ="%d-%m-%Y")
  YukiDF$Afschrift<-paste(substr(YukiDF$Datum ,7,11),substr(YukiDF$Datum ,4,5),sep="")
  YukiDF$Omschrijving<-paste(REVRaw$Type,":",REVRaw$Description)
  YukiDF$Bedrag<-REVRaw$Amount
  YukiDF$Omschrijving[which(REVRaw$Fee!=0)]<-paste(YukiDF$Omschrijving[which(REVRaw$Fee!=0)],"FEE:",REVRaw$Fee[which(REVRaw$Fee!=0)])
  YukiDF$Naam_tegenrekening<-""
  YukiDF$Naam_tegenrekening[which(REVRaw$Type=="CASHBACK")]<-"Revolut"
  YukiDF$Naam_tegenrekening[which(REVRaw$Type=="CARD_PAYMENT")]<-REVRaw$Description[which(REVRaw$Type=="CARD_PAYMENT")]
  YukiDF$Naam_tegenrekening[which(REVRaw$Type=="TRANSFER")]<-REVRaw$Description[which(REVRaw$Type=="TRANSFER")]
  YukiDF$Naam_tegenrekening[which(substr(YukiDF$Naam_tegenrekening,1,3)=="To ")]<-paste(substr(YukiDF$Naam_tegenrekening[which(substr(YukiDF$Naam_tegenrekening,1,3)=="To ")],4,25))
  YukiDF$Naam_tegenrekening[which(REVRaw$Type=="TOPUP")]<-substr(REVRaw$Description[which(REVRaw$Type=="TOPUP")],14,35)
  FeeRecords<-nrow(REVRaw[REVRaw$Fee!=0,]) #number of records containing a Fee (transaction cost)
  if (FeeRecords>0) {
    FeeDF<-CreateFeeDF(FeeRecords)
    FeeDF$IBAN<-paste("LT773250030128232439",REVRaw$Currency[1],sep = "")
    FeeDF$Valuta <-REVRaw$Currency[1]
    FeeDF$Naam_tegenrekening<-"Revolut"
    FeeDF$Rentedatum<-format(as.Date(REVRaw$Completed.Date[REVRaw$Fee!=0],"%Y-%m-%d"),format ="%d-%m-%Y")
    FeeDF$Datum<-format(as.Date(REVRaw$Started.Date[REVRaw$Fee!=0],"%Y-%m-%d"),format ="%d-%m-%Y")
    FeeDF$Bedrag<--REVRaw$Fee[REVRaw$Fee!=0]
    FeeDF$Afschrift<-YukiDF$Afschrift[REVRaw$Fee!=0]
    FeeDF$Omschrijving<-paste("[FEE:]",REVRaw$Description[REVRaw$Fee!=0])
    View(FeeDF)
    message("Fee Split in seperate Records: ",FeeRecords, " Fee Amount: ",sum(FeeDF$Bedrag))
    YukiDF<-rbind(YukiDF,FeeDF)
    } #Split fee to separate records
} 
# ==== Post processing and Write output file ====
if (DType != "UNKNOWN") {
  YukiDF$Naam_tegenrekening<-gsub("\\s+", " ", YukiDF$Naam_tegenrekening)
  YukiDF$Omschrijving<-gsub("\\s+", " ", YukiDF$Omschrijving)
  View(YukiDF)
  Smry<-aggregate.data.frame(YukiDF$Bedrag, list(YukiDF$Naam_tegenrekening) ,sum)
  View(Smry)
  aggregate.data.frame(YukiDF$Bedrag, list(substr(YukiDF$Omschrijving,1,2)) ,sum)
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
  message("File Created from: ", DType, " Nof Raw Records: ", NROF_Rawrecords," Total Records created:",nrow(YukiDF), " Amount: ",sum(YukiDF$Bedrag))
} #s Recognised document so write output file
