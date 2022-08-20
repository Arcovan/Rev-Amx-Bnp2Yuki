# This program converts convert statements from AMEX or BNP to format to import in
# accounting system CSV provided by AMEX/BNP is converted to CSV
# ===== Define Functions =====
CheckDocType <- function(x) {
  DocType<-"UNKNOWN"
  header<-readLines(x,skip=0, n = 1)
  NrOfColumns<-length(strsplit(header,c(",",";"))[[1]]) #splits header in velden door seperator , of ;
  
  if ( length(grep("Product",header))!= 0) { DocType<-"REV"} # soms zit er een extra kolom bij categorie
  if ( length(grep("Kaartlid",header))!= 0) { DocType<-"AMX"} # soms zit er een extra kolom bij categorie
  if ( length(grep("Uitvoeringsdatum",header)) != 0) { DocType<-"BNP"} # door de ; wordt deze ingelezen als 1 kolom
  message("Aantal kolommen: ", NrOfColumns, "\nDocType:",DocType)
  return(DocType)
}
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
  return(YukiDF) # return data frame which will be used for export
}
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
} # Dataframe to split fee from lines

ifile <- file.choose()    #Select Import file Stop is incorrect
if (length(grep(".csv", ifile)) == 0) {
  stop("Please choose file with extension 'csv'.\n")
}
if (ifile == "") {
  stop("Empty File name [ifile]\n")
}
# ==== SET Environment ====
getOption("OutDec")       #check what decimal point is and return "." or ","
options(OutDec = ".")     #set decimal point to "."
setwd(dirname(ifile))     #set working directory to input directory where file is
ofile <- gsub(".csv", "-YukiR.csv", ifile)
message("Input file: ", ifile, "\nOutput file: ", ofile)
message("Output file to directory: ", getwd())
# Filename <- readline(prompt = "Welke filename?")

DType<-CheckDocType(ifile)
# Test Switch functie
switch (DType,
        "AMX" = message("Document: AMEX"),
        "BNP" = message("Document: BNP-Fortis"),
        "REV" = message("Document: Revolut")
)
# ==== Process Input file and create output DATAFRAME ====
if (DType =="AMX") {
  AmexRaw <-read.csv(ifile,header = TRUE ,stringsAsFactors = FALSE)   #Reads field as factors or as characters
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
  if (length(AfschriftNR) == 0) { stop("Betaling niet gevonden") }
  # unlist(gregexpr(" ",YukiDF$Naam_tegenrekening))    # positie van Spaties
  YukiDF$IBAN<-"3753.8822.3759.008"
  YukiDF$Valuta<-"EUR"
  YukiDF$Datum<-format(as.Date(AmexRaw[,"Datum"],"%m/%d/%Y"),format ="%d-%m-%Y")
  YukiDF$Rentedatum<-YukiDF$Datum
  YukiDF$Afschrift<-paste(substr(YukiDF$Datum,7,11),AfschriftNR, sep="")
  YukiDF$Naam_tegenrekening<-gsub(",",".",gsub("^([^ ]* [^ ]*) .*$", "\\1", AmexRaw$Omschrijving)) # CopY all until 2nd SPACE, "paste(strsplit(s," ")[[1]][1:2],collapse = " ")" kan ook
  YukiDF$Omschrijving<-paste(substr(AmexRaw$Kaartlid,1,2),":",gsub(",",".",AmexRaw$Omschrijving))
  YukiDF$Omschrijving[which(AmexRaw$Aanvullende.informatie!="")]<-paste(YukiDF$Omschrijving[which(AmexRaw$Aanvullende.informatie!="")],"#",gsub(",",".",AmexRaw$Aanvullende.informatie[which(AmexRaw$Aanvullende.informatie!="")]))
  YukiDF$Omschrijving<- gsub("\n","##",YukiDF$Omschrijving)  #Vervang line feed door twee hash tags
  YukiDF$Bedrag<-as.numeric(gsub(",",".",AmexRaw$Bedrag))*-1
  YukiDF$Naam_tegenrekening[which(AmexRaw$Omschrijving=="HARTELIJK BEDANKT VOOR UW BETALING")]<-"American Express"
  aggregate.data.frame(YukiDF$Bedrag, list( substr(YukiDF$Omschrijving,1,2)) ,sum)
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
  substr(BNPRaw$Uitvoeringsdatum ,4,5)
  substr(BNPRaw$Uitvoeringsdatum ,7,11)
  YukiDF$Naam_tegenrekening<-"" 
  YukiDF$Omschrijving<-gsub("BANKREFERENTIE","Ref",gsub("VALUTADATUM ","ValDat",BNPRaw$Details))
  YukiDF$Bedrag<-BNPRaw$Bedrag
  YukiDF$Tegenrekening<-""
  YukiDF$Tegenrekening[grep("BIC",BNPRaw$Details)]<-BNPRaw$Tegenpartij[grep("BIC",BNPRaw$Details)]
  # split string in losse woorden en selecteer 6e woord als tegenpartij voor alle BETALING MET DEBETKAART
  YukiDF$Naam_tegenrekening[grep("BETALING MET DEBETKAART", BNPRaw$Tegenpartij)]<-sapply(strsplit(BNPRaw$Details[grep("BETALING MET DEBETKAART", BNPRaw$Tegenpartij)]," "),'[',6)
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
    message("Fee")
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
  message("File Created from: ", DType, " Aantal: ", NROF_Rawrecords, " Bedrag: ",sum(YukiDF$Bedrag))
}  #schrijf gegegevns weg
