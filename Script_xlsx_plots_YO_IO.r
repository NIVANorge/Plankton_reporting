######################################################################
##SCRIPT FOR CREATING OVERVIEW EXCEL FILES AND PLOTS FROM YO/IO DATA##
##																                                	##
##Sonja Kistenich, last updated 24.06.2024				            			##
##For Plankton Toolbox version 1.4.1							                	##
######################################################################

#Set working directory if necessary
setwd("C:/Users/SKI/R/R_Projects/Excel_modification")

#load libraries
library(readxl)
library(tidyverse)
library(reshape2)
library(openxlsx)
library(scales)
library(grid)
library(cowplot)

#ensure correct encoding of special characters
options(encoding = "UTF-8")			#for older R versions
Sys.setlocale(locale='no_NB.utf8') 	#for newer R versions

#import PTB report and chla file
file1 <- list.files(pattern = "rep_")
report1 <- read_xlsx(file1)
file2 <- list.files(pattern = "klf_")
report2 <- read_xlsx(file2, sheet = "Water chemistry")

rep_split <- split(report1, report1$Station_code)
my.list <- names(rep_split)
len <- length(my.list)



#_______________________run Script_xlsx.R for constructing abundance, carbon and net haul xlsx files_______________________

for (x in rep_split) {
    
	
	rep <- as.data.frame(x)
	rep <- rep %>% arrange(Station_code, Sample_date, Min_depth) %>% 
		rename(Abundance=Abundance_ind_l, Calculated_carbon=Calculated_carbon_ugC_l) %>% 
		mutate_at(vars(Abundance, Calculated_carbon), as.numeric) 
		
	sizeclasses <- read_xlsx("SizeClasses_2024_2.xlsx", col_types = "text")

	# Taxonomic corrections
	rep$Taxon_name_extended[rep$Taxon_name_extended == "Protoperidinium marie-lebouriae"] <- "Protoperidinium marielebouriae"
	rep$Taxon_name_extended[rep$Taxon_name_extended == "cf. Protoperidinium marie-lebouriae"] <- "cf. Protoperidinium marielebouriae"
	rep$Taxon_name_extended[rep$Taxon_name_extended == "Protoperidinium cf. marie-lebouriae"] <- "Protoperidinium cf. marielebouriae"


	# Retrieve size range, combine files and fill empty Taxon_class
	rep1 <- left_join(rep, sizeclasses, by = c("Taxon_name" = "Species","Size_class" = "SizeClass"))
	rep1$Taxon_class[is.na(rep1$Taxon_class) & rep1$Taxon_name_extended == "Ciliophora"] <- "Ciliophora"
	rep1$Taxon_class[is.na(rep1$Taxon_class) & rep1$Taxon_name_extended == "Cyanobacteria"] <- "Cyanobacteria"
	rep1$Taxon_class[is.na(rep1$Taxon_class) & rep1$Taxon_name_extended == "Haptophyta"] <- "Prymnesiophyceae"
	rep1$Taxon_class[is.na(rep1$Taxon_class) & rep1$Taxon_name_extended == "Solenicola setigera"] <- "Protozoa classes incertae sedis"
	rep1$Taxon_class[is.na(rep1$Taxon_class)] <- "Classes incertae sedis"
	rep1$Taxon_class[rep1$Taxon_class=="Unicells classes incertae sedis" | rep1$Taxon_class == "Flagellates classes incertae sedis" | rep1$Taxon_class == "Unicells and flagellates classes incertae sedis"] <- "Classes incertae sedis"
	rep1$Taxon_class[rep1$Taxon_class=="Euglenoidea"] <- "Euglenophyceae"
	rep1$Taxon_class[rep1$Taxon_class=="Oligotrichea"|rep1$Taxon_class == "Prostomatea"|rep1$Taxon_class == "Litostomatea"|rep1$Taxon_class == "Oligohymenophorea"] <- "Ciliophora"
	rep1$Taxon_class[rep1$Taxon_class=="Cyanophyceae"] <- "Cyanobacteria"
	rep1$Taxon_class[rep1$Taxon_class=="Prymnesiophyceae"] <- "Coccolithophyceae"
	rep1$Taxon_class[rep1$Taxon_class=="Trebouxiophyceae"|rep1$Taxon_class == "Chlorophyceae"|rep1$Taxon_class == "Zygnematophyceae"|rep1$Taxon_class == "Nephroselmidophyceae"|rep1$Taxon_class == "Mamiellophyceae"] <- "Chlorophyta"

	# Insert Norwegian names in Taxon_class column
	rep1$Taxon_class[rep1$Taxon_class == "Bacillariophyceae"] <- "Bacillariophyceae (kiselalger)"
	rep1$Taxon_class[rep1$Taxon_class == "Chlorophyta"] <- "Chlorophyta (grønnalger)"
	rep1$Taxon_class[rep1$Taxon_class == "Choanoflagellatea"] <- "Choanoflagellatea (krageflagellater)"
	rep1$Taxon_class[rep1$Taxon_class == "Chrysophyceae"] <- "Chrysophyceae (gullalger)"
	rep1$Taxon_class[rep1$Taxon_class == "Ciliophora"] <- "Ciliophora (ciliater)"
	rep1$Taxon_class[rep1$Taxon_class == "Coccolithophyceae"] <- "Coccolithophyceae (kalk- og svepeflagellater)"
	rep1$Taxon_class[rep1$Taxon_class == "Cryptophyceae"] <- "Cryptophyceae (svelgflagellater)"
	rep1$Taxon_class[rep1$Taxon_class == "Cyanobacteria"] <- "Cyanobacteria (blågrønnbakterier)"
	rep1$Taxon_class[rep1$Taxon_class == "Dictyochophyceae"] <- "Dictyochophyceae (kiselflagellater og pedineller)"
	rep1$Taxon_class[rep1$Taxon_class == "Dinophyceae"] <- "Dinophyceae (fureflagellater)"
	rep1$Taxon_class[rep1$Taxon_class == "Ebriophyceae"] <- "Ebriophyceae (skjelettflagellater)"
	rep1$Taxon_class[rep1$Taxon_class == "Euglenophyceae"] <- "Euglenophyceae (øyealger)"
	rep1$Taxon_class[rep1$Taxon_class == "Prasinophyceae"] <- "Prasinophyceae (olivengrønnalger)"
	rep1$Taxon_class[rep1$Taxon_class == "Protozoa classis incertae sedis"] <- "Protozoa"
	rep1$Taxon_class[rep1$Taxon_class == "Raphidophyceae"] <- "Raphidophyceae (nålflagellater)"
	rep1$Taxon_class[rep1$Taxon_class == "Xanthophyceae"] <- "Xanthophyceae (gulgrønnalger)"
	rep1$Taxon_class[rep1$Taxon_class == "Classes incertae sedis"] <- "Classes incertae sedis (ubestemte klasser)"

	# Rename taxa
	rep1$Taxon_name_extended[rep1$Taxon_name_extended == "Haptophyta"] <- "Haptofytter"
	rep1$Taxon_name_extended[rep1$Taxon_name_extended == "Pennales"] <- "Pennate kiselalger"
	rep1$Taxon_name_extended[rep1$Taxon_name_extended == "Centrales"] <- "Sentriske kiselalger"
	rep1$Taxon_name_extended[rep1$Taxon_name_extended == "Peridiniales"] <- "Tekate fureflagellater"
	rep1$Taxon_name_extended[rep1$Taxon_name_extended == "Gymnodiniales"] <- "Atekate fureflagellater"
	rep1$Taxon_name_extended[rep1$Taxon_name_extended == "Flagellates"] <- "Flagellater"
	rep1$Taxon_name_extended[rep1$Taxon_name_extended == "Unicell"] <- "Monader"
	rep1$Taxon_name_extended[rep1$Taxon_name_extended == "Unicells_and_flagellates" ] <- "Monader"

	#Delete unneccessary columns and reshaping the data
	rep_h <- rep1 %>% 
		filter(Presence > 0) %>% 
		select(Sample_date, Station_code, Taxon_class, Taxon_name_extended, Presence) %>% 
		mutate_at(vars(Presence), as.numeric)
	
	rep2 <- rep1 %>% 
	  filter(is.na(Presence)) %>%
	  select(Sample_date, Station_name, Station_code, Max_depth, Taxon_class, Taxon_name_extended, Abundance, Calculated_carbon, SizeRange) %>%
	  replace(is.na(.), "") %>%
	  unite(Station, c("Station_code", "Station_name"), sep = " ") %>%
	  unite(Taxon, c("Taxon_name_extended", "SizeRange"), sep = " ")

	#for abundance
	rep_a <- rep2 %>%
	  select(!c(7)) %>%
	  mutate(Abundance = round(Abundance, 0)) 

	#for carbon
	rep_c <- rep2 %>%
	  select(!c(6)) %>%
	  mutate(Calculated_carbon = round(Calculated_carbon, 3))

	# Change date format
	rep_a$Sample_date <- format(as.Date(rep_a$Sample_date, format = "%Y-%m-%d"), "%d/%m/%Y")
	rep_c$Sample_date <- format(as.Date(rep_c$Sample_date, format = "%Y-%m-%d"), "%d/%m/%Y")
	rep_h$Sample_date <- format(as.Date(rep_h$Sample_date, format = "%Y-%m-%d"), "%d/%m/%Y")

	# Reshaping the data
	rep_a1 <- melt(rep_a, id = 1:5, na.rm = FALSE)
	rep_a2 <- dcast(rep_a1, Taxon_class + Taxon ~ factor(Sample_date, levels = unique(Sample_date)), sum)
	rep_c1 <- melt(rep_c, id = 1:5, na.rm = FALSE)
	rep_c2 <- dcast(rep_c1, Taxon_class + Taxon ~ factor(Sample_date, levels = unique(Sample_date)), sum)
	rep_h1 <- melt(rep_h, id = 1:4, na.rm = FALSE)
	rep_h2 <- dcast(rep_h1, Taxon_class + Taxon_name_extended ~ factor(Sample_date, levels = unique(Sample_date)), sum)

	# Including sums for each Taxon class
	numericCols = sapply(rep_a2, is.numeric)
	numericCols2 = sapply(rep_c2, is.numeric)
	numericColsh = sapply(rep_h2, is.numeric)

	func = function(df,numCols) {
	  sums <- colSums(df[, numCols])
	  names(sums) <- colnames(df[c(3:ncol(df))])
	  sums[is.na(sums)] <- 0
	  result <- rep(NA, ncol(df))
	  names(result) <- colnames(df)
	  result[names(sums)] <- sums
	  result[2] <- "Sum:"
	  df[df == 0] <- "."
	  rbind(df, result, rep(NA,ncol(df)), rep(NA,ncol(df)))
	}

	rep_a3 <- split(rep_a2, rep_a2$Taxon_class) %>% map_dfr(func, numCols = numericCols)
	rep_c3 <- split(rep_c2, rep_c2$Taxon_class) %>% map_dfr(func, numCols = numericCols2)
	rep_h3 <- split(rep_h2, rep_h2$Taxon_class) %>% map_dfr(func, numCols = numericColsh)


	# Including total sum
	sums_all <- colSums(rep_a2[, numericCols])
	names(sums_all) <- colnames(rep_a2[c(3:ncol(rep_a2))])
	result_all <- rep(NA, ncol(rep_a2))
	names(result_all) <- colnames(rep_a2)
	result_all[names(sums_all)] <- sums_all
	result_all[2] <- "Sum totalt:"
	rep_a4 <- rbind(rep(NA, ncol(rep_a2)), rep(NA, ncol(rep_a2)), rep_a3, result_all)

	sums_all2 <- colSums(rep_c2[, numericCols2])
	names(sums_all2) <- colnames(rep_c2[c(3:ncol(rep_c2))])
	result_all2 <- rep(NA, ncol(rep_c2))
	names(result_all2) <- colnames(rep_c2)
	result_all2[names(sums_all2)] <- sums_all2
	result_all2[2] <- "Sum totalt:"
	rep_c4 <- rbind(rep(NA, ncol(rep_c3)), rep(NA, ncol(rep_c3)), rep_c3, result_all2)

	rep_h4 <- rbind(rep(NA, ncol(rep_h3)), rep(NA, ncol(rep_h3)), rep_h3)

	# Fix Taxon_class names
	rep_a4$Taxon_class[duplicated(rep_a4$Taxon_class)] <- ""
	rep_c4$Taxon_class[duplicated(rep_c4$Taxon_class)] <- ""
	rep_h4$Taxon_class[duplicated(rep_h4$Taxon_class)] <- ""

	# Move Taxon_class column 1 up
	rep_a4 <- mutate_at(rep_a4, c(1), list(lead), n = 1)
	rep_a4[is.na(rep_a4)] <- ""
	rep_c4 <- mutate_at(rep_c4, c(1), list(lead), n = 1)
	rep_c4[is.na(rep_c4)] <- ""
	rep_h4 <- mutate_at(rep_h4, c(1), list(lead), n = 1)
	rep_h4[is.na(rep_h4)] <- ""

	# Rename columns
	StationCode <- paste(rep[1,9])
	Station <- paste(rep_a1[1,2], rep_a1[1,3], "m")
	Station2 <- paste(rep1[1,9], rep1[1,8], sep = "_")
	Station3 <- paste(rep_a1[1,2], "0-30 m")
	colnames(rep_a4)[colnames(rep_a4) == "Taxon_class"] <- ""
	colnames(rep_a4)[colnames(rep_a4) == "Taxon"] <- paste("Antall celler/liter for", Station, sep = " ")
	colnames(rep_c4)[colnames(rep_c4) == "Taxon_class"] <- ""
	colnames(rep_c4)[colnames(rep_c4) == "Taxon"] <- paste("Karbon i µg/liter for", Station, sep = " ")
	colnames(rep_h4)[colnames(rep_h4) == "Taxon_class"] <- ""
	colnames(rep_h4)[colnames(rep_h4) == "Taxon_name_extended"] <- paste("Håvtrekk for", Station3, sep = " ")

	# Combining abundance and carbon tables
	rep_c5 <- select(rep_c4, -c(1,2)) 
	rep_a_c <- cbind(rep_a4, rep_c5)
	rep_a_c <- rbind(rep(NA, ncol(rep_a_c)), rep_a_c) 
	rep_a_c[is.na(rep_a_c)] <- ""
	colnames(rep_a_c)[2] <- Station



	## Formatting final excel-workbooks --------------------------------------------------

	# Defining some objects
	cmax <- ncol(rep_a4)
	cmax2 <- ncol(rep_a_c)
	cmaxh <- ncol(rep_h4)
	rmax <- nrow(rep_a4)
	rmaxh <- nrow(rep_h4)
	l <- which(startsWith(rep_a4[, 2], "Sum"))
	lh <- which(startsWith(rep_h4[, 2], "Sum"))
	m <- which(rep_a4[, 1] != "")
	mh <- which(rep_h4[, 1] != "")

	# Create new empty workbooks
	wb1 <- createWorkbook()
	addWorksheet(wb1, Station2)
	writeData(wb1, sheet = 1, x = rep_a4)  #for abundance data
	wb2 <- createWorkbook()
	addWorksheet(wb2, Station2)
	writeData(wb2, sheet = 1, x = rep_c4)  #for carbon data
	wb3 <- createWorkbook()
	addWorksheet(wb3, Station2)
	writeData(wb3, sheet = 1, x = rep_h4)  #for håvtrekk data
	wb4 <- createWorkbook()
	addWorksheet(wb4, Station2)
	writeData(wb4, sheet = 1, x = rep_a_c) #for abundance+carbon data
	writeData(wb4, sheet = 1, x= "", startCol = 1, startRow = 1)
	writeData(wb4, sheet = 1, x= "Antall celler/liter", startCol = 3, startRow = 2)
	writeData(wb4, sheet = 1, x= "Karbon µg/liter", startCol = cmax+1, startRow = 2)
	mergeCells(wb4, sheet = 1, cols = cmax+1:(cmax-2), rows = 2)
	mergeCells(wb4, sheet = 1, cols = 3:cmax, rows = 2)

	# Merge cells for classes
	for (i in m) {
	  mergeCells(wb1, sheet = 1, cols = 1:2, rows = i + 1)
	  mergeCells(wb2, sheet = 1, cols = 1:2, rows = i + 1)
	  mergeCells(wb4, sheet = 1, cols = 1:2, rows = i + 2)
	}

	for (i in mh) {
	  mergeCells(wb3, sheet = 1, cols = 1:2, rows = i + 1)
	}

	# Set column width
	setColWidths(wb1, sheet = 1, cols = 1:cmax, widths = "auto")
	setColWidths(wb1, sheet = 1, cols = 1, widths = c(2))
	setColWidths(wb1, sheet = 1, cols = 2, widths = c(43))
	setColWidths(wb2, sheet = 1, cols = 1:cmax, widths = "auto")
	setColWidths(wb2, sheet = 1, cols = 1, widths = c(2))
	setColWidths(wb2, sheet = 1, cols = 2, widths = c(43))
	setColWidths(wb3, sheet = 1, cols = 1:cmaxh, widths = "auto")
	setColWidths(wb3, sheet = 1, cols = 1, widths = c(2))
	setColWidths(wb3, sheet = 1, cols = 2, widths = c(43))
	setColWidths(wb4, sheet = 1, cols = 1, widths = c(2))
	setColWidths(wb4, sheet = 1, cols = 2, widths = c(43))
	setColWidths(wb4, sheet = 1, cols = 3:cmax2, widths = c(11))
	

	# Create general style
	alls <- createStyle(fontSize = 10, fontName = "Arial", halign = "left", valign = "bottom")
	addStyle(wb1, sheet = 1, alls, rows = 1:rmax+1, cols = 1:cmax, gridExpand = TRUE)
	addStyle(wb2, sheet = 1, alls, rows = 1:rmax+1, cols = 1:cmax, gridExpand = TRUE)
	addStyle(wb3, sheet = 1, alls, rows = 1:rmaxh+1, cols = 1:cmaxh, gridExpand = TRUE)
	addStyle(wb4, sheet = 1, alls, rows = 1:rmax+1, cols = 1:cmax2, gridExpand = TRUE)

	# Create style for header
	hs <- createStyle(fontSize = 10, fontName = "Arial", textDecoration = "BOLD", halign = "center", valign = "bottom")
	addStyle(wb1, sheet = 1, hs, rows = 1, cols = 1:cmax)
	addStyle(wb2, sheet = 1, hs, rows = 1, cols = 1:cmax)
	addStyle(wb3, sheet = 1, hs, rows = 1, cols = 1:cmaxh)
	addStyle(wb4, sheet = 1, hs, rows = 1:2, cols = 1:cmax2, gridExpand = TRUE)
	
	# Create style for sample columns
	ss <- createStyle(halign = "right")
	addStyle(wb1, sheet = 1, ss, cols = 3:ncol(rep_a4), rows = 2:rmax+1, gridExpand = TRUE, stack = TRUE)
	addStyle(wb2, sheet = 1, ss, cols = 3:ncol(rep_a4), rows = 2:rmax+1, gridExpand = TRUE, stack = TRUE)
	addStyle(wb3, sheet = 1, ss, cols = 3:ncol(rep_h4), rows = 2:rmax+1, gridExpand = TRUE, stack = TRUE)
	addStyle(wb4, sheet = 1, ss, cols = 3:ncol(rep_a_c), rows = 3:rmax+1, gridExpand = TRUE, stack = TRUE)
	
	# Create style for taxon names in italics
	c2s <- createStyle(textDecoration = "italic")
	addStyle(wb1, sheet = 1, c2s, cols = 2, rows = 2:rmax+1, gridExpand = TRUE, stack = TRUE)
	addStyle(wb2, sheet = 1, c2s, cols = 2, rows = 2:rmax+1, gridExpand = TRUE, stack = TRUE)
	addStyle(wb3, sheet = 1, c2s, cols = 2, rows = 2:rmaxh+1, gridExpand = TRUE, stack = TRUE)
	addStyle(wb4, sheet = 1, c2s, cols = 2, rows = 2:rmax+1, gridExpand = TRUE, stack = TRUE)
	
	# Create style for taxon classes in bold
	c3s <- createStyle(textDecoration = "bold")
	addStyle(wb1, sheet = 1, c3s, cols = 1, rows = 2:rmax+1, gridExpand = TRUE, stack = TRUE)
	addStyle(wb2, sheet = 1, c3s, cols = 1, rows = 2:rmax+1, gridExpand = TRUE, stack = TRUE)
	addStyle(wb3, sheet = 1, c3s, cols = 1, rows = 2:rmaxh+1, gridExpand = TRUE, stack = TRUE)
	addStyle(wb4, sheet = 1, c3s, cols = 1, rows = 2:rmax+1, gridExpand = TRUE, stack = TRUE)
	
	# Create style for sum rows
	sus <- createStyle(fontSize = 10, fontName = "Arial", halign = "right", valign = "bottom", border = "top")
	addStyle(wb1, sheet = 1, sus, cols = 2:cmax, rows = l+1, gridExpand = TRUE, stack = FALSE)
	addStyle(wb2, sheet = 1, sus, cols = 2:cmax, rows = l+1, gridExpand = TRUE, stack = FALSE)
	addStyle(wb3, sheet = 1, sus, cols = 2:cmax, rows = lh+1, gridExpand = TRUE, stack = FALSE)
	addStyle(wb4, sheet = 1, sus, cols = 2:cmax2, rows = l+2, gridExpand = TRUE, stack = FALSE)
	
	# Save workbooks
	rep_ds <- paste(rep[1, 1])
	ifelse(!dir.exists(rep_ds), dir.create(rep_ds), FALSE)
	newfolder <- paste(rep_ds, "/", StationCode, sep = "")
	ifelse(!dir.exists(newfolder), dir.create(newfolder), FALSE)
	filepath1 <- paste(newfolder, "/", Station2, "_abundans", ".xlsx", sep = "")
	filepath2 <- paste(newfolder, "/",Station2, "_karbon", ".xlsx", sep = "")
	filepath3 <- paste(newfolder, "/",Station2, "_håvtrekk", ".xlsx", sep = "")
	filepath4 <- paste(newfolder, "/",Station2, "_a+c", ".xlsx", sep = "")
	saveWorkbook(wb1, file = filepath1)
	saveWorkbook(wb2, file = filepath2)
	saveWorkbook(wb3, file = filepath3)
	saveWorkbook(wb4, file = filepath4)
	
    	
	print(Station2)

}



#_______________________run Script_plots.R for creating a plot of chla, abundance and carbon data_______________________

for (i in 1:len) {
  y <- my.list[i]
  
  ## Abundance & carbon
 
  # Load files and rename classes 
  df <- report1 %>%
    filter(Station_code == y) %>%
    filter((Taxon_name != "Amoeba" & Taxon_name != "Ciliophora" & Taxon_name != "Solenicola setigera" & Taxon_class != "Oligotrichea" & Taxon_class != "Protozoa classis incertae sedis" & Taxon_class != "Imbricatea" & Taxon_class != "Oligohymenophorea") %>%
    replace_na(TRUE))
  
  df$Taxon_class[is.na(df$Taxon_class)] <- "Andre flagellater\nog monader"
  df$Taxon_class[df$Taxon_class != "Dinophyceae" & df$Taxon_class != "Bacillariophyceae"] <- "Andre flagellater\nog monader"
  df$Taxon_class[df$Taxon_class == "Dinophyceae"] <- "Fureflagellater"
  df$Taxon_class[df$Taxon_class == "Bacillariophyceae"] <-"Kiselalger"
  
  # Change column format, delete unnecessary columns, summarize, add 0
  df1 <- df %>% 
    filter(is.na(Presence)) %>% 
    rename(Abundance=Abundance_ind_l, Calculated_carbon=Calculated_carbon_ugC_l) %>%
    mutate_at(vars(Abundance, Calculated_carbon), as.numeric) %>%
    select(Sample_date, Taxon_class, Abundance, Calculated_carbon) %>% 
    group_by(Sample_date, Taxon_class) %>% 
    summarise_all(sum) %>% 
    ungroup(.) %>% 
    complete(Sample_date, nesting(Taxon_class), fill = list(Abundance = 0, Calculated_carbon = 0)) %>% 
    mutate_at(vars(Abundance, Calculated_carbon), as.numeric)
  df1 %>% select(Sample_date, Taxon_class, Abundance) -> df2_a
  df1 %>% select(Sample_date, Taxon_class, Calculated_carbon) -> df2_c
  
  # Change date format and define some variables
  df2_a$Sample_date <- format(as.Date(df2_a$Sample_date, format = "%Y-%m-%d"), "%d/%m/%Y")
  df2_c$Sample_date <- format(as.Date(df2_c$Sample_date, format = "%Y-%m-%d"), "%d/%m/%Y")
  StationCode <- paste(df[1, 9])
  Station2 <- paste(df[1, 9], df[1, 8], sep = "_")
  title <- paste(StationCode, " ", df[1, 8], ": Planteplankton", sep = "")
  
  
  ## Chlorophyll a 
  
  fil <- report2 %>% 
    select_all(~gsub("\\s+", "_", .)) %>%
    filter(Station_code == y)
  
  # Remodel dataframe
  fil1 <- fil %>%
    rename(chla = "KlfA_µg/l") %>%
    filter(!is.na(chla)) %>%
    filter(Depth_2 == "2") %>%
    relocate(chla, .after = Depth_2) %>%
    select(c(1:9)) %>%
    select_all(~gsub("\\< ", "", .)) %>%
    select_all(~gsub("\\,", ".", .)) %>%
    mutate_at(vars(chla, Sample_date), as.character)
  
  klf <- fil1 %>%
    select(c(6, 9)) %>% 
    mutate_at(vars(chla), as.numeric) 
  
  klf$Sample_date <- format(as.Date(klf$Sample_date, format = "%Y-%m-%d"), "%d/%m/%Y")
  
  #Define some variables
  title2 <- paste(StationCode, " ", fil[1, 5], ": Klorofyll a", sep = "")
  dates <- unique(klf$Sample_date)
  min <- as.Date(dates[1], format = "%d/%m/%Y")
  nval <- length(dates)
  max <- as.Date(dates[nval], format = "%d/%m/%Y")
  
  
  # Creating plots -----------------------------------------------------------------------------------------------------
  
  #Plot for chlorofyll a
  plot_chla <- 
    ggplot(klf, aes(x = as.Date(Sample_date, format = "%d/%m/%Y"), y = chla)) + 
    geom_line(colour = "black", linewidth = 1) + 
    scale_x_date(
      labels = date_format("%d/%m/%Y"), 
      breaks = as.Date(klf$Sample_date, format = "%d/%m/%Y"), 
      limits = c(min, max), expand = c(0,0)
    ) +
    theme(
      axis.text.x = element_text(angle = 40, vjust = 0.5), 
      plot.margin = margin(0.4, 2, -0.1, 0.1, unit = "cm"), 
      axis.title.y=element_text(size=20), 
      axis.text=element_text(size=18), 
      plot.title = element_text(hjust = 0, size = 24), 
      panel.grid.minor.x = element_blank()
    ) +
    ggtitle(title2) +
    expand_limits(y = 0) +
    labs(x = "", y = "µg/liter")
  
  #Plot for abundance
  plot_a <- 
    ggplot(df2_a, aes(x = as.Date(Sample_date, format = "%d/%m/%Y"), y = Abundance, fill = Taxon_class)) + 
    geom_area(stat = "identity", position = "stack") +
    scale_x_date(
      labels = date_format("%d/%m/%Y"), 
      breaks = as.Date(df2_c$Sample_date, format = "%d/%m/%Y"), 
      limits = c(min, max), 
      expand = c(0,0)
    ) +
    theme(
      axis.text = element_text(size = 18), 
      axis.text.x = element_blank(), 
      plot.title = element_text(hjust = 0, size = 24), 
      legend.text = element_text(size = 18), 
      legend.key.size = unit(1.1, "cm"),
      legend.position = "inside",
      legend.position.inside = c(0.91, 0.85), 
      legend.title = element_blank(), 
      axis.title.y = element_text(size = 20), 
      axis.ticks.x = element_blank(), 
      panel.grid.minor.x = element_blank(), 
      plot.margin = unit(c(0.2, 2, -0.2, 0), "cm")
    ) +
    ggtitle(title) +
    scale_fill_manual(values = c("#c0c0c0", "#3366ff", "#f4bb64")) +
    guides(fill = guide_legend(reverse=TRUE)) +
    scale_y_continuous(labels = function(x) format(x, big.mark = " ", scientific = FALSE)) + 
    labs(x = "", y = "antall celler/liter") +
    coord_cartesian(xlim = c(min, max))
  
  #Plot for carbon
  plot_c <- 
    ggplot(df2_c, aes(x = as.Date(Sample_date, format = "%d/%m/%Y"), y = Calculated_carbon, fill = Taxon_class)) + 
    geom_area(stat = "identity", position = "stack") + 
    labs(x = "", y = "karbon µg/liter") +
    scale_x_date(
      labels = date_format("%d/%m/%Y"), 
      breaks = as.Date(df2_c$Sample_date, format = "%d/%m/%Y"), 
      limits = c(min, max), expand = c(0,0)
    ) +
    theme(
      axis.text.x = element_text(angle = 40, vjust = 0.5), 
      plot.margin = unit(c(-0.2, 2, 0, 0), "cm"), 
      axis.text = element_text(size = 18), 
      legend.position = "none", 
      axis.title.y = element_text(size = 20), 
      panel.grid.minor.x = element_blank()
    ) +
    scale_fill_manual(values = c("#c0c0c0", "#3366ff", "#f4bb64")) +
    scale_y_continuous(labels = function(x) format(x, big.mark = " ", scientific = FALSE)) +
    coord_cartesian(xlim = c(min, max))
  
  plots <- list(plot_chla, plot_a, plot_c)
  
  # Save in one figure
  grid.newpage()
  rep_ds <- paste(df[1, 1])
  ifelse(!dir.exists(rep_ds), dir.create(rep_ds), FALSE)
  newfolder <- paste(rep_ds, "/", StationCode, sep = "")
  ifelse(!dir.exists(newfolder), dir.create(newfolder), FALSE)
  file_station <- paste(newfolder, "/", Station2, "_plot_chla+a+c.png", sep="")
  ggsave(
    plot_grid(plotlist = plots, ncol = 1, align = "v"), 
    filename = file_station, 
    width = unit(15, "cm"), 
    height = unit(15, "cm")
  )
    
    print(paste("plot saved for", Station2))

}
 print("All done")