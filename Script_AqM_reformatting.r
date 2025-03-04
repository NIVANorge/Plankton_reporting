######################################################################
##SCRIPT FOR CONVERTING PLANKTON DATA TO FIT AQUAMONITOR IMPORT FORM##
##																	                                ##
##Sonja Kistenich, last updated 10.06.2024						            	##
##For Plankton Toolbox version 1.4.1							                	##
######################################################################

#Set working directory if necessary
setwd("C:/Users/SKI/R/R_Projects/Excel_modification")


#load libraries
library(readxl)
library(tidyverse)
library(writexl)
library(reshape2)

#ensure correct encoding of special characters
options(encoding = "UTF-8")			#for older R versions
Sys.setlocale(locale='no_NB.utf8') 	#for newer R versions

#import report file from PTB
file1 <- list.files(pattern = "rep_")
rep <- read_xlsx(file1)

df <- as.data.frame(rep)
df1 <- df %>% 
  arrange(Station_code, Sample_date, Min_depth) %>% 
  mutate_at(vars(Abundance_ind_l, Biovolume_mm3_l, Calculated_carbon_ugC_l, Presence), as.numeric) %>%
  select(-Size_range, -Aphia_id)

sizeclasses <- read_xlsx("SizeClasses_2024_3.xlsx", col_types = "text")
rep1 <- left_join(df1, sizeclasses, by = c("Taxon_name" = "Species", "Size_class" = "SizeClass"))

#correct taxonomic names
rep$Taxon_name_extended[rep$Taxon_name_extended == "Protoperidinium marie-lebouriae "] <- "Protoperidinium marielebouriae"
rep$Taxon_name_extended[rep$Taxon_name_extended == "Protoperidinium cf. marie-lebouriae "] <- "Protoperidinium cf. marielebouriae"
rep$Taxon_name_extended[rep$Taxon_name_extended == "cf. Protoperidinium marie-lebouriae "] <- "cf. Protoperidinium marielebouriae"

# Change date format
rep1$Sample_date <- format(as.Date(rep1$Sample_date, format = "%Y-%m-%d"), "%d.%m.%Y")

#correct Presence column
rep1$Presence <- replace(rep1$Presence, rep1$Presence > 1, 1)

#correct spp column
rep1$Species_flag <- replace(rep1$Species_flag, rep1$Species_flag == "sp.", "spp.")

#correct Rank column for Chaetoceros (Phaeoceros)
list1 <- grep("Phaeoceros", rep1$Taxon_name)
for (i in list1) {
  rep1[i, 19] <- "Genus"
}

#correct Rank column for Ptychocylis
list2 <- grep("Ptychocylis", rep1$Taxon_name)
for (i in list2) {
  rep1[i, 19] <- "Genus"
}

##correct column Species_flag
grep("GRP", rep1$Species_flag)
grep("CPX", rep1$Species_flag)
list3 <- which(rep1$Rank == "Genus" & is.na(rep1$Species_flag))
for (i in list3) {
  rep1[i, 20] <- "spp."
}

list4 <- which(rep1$Rank != "Genus" & rep1$Species_flag == "spp.")
for (i in list4) {
  rep1[i, 20] <- NA
}


#combine taxon names
rep1a <- rep1 %>%
	group_by(across(c(1:21, 26, 34:39))) %>%
	summarise(across(c(Abundance_ind_l, Biovolume_mm3_l, Calculated_carbon_ugC_l), sum)) %>%
	ungroup(.) %>%
	unite(Taxonomy, Taxon_name, SizeRange, sep= " ", na.rm = TRUE)
	

#change columns
rep2 <- rep1a %>% 
  #mutate(Sampling_method = if_else(.$Sampler_type_code == "NET", 'Håv20', 'Vannhenter (NS-EN 15972:2011)')) %>%   ## NB: change for IO/YO and Økokyst!!
  mutate(Sampling_method = "FBOX-AUTSAMPLE") %>%   ## for Ferrybox
  mutate(Cf = replace(.$Cf, Cf == "cf. (genus)", "genus")) %>%
  mutate(Cf = replace(.$Cf, Cf == "cf. (species)", "species")) %>%
  mutate(Species_flag = replace(.$Species_flag, Species_flag == "spp.", 2)) %>%
  mutate(Method = NA) %>%
  mutate(Taxonomy_domain = "Phytoplankton - Plankton Toolbox") %>%
  mutate(Stage_code = NA) %>%
  rename(Method_ref = Method_reference_code) %>%
  rename(Depth1 = Min_depth) %>%
  rename(Depth2 = Max_depth) %>%
  rename(SP = Species_flag) %>%
  rename(Laboratory = Analytical_laboratory)


#rearrange and delete columns
rep3 <- rep2[, c(2, 9, 8, 4, 31, 12, 13, 33, 18, 20, 21, 34, 32, 27, 23, 22, 28:30)] 

#reshaping data
rep4 <- melt(rep3)

rep5 <- rep4 %>%
	drop_na(value) %>%
	rename(unit = variable) %>%
	mutate(Method = case_when(unit == "Abundance_ind_l" ~ "Celletall",
							unit == "Biovolume_mm3_l" ~ "Biovolum",
							unit == "Calculated_carbon_ugC_l" ~ "Cellekarbon",
							unit == "Presence" ~ "Presence/absence")) %>%
	mutate(unit = case_when(unit == "Abundance_ind_l" ~ "Celler/l",
							unit == "Biovolume_mm3_l" ~ "mm3/l",
							unit == "Calculated_carbon_ugC_l" ~ "µg C/l",
							unit == "Presence" ~ "")) %>%
  mutate(Method_ref = case_when(Method == "Celletall" ~ "NS-EN 15204/NS-EN 15972:2011",
                          Method == "Biovolum" ~ "NS-EN 15204/NS-EN 15972:2011/Olenina et al. 2006",
                          Method == "Cellekarbon" ~ "NS-EN 15204/NS-EN 15972:2011/Olenina et al. 2006",
                          Method == "Presence/absence" ~ "")) 

names(rep5) <- toupper(names(rep5))


	

#save to new file
cur_date <- Sys.Date()
cur_date <- format(as.Date(cur_date, format = "%Y-%m-%d"), "%Y%m%d")
project <- paste(rep2[1,1], "_plankton_data_2024_", cur_date, ".xlsx", sep = "")
write_xlsx(rep5, project)
