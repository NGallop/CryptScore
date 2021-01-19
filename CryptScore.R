# Automating Crypt Reports V2.0
# Author: Niall Gallop
# Date: 29/06/2020


library('xlsx')
library('dplyr')

path =  # the folder the Qupath files will be read from
compiled_case = data.frame() # sets up empty data frame for cumulative data to go into
slides = dir(path, pattern='.txt') # assigning the Qupath files a value
save_path =  # the folder that the excel spreadsheet will be saved into


for (s in 1:length(slides)){ # loop through files within a directory

  rawdata = read.table(file.path(path, slides[s]), header=TRUE, sep='\t') # read qupath table


  annotations = data.frame(
    Class = c('Negative', 'Other', 'Partial', 'Patch'),
    Count = c(sum(rawdata$Class == 'Negative'),
              sum(rawdata$Class == 'Other'),
              sum(rawdata$Class == 'Partial'),
              sum(rawdata$Class == 'Patch'))
  )

  fact_char = as.character(rawdata$Name)
  char_num = as.numeric(fact_char) # coerce data into numerical value
  patch_crypts = na.omit(char_num) # remove n/a
  patch_total = sum(patch_crypts) # patch_total is the amount of crypts within patches


  Npartial = as.numeric(annotations$Count[3]) #number of partials extracted from annotations data frame
  Npatch = as.numeric(annotations$Count[4]) #number of patches extracted from annotations data frame
  Nloss = as.numeric(annotations$Count[1]) #number of negative crypts extracted from annotations data frame
  Nclone = Npatch + Npartial + (Nloss - patch_total) # formula to generate number of clones: number of patches and partials added to the total number of crypts scored minus the number of crypts within patches

  Patch_analysis = data.frame(patch_crypts) #breakdown of patch sizes

  Crypt_analysis = data.frame(
    Info = c('No. of crypts', 'No. of clones', 'No. of patches', 'No. of partials'),
    number = c(Nloss, Nclone, Npatch, Npartial)
  ) # Processed data frame with desired information

  compiled_case = rbind.data.frame(compiled_case, c(slides[s], Crypt_analysis$number), stringsAsFactors = FALSE) # build cumulative data frame

  savefilename = strtrim(slides[s], nchar(slides[s])-4) # create name for the save file

  write.xlsx(
    Crypt_analysis,
    paste0(save_path, savefilename, '.xlsx'),
    sheetName = "Crypt Analysis",
    col.names=TRUE,
    row.names = FALSE,
    append = FALSE)
  write.xlsx(
    Patch_analysis,
    paste0(save_path, savefilename, '.xlsx'),
    sheetName = 'Patch Breakdown',
    col.names = TRUE,
    row.names = FALSE,
    append = TRUE)
  write.xlsx(
    rawdata,
    paste0(save_path, savefilename, '.xlsx'),
    sheetName = 'Raw Data',
    col.names = TRUE,
    row.names = FALSE,
    append = TRUE) # write all data required into an excel spreadsheet
} # repeat loop for all files until the last file has been processed, then end and move onto:

colnames(compiled_case) = c('File ID', 'No. of crypts', 'No. of clones', 'No. of patches', 'No. of partials') # give cumulative data frame column names
totals = c('Totals',
           sum(as.numeric(compiled_case[, 2])),
           sum(as.numeric(compiled_case[, 3])),
           sum(as.numeric(compiled_case[, 4])),
           sum(as.numeric(compiled_case[, 5]))
) # create a vector of totals from the cumulative data frame

compiled_case = rbind.data.frame(compiled_case, totals) # stick the totals vector to the bottom of the cumulative data frame

cum_save_name = strtrim(savefilename, 4)

write.xlsx(
  compiled_case,
  paste0(save_path, cum_save_name, '_totals.xlsx'),
  sheetName = 'Case Overview',
  col.names = TRUE,
  row.names = FALSE,
  append = FALSE
) # save cumulative data frame to the same folder as the
