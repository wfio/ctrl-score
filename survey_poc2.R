################################################################################
# Author: Zach                                                                 #
# License: MIT, free to use, re-distrubte and modify w/o commercial gain.      #
# Date: ~ 20 Aug 2017                                                          #
# Purpose: # This script automates the processing of tidy data in Excel workbo #
#oks. It eliminates moving all of the spreadsheets into a directory by manual- #
#ly copying and pasting all of the content into a semi-static worksheet.       #
# The semi-static worksheet approach is subject to errors or ommissions if the #
#underlying original worksheets are modified at a later date. As a result, the #
#underlying workbooks can continue to be modified and the script just needs to #
# be run once after all modifications are completed. The script performs vari- #
#ous grouping actions and then summarizes the raw scores by producing simple   #
#averages and the corresponding ratings.                                       #
################################################################################

#load packages
require(tidyverse) #contains dplyr, tidyr, tibbles and other nifty things
require(readxl) #fast XSLX parser

# Establish the file names to read and read the contents of the Excel files into
# a list of data.frames
files.list <-list.files(pattern = '*.xlsx')  # Get a list of file names in the directory
df.files <-setNames(lapply(files.list, read_excel), files.list)  # Write the 
# contents of those files into a list of data.frames
df <- bind_rows(df.files, .id = "file_src") #bind the rows together to form a 
# tall data set of tidy data, with every variable with an observation.

# Tidy up yo-self! And don't leave no trash!
rm(df.files, files.list)

# Create Assessment Unit and Business Unit Summary Statistics ##################
# Stage 2 processing                                                           #
# Create summary statistics (simple average) data.frames at the assessment and #
# business unit levels across the control categories. This section of the      #
# script takes the data.frame and collapses the data at the levels of AU and   #
# CC. This allows for simple summary statistics via (dplyr::summarise) to      #
# operate on each vector of the data.frame. Thus, each CC of a given AU or BU  #
# is the simple mathetmatical average of the raw scores from the vector of     #
# corresponding scores for the given category (GOV, CORC, etc scores [0-4])    #
# where 0 = N/A, 1 = "Highly Effective", 2 = "Generally Effective",            #
# 3 = "Marginally Effective" and # 4 = "Ineffective". These scores and ratings #
# are established and defined in the methodology. Subsequent sections handle   #
# the weighted factorization of the scores and the sum of weighted scores.     #
################################################################################

# This section produces a table of simple averages (scores) for each business
# and each assessment unit across the respective control categories for each 
# level of the unit AU vs BU.
scores.list <- list(AU = df %>%
                  # scores.list$AU
                  group_by(AU, CC) %>% # groups df by AU and CC
                  summarise(avg = mean(Score, na.rm = TRUE)), #AU means
              
                      # scores.list$BU
                  BU = df %>% # 
                  group_by(BU, CC) %>% # groups df by BU and CC
                  summarise(avg = mean(Score, na.rm = TRUE)) # BU means
              ) #end-list

df.table <- bind_rows(scores.list, .id = "Unit_Level") %>% 
  replace(., is.na(.), "") %>%  # To avoid "NA" values when we "unite" below
  unite(Unit_Name, AU, BU, sep="") %>% 
  spread(CC, avg)
# df.table
rm(df) # clean up space
#
# End Stage 2 Processing #######################################################

# Create table of control category weights
weight.table <- tribble(
  ~AU, ~AUDIT, ~CORC, ~GOV, ~PPS, ~TMSC, ~TRAIN,
  "BSA_Admin",.10,.10,.30,.10,.30,.10,
  "BSG",.10,.10,.30,.10,.30,.10,
  "CBD",.10,.30,.30,.10,.10,.10,
  "CBG",.10,.30,.30,.10,.10,.10,
  "WMG",.10,.30,.30,.10,.10,.10
)

convert_scores <- function(score) {
  # Creates a reference data.frame of score ranges to string values
  #
  # Args: 
  #  score: The simple average from a data.frame of scores
  #
  # Returns:
  #  A string value that corresponds to a given score in the range.
  #  e.g., 2.45 would return "Generally Effective"
  rating <- case_when(
    score > 0 & score < 1.5 ~ "Highly Effective",
    score >= 1.5 & score < 2.5 ~ "Generally Effective",
    score >= 2.5 & score < 3.5 ~ "Marginally Effective",
    score >= 3.5 ~ "Ineffective"
  )
  return(rating)
} #end-function(convert_scores)

# Create tables of Control Ratings using the covert_scores function to translate
# the simple average to a Control rating. Creates AU and BU level tables. Paste
# these back into Excel and conditionally color to give a heat map. At least,
# until I figure out a way to conditionally color within R.

au.ratings.tbl <- df.table %>%
  filter(Unit_Level == "AU") %>%
  mutate_at(c("AUDIT", "CORC", "GOV", "PPS", "TMSC", "TRAIN"), convert_scores)
#  au.ratings.tbl[,1:8]

bu.ratings.tbl <- df.table %>%
  filter(Unit_Level == "BU") %>%
  mutate_at(c("AUDIT", "CORC", "GOV", "PPS", "TMSC", "TRAIN"), convert_scores)
#  bu.ratings.tbl[,1:8]

# Calculate weighted scores for the assessment units. Matrix multiplication occ-
# urs in an element-wise nature for the two given data.frames. A raw score
# multiplied by its associated weight in the weight.table df.
au.names <- c("BSA_Admin", "BSG", "CBD", "CBG", "WMG")
au.weighted.scores <- spread(scores.list$AU, CC, avg)[,-1]*weight.table[,-1]
au.weighted.scores <- cbind(AU = paste0(au.names), au.weighted.scores)

lob.weighted.scores <- au.weighted.scores %>%
  gather(Control_Category, weighted.score, -AU) %>%
  group_by(Control_Category) %>%
  summarise(
    # Order: BSA_Admin, BSG, CBD, CBG, WMG
            CBD = round(weighted.mean(weighted.score, c(1,1,1,0,0)),2),
            CBG = round(weighted.mean(weighted.score, c(1,1,0,1,0)),2),
            WMG = round(weighted.mean(weighted.score, c(1,1,0,0,1)),2), 
            BSA_Admin = round(weighted.mean(weighted.score, c(1,0,0,0,0)),2),
            BSG = round(weighted.mean(weighted.score, c(0,1,0,0,0)),2)) %>%
  ungroup()
lob.weighted.scores

# rearrange result & calculate overall sum for each line of business
lob.control.results <- lob.weighted.scores %>%
  gather(AU, score, -Control_Category) %>%
  spread(Control_Category, score) %>%
  select(AU, GOV, CORC, TMSC, AUDIT, PPS, TRAIN) %>%
  mutate(Overall_Score = AUDIT + CORC + GOV + PPS + TMSC + TRAIN) %>%
  mutate(Overall_Rating = convert_scores(Overall_Score)
) # end-chain
lob.control.results

# calculate weighted scores for BSA Admin and BSG
central.control.scores <- lob.control.results[1:2, ]
