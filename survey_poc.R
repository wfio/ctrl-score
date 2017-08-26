# Author: Zach                                                                 #
# Description of Script Purpose:                                               #
# This script automates the processing of tidy data in Excel workbooks. It eli-#
# minates moving all of the spreadsheets into a directory by manually copying  #
# and pasting all of the content into a semi-static worksheet. The semi-static #
# worksheet approach is subject to errors or ommissions if the underlying orig-#
# inal worksheets are modified at a later date. As a result, the underlying wo-#
# rkbooks can continue to be modified and the script just needs to be run once #
# after all modifications are completed. The script performs various grouping  #
# actions and then summarizes the raw scores by producing simple averages and  #
# the corresponding ratings.                                                   #
################################################################################

#load packages
require(tidyverse) #contains dplyr, tidyr, tibbles and other nifty things
require(readxl) #fast XSLX parser
require(reshape2) #reshapes long dfs into wide dfs via dcast

# Establish the file names to read and read the contents of the Excel files into
# a list of data.frames
files.list <-list.files(pattern = '*.xlsx')  # Get a list of file names in the directory
df.files <-setNames(lapply(files.list, read_excel), files.list)  # Read the contents of those files into a list of data.frames
df <- bind_rows(df.files, .id = "file_src")

# Tidy up yo-self! And don't leave no trash!
#rm(df.files, files.list)

# Begin Assessment Unit and Business Unit Summary Statistics ###################
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
# are established and defined in the methodology.                              #
################################################################################

# Assessment unit scores (decide spread spread and dcast)
# The script groups the raw data by group AU (rownames) and CC (colnames). Once
# grouped, it filters out rownames == 'AU_name' then derives the simple avg
# for each CC for each AU. The data is then pivoted from long to wide form and 
# stored as a tibble or data.frame 'au_name.au.stat.wide'.

bsa.au.stat.wide <- df %>% # 
  group_by(AU, CC) %>% # groups df by AU and CC
  filter(AU == "BSA Admin") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(AU ~ CC, value.var = "avg") %>% 
  print()  # end chain

bsa.au.stat.spread <- df %>% # 
  group_by(AU, CC) %>% # groups df by AU and CC
  filter(AU == "BSA Admin") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  spread(key = CC, value = avg) %>% 
  print()  # end chain

# Creates the summary statistics of the BSA Admin business units by Control Category
bsa.bu.stat.wide <- df %>%
  group_by(BU, CC) %>%
  filter(AU == "BSA Admin") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(BU ~ CC, value.var = "avg") %>%
  print()  # end chain

# BSG assessment unit scores
bsg.au.stat.wide <- df %>%
  group_by(AU, CC) %>%
  filter(AU == "BSG") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(AU ~ CC, value.var = "avg") %>%
  print()  # end chain
# Creates the summary statistics of the BSG business units by Control Category
bsg.bu.stat.wide <- df %>%
  group_by(BU, CC) %>%
  filter(AU == "BSG") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(BU ~ CC, value.var = "avg") %>%
  print()  # end chain

# CBD Assessment Unit scores
cbd.au.stat.wide <- df %>%
  group_by(AU, CC) %>%
  filter(AU == "CBD") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(AU ~ CC, value.var = "avg") %>%
  print()  # end chain
# CBD business units by Control Category
cbd.bu.stat.wide <- df %>%
  group_by(BU, CC) %>%
  filter(AU == "CBD") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(BU ~ CC, value.var = "avg") %>%
  print() # end chain

# CBG assessment unit scores
cbg.au.stat.wide <- df %>%
  group_by(AU, CC) %>%
  filter(AU == "CBG") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(AU ~ CC, value.var = "avg") %>%
  print()  # end chain
# CBG business units by Control Category
cbg.bu.stat.wide <- df %>%
  group_by(BU, CC) %>%
  filter(AU == "CBG") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(BU ~ CC, value.var = "avg") %>%
  print() # end chain

# WMG assessment unit scores
wmg.au.stat.wide <- df %>%
  group_by(AU, CC) %>%
  filter(AU == "WMG") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(AU ~ CC, value.var = "avg") %>%
  print() # end chain
# Creates the summary statistics of the WMG business units by Control Category
wmg.bu.stat.wide <- df %>%
  group_by(BU, CC) %>%
  filter(AU == "WMG") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(BU ~ CC, value.var = "avg") %>%
  print() # end chain, 

# Bind all of the rows together as a wide and tidy data frame of averages 
# across the given control category of a given assessment or business unit.
# assessment unit level
au.avg.scores <- bind_rows(
  bsa.au.stat.wide,bsg.au.stat.wide,cbd.au.stat.wide,
  cbg.au.stat.wide, wmg.au.stat.wide)
# business unit level
bu.avg.scores <- bind_rows(bsa.bu.stat.wide, bsg.bu.stat.wide, cbd.bu.stat.wide,
                           cbg.bu.stat.wide, wmg.bu.stat.wide)

# End Stage 2 Processing #######################################################

# Create table of control category weights
weight.table <- tribble(
  ~AU, ~AUDIT, ~CORC, ~GOV, ~PPS, ~TMSC, ~TRAIN,
  "BSA Admin",.10,.10,.30,.10,.30,.10,
  "BSG",.10,.10,.30,.10,.30,.10,
  "CBD",.10,.30,.30,.10,.10,.10,
  "CBG",.10,.30,.30,.10,.10,.10,
  "WMG",.10,.30,.30,.10,.10,.10
)

derive_ratings <- function(score) {
  # Creates a reference data.frame of score ranges to string values
  #
  # Args: 
  #  score: The simple average from a data.frame of scores
  #
  # Returns:
  #  A string value that corresponds to a given score in the range.
  #  e.g., 2.45 would return "Generally Effective"
  lbl <- case_when(
    score > 0 & score < 1.5 ~ "Highly Effective",
    score >= 1.5 & score < 2.5 ~ "Generally Effective",
    score >= 2.5 & score < 3.5 ~ "Marginally Effective",
    score >= 3.5 ~ "Ineffective"
  )
  return(lbl)
} #end-rating

au.ratings.cc <- au.avg.scores %>% 
  mutate_at(c("AUDIT", "CORC", "GOV", "PPS", "TMSC", "TRAIN"), score_label)
au.ratings.cc[,1:4]

bu.ratings.cc <- bu.avg.scores %>%
  mutate_at(c("AUDIT", "CORC", "GOV", "PPS", "TMSC", "TRAIN"), score_label)
bu.ratings.cc[,1:4]

# arrange columns in au.scores to match order of columns in weight.table and
# re-name au.scores to lob scores
lob.avg.scores <- au.avg.scores %>% arrange(AU, GOV, CORC, TMSC, AUDIT, PPS, TRAIN)

# calculate weighted scores for the lines of business
lob.weighted.avg.scores <- lob.avg.scores[,-1]*weight.table[,-1]
lob.weighted.avg.scores$AU <- lob.avg.scores$AU

# calculate weighted average scores for each line of business (CBD, CBG, WMG)
lob.weighted.avg.scores <- lob.weighted.avg.scores %>%
  gather(Control_Category, weighted.score, -AU) %>%
  group_by(Control_Category) %>%
  summarise(CBD = round(weighted.mean(weighted.score, c(1,1,1,0,0)),2),
            CBG = round(weighted.mean(weighted.score, c(1,1,0,1,0)),2),
            WMG = round(weighted.mean(weighted.score, c(1,1,0,0,1)),2)) %>%
  ungroup()

# rearrange result & calculate overall sum for each line of business
lob.control.scores <- lob.weighted.scores %>%
  gather(LOB, score, -Control_Category) %>%
  spread(Control_Category, score) %>%
  select(LOB, GOV, CORC, TMSC, AUDIT, PPS, TRAIN) %>%
  mutate(Control_Score = GOV + CORC + TMSC + AUDIT + PPS + TRAIN) %>%
  mutate(Control_Rating = case_when(
    Control_Score > 3.499 ~ "Ineffective",
    Control_Score > 2.499 & Control_Score <= 3.499 ~ "Marginally Effective",
    Control_Score >= 1.5 & Control_Score <= 2.499 ~ "Generally Effective",
    Control_Score < 1.5 ~ "Highly Effective"
  )
)# end-chain

################
# populate df of au.scores by the name bsa.bsg.scores
bsa.bsg.scores <- au.scores %>% arrange(AU, GOV, CORC, TMSC, AUDIT, PPS, TRAIN)

# calculate weighted scores for BSA Admin and BSG
bsa.bsg.weighted.scores <- au.scores[,-1]*weight.table[,-1]
bsa.bsg.weighted.scores$AU <- au.scores$AU

# calculate scores for BSA Admin and BSG
bsa.bsg.weighted.scores <- bsa.bsg.weighted.scores %>%
  gather(Control_Category, weighted.score, -AU) %>%
  group_by(Control_Category) %>%
  summarise(BSA_Admin = weighted.mean(weighted.score, c(1,0,0,0,0)),
            BSG = weighted.mean(weighted.score, c(0,1,0,0,0))) %>%
  ungroup()

# rearrange result & calculate overall sum for BSA Admin and BSG
bsa.bsg.control.scores <- bsa.bsg.weighted.scores %>%
  gather(group, score, -Control_Category) %>%
  spread(Control_Category, score) %>%
  select(group, GOV, CORC, TMSC, AUDIT, PPS, TRAIN) %>%
  mutate(Control_Score = GOV + CORC + TMSC + AUDIT + PPS + TRAIN) %>%
  mutate(Rating = case_when(
    Control_Score > 3.499 ~ "Ineffective",
    Control_Score > 2.499 & Control_Score <= 3.499 ~ "Marginally Effective",
    Control_Score >= 1.5 & Control_Score <= 2.499 ~ "Generally Effective",
    Control_Score < 1.5 ~ "Highly Effective"
  )
         )