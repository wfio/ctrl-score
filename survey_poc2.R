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

# This section produces a table of raw average scores for each business unit
# and each assessment unit across the respective control categories.
scores.list <- list(AU = df %>%
                  group_by(AU, CC) %>% # groups df by AU and CC
                  summarise(avg = mean(Score, na.rm = TRUE)) %>%
                    spread(AU, CC, avg),
                
                BU = df %>% # 
                  group_by(BU, CC) %>% # groups df by AU and CC
                  summarise(avg = mean(Score, na.rm = TRUE)) %>%
                    spread(BU, CC, avg))

df.table <- bind_rows(scores.list, .id = "Unit Level") %>% 
  replace(., is.na(.), "") %>%  # To avoid "NA" values when we "unite" below
  unite(Unit, AU, BU, sep="") %>% 
  spread(CC, avg)
df.table
#
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

convert_scores <- function(score) {
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

# Create tables of Control Ratings using the covert_scores function to translate
# the simple average to a Control rating. Creates AU and BU level tables.

au.ratings.tbl <- df.table %>%
  filter(`Unit Level` == "AU") %>%
  mutate_at(c("AUDIT", "CORC", "GOV", "PPS", "TMSC", "TRAIN"), convert_scores)
  au.ratings.tbl[,1:8]

bu.ratings.tbl <- df.table %>%
  filter(`Unit Level` == "BU") %>%
  mutate_at(c("AUDIT", "CORC", "GOV", "PPS", "TMSC", "TRAIN"), convert_scores)
bu.ratings.tbl[,1:8]

# arrange columns in au.scores to match order of columns in weight.table and
# re-name au.scores to lob scores
#lob.avg.scores <- au.avg.scores %>% arrange(AU, GOV, CORC, TMSC, AUDIT, PPS, TRAIN)

# calculate weighted scores for the lines of business
au.names <- c("BSA Admin", "BSG", "CBD", "CBG", "WMG")
au.weighted.scores <- spread(scores.list$AU, CC, avg)[,-1]*weight.table[,-1]
rownames(au.weighted.scores) <- au.names

# calculate weighted average scores for each line of business (CBD, CBG, WMG)
au.weighted.scores <- au.weighted.scores %>%
  gather(Control_Category, weighted.score) %>%
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