# This script automates the processing of the control inventory survey
# workbooks. It eliminates moving all of the spreadsheets into a directory by
# manually copying and pasting all of the content into a semi-static worksheet.
# The semi-static worksheet approach is subject to errors or ommissions if the
# underlying original worksheets are modified at a later date. Now, the
# underlying workbooks can continue to be modified and the script just needs to
# be run once after all modifications are completed. The script performs various
# grouping actions and then summarizes the raw scores by producing simple
# averages and the corresponding ratings.

#load packages
require(dplyr)
require(readxl)
require(reshape2)
require(tidyr)
require(tibble)

# Establish the file names to read and read the contents of the Excel files into
# a list of data.frames
files.list <-list.files(pattern = '*.xlsx')  # Get a list of file names in the directory
df.files <-setNames(lapply(files.list, read_excel), files.list)  # Read the contents of those files into a list of data.frames
df <- bind_rows(df.files, .id = "file_src")

# Tidy up yo-self! And don't leave no trash!
#rm(df.files, files.list)

# Create summary statistics data.frames at the assessment and business unit
# levels across the control categories BSA Admin assessment unit scores
bsa.au.stat.wide <- df %>%
  group_by(AU, CC) %>%
  filter(AU == "BSA Admin") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(AU ~ CC, value.var = "avg") %>% 
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

#bind all of the rows for all assessment units together
au.scores <- bind_rows(
  bsa.au.stat.wide,bsg.au.stat.wide,cbd.au.stat.wide,
  cbg.au.stat.wide,wmg.au.stat.wide)

# Create table of control category weights
weight.table <- tribble(
  ~AU, ~AUDIT, ~CORC, ~GOV, ~PPS, ~TMSC, ~TRAIN,
  "BSA Admin",.10,.10,.30,.10,.30,.10,
  "BSG",.10,.10,.30,.10,.30,.10,
  "CBD",.10,.30,.30,.10,.10,.10,
  "CBG",.10,.30,.30,.10,.10,.10,
  "WMG",.10,.30,.30,.10,.10,.10
)

# arrange columns in au.scores to match order of columns in weight.table and
# re-name au.scores to lob scores
lob.scores <- au.scores %>% arrange(AU, GOV, CORC, TMSC, AUDIT, PPS, TRAIN)

# calculate weighted scores for the lines of business
lob.weighted.scores <- lob.scores[,-1]*weight.table[,-1]
lob.weighted.scores$AU <- lob.scores$AU

# calculate scores for each line of business
lob.weighted.scores <- lob.weighted.scores %>%
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