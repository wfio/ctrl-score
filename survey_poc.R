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

# Establish the file names to read and read the contents of the Excel files into
# a list of data.frames
files.list <-
  list.files(pattern = '*.xlsx')  # Get a list of file names in the directory
df.files <-
  setNames(lapply(files.list, read_excel), files.list)  # Read the contents of those files into a list of data.frames
df <- bind_rows(df.files, .id = "file_src")

# Tidy up yo-self! And don't leave no trash!
#rm(df.files, files.list)

# Create summary statistics data.frames at the assessment and business unit
# levels across the control categories BSA Admin assessment unit scores
df %>%
  group_by(AU, CC) %>%
  filter(AU == "BSA Admin") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(AU ~ CC, value.var = "avg") %>%
  print() -> bsa.au.stat.wide  # end chain
# Creates the summary statistics of the BSA Admin business units by Control Category
df %>%
  group_by(BU, CC) %>%
  filter(AU == "BSA Admin") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(BU ~ CC, value.var = "avg") %>%
  print() -> bsa.bu.stat.wide  # end chain

# BSG assessment unit scores
df %>%
  group_by(AU, CC) %>%
  filter(AU == "BSG") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(AU ~ CC, value.var = "avg") %>%
  print() -> bsg.au.stat.wide  # end chain
# Creates the summary statistics of the BSG business units by Control Category
df %>%
  group_by(BU, CC) %>%
  filter(AU == "BSG") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(BU ~ CC, value.var = "avg") %>%
  print() -> bsg.bu.stat.wide  # end chain

# CBD Assessment Unit scores
df %>%
  group_by(AU, CC) %>%
  filter(AU == "CBD") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(AU ~ CC, value.var = "avg") %>%
  print() -> cbd.au.stat.wide  # end chain
# CBD business units by Control Category
df %>%
  group_by(BU, CC) %>%
  filter(AU == "CBD") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(BU ~ CC, value.var = "avg") %>%
  print() -> cbd.bu.stat.wide  # end chain

# CBG assessment unit scores
df %>%
  group_by(AU, CC) %>%
  filter(AU == "CBG") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(AU ~ CC, value.var = "avg") %>%
  print() -> cbg.au.stat.wide  # end chain
# CBG business units by Control Category
df %>%
  group_by(BU, CC) %>%
  filter(AU == "CBG") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(BU ~ CC, value.var = "avg") %>%
  print() -> cbg.bu.stat.wide  # end chain

# WMG assessment unit scores
df %>%
  group_by(AU, CC) %>%
  filter(AU == "WMG") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(AU ~ CC, value.var = "avg") %>%
  print() -> wmg.au.stat.wide  # end chain
# Creates the summary statistics of the WMG business units by Control Category
df %>%
  group_by(BU, CC) %>%
  filter(AU == "WMG") %>%
  summarise(avg = mean(Score, na.rm = TRUE)) %>%
  dcast(BU ~ CC, value.var = "avg") %>%
  print() -> wmg.bu.stat.wide  # end chain


cbd.overall <- data.frame(
  cbd.gov = (au.scores[3, 4] * .30) * .33 + (au.scores[1, 4] * .30) * .33 +
    (au.scores[2, 4] * .30) * .33,
  cbd.corc = (au.scores[3, 3] * .30) * .33 + (au.scores[1, 3] * .1) * .33 +
    (au.scores[2, 3] * .1) * .33,
  cbd.tmsc = (au.scores[3, 6] * .30) * .33 + (au.scores[1, 6] * .30) * .33 +
    (au.scores[2, 6] * .30) * .33,
  cbd.audit = (au.scores[3, 2] * .30) * .33 + (au.scores[1, 2] * .30) * .33 +
    (au.scores[2, 2] * .30) * .33,
  cbd.pps = (au.scores[3, 5] * .30) * .33 + (au.scores[1, 5] * .30) * .33 +
    (au.scores[2, 5] * .30) * .33,
  cbd.train = (au.scores[3, 7] * .30) * .33 + (au.scores[1, 7] * .30) * .33 +
    (au.scores[2, 7] * .30) * .33
)