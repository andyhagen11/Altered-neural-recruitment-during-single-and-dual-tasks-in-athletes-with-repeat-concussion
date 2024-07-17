## RMANOVA or RMANCOVA script (one-way or two-way) with pairwise comparisons that are corrected for multiple comparisons 

library(stats)
library(psych)
library(rstudioapi)
rm(list=ls())
 
#Print warnings as they come up for troubleshooting purposes
options(warn=1)

# Import Data in long format
data <- read_excel("Library/CloudStorage/OneDrive-Colostate/Colorado State/SNL/SRC Dual Task fNIRS/Concussion (n=20) and Controls (n=17) DTS LE Results.xlsx", sheet = "Long Form PFC")
#data<- data[data$Group == "SRC", ] # For if you want to filter and test groups/conditions separately

# Specify withinFactor and betweenFactor as variables
variable = readline("Which variable would you like to run a RMANOVA on?: ") 
withinFactor <- "Task"
betweenFactor <- "Group"

# Generate Summary Statistics
summaryStats = describeBy(data[[variable]], group = list(data[[withinFactor]], data[[betweenFactor]]), mat = TRUE)
#summaryStats = describeBy(data[[variable]], group = list(data[[withinFactor]]), mat = TRUE)

# Create linear model of an outcome of interest
# Two-way ANOVA
formulaString <- paste(variable, "~", withinFactor, "*", betweenFactor, "* Sex + Age") # using * to include interaction
# One-way ANOVA
#formulaString <- paste(variable, "~", withinFactor)
anovaModel <- aov(as.formula(formulaString), data = data)

# Create an ANOVA table of the model
anovaResults = anova(anovaModel)

# Calculate interaction pairwise comparisons
pairwiseResults <- list()
pairwiseResults[["Interaction"]] <- pairwise.t.test(data[[variable]], interaction(data[[withinFactor]], data[[betweenFactor]]), p.adj = "none", paired = FALSE)
#pairwiseResults <- pairwise.t.test(data[[variable]], (data[[withinFactor]]), p.adj = "none", paired = TRUE) # For one-way

# Calculate within group pairwise comparisons (if group main effect is significant) # exclude if doing one-way
splitData <- split(data, data[[betweenFactor]])
# Loop through each subset of the data
for (group in names(splitData)) {
  groupData <- splitData[[group]]
  # Perform pairwise t-tests for the within-subject factor within this group
  pairwiseResults[[group]]<- pairwise.t.test(groupData[[variable]], groupData[[withinFactor]], 
                                      p.adj = "none", paired = TRUE)
}

# Calculate adjusted p-values for all comparisons (excluding N/A in our correction)
all_pvalues <- unlist(lapply(pairwiseResults, function(x) x$p.value))
adjustedPValues <- p.adjust(all_pvalues, method = "fdr")

# Add unadjusted and adjusted p-values back to the results
index <- 1
for (name in names(pairwiseResults)) {
  num_pvalues <- length(pairwiseResults[[name]]$p.value)
  pairwiseResults[[name]]$adjusted_p.value <- matrix(adjustedPValues[index:(index + num_pvalues - 1)], 
                                                     nrow = nrow(pairwiseResults[[name]]$p.value),
                                                     ncol = ncol(pairwiseResults[[name]]$p.value))
  pairwiseResults[[name]]$unadjusted_p.value <- pairwiseResults[[name]]$p.value
  index <- index + num_pvalues
}

# Export CSVs if desired
exportCsv <- readline("Do you want to export the CSV files? (Y/N): ")
if (exportCsv == "Y") {
  # Choose export directory
  exportDir <- selectDirectory()
  # Write CSVs 
  write.csv(summaryStats, file.path(exportDir, "SummaryStats.csv"), row.names = TRUE)
  write.csv(anovaResults, file.path(exportDir, "ANOVA_Table.csv"), row.names = TRUE)
  
  for (group in names(pairwiseResults)) {
      rownames(pairwiseResults[[group]]$adjusted_p.value) <- rownames(pairwiseResults[[group]]$unadjusted_p.value)
      colnames(pairwiseResults[[group]]$adjusted_p.value) <- colnames(pairwiseResults[[group]]$unadjusted_p.value)
      write.csv(as.data.frame(pairwiseResults[["Interaction"]]$adjusted_p.value), 
              file.path(exportDir, "ANOVA_Pairwise_Adjusted_Interaction.csv"), row.names = TRUE)
      write.csv(as.data.frame(pairwiseResults[["Interaction"]]$unadjusted_p.value), 
              file.path(exportDir, "ANOVA_Pairwise_Unadjusted_Interaction.csv"), row.names = TRUE)
      write.csv(as.data.frame(pairwiseResults[[group]]$adjusted_p.value), 
                file.path(exportDir, paste0("ANOVA_Pairwise_Adjusted_", group, ".csv")), row.names = TRUE)
      write.csv(as.data.frame(pairwiseResults[[group]]$unadjusted_p.value), 
                file.path(exportDir, paste0("ANOVA_Pairwise_Unadjusted_", group, ".csv")), row.names = TRUE)
  }
} else {
  cat("CSV files will not be exported.\n")
}
