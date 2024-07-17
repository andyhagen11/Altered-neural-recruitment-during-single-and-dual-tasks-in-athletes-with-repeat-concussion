#Simple script to run a two-way (repeated measures and group) MANVOVA or MANCOVA for four different dependent variables

library(stats)  
library(psych)  
library(rstudioapi)  
rm(list=ls())

#Print warnings as they come up for troubleshooting purposes
options(warn=1)

# Import Data in long format
Data <- read_excel("/Users/andyhagen/Library/CloudStorage/OneDrive-Colostate/Colorado State/SNL/SRC Dual Task fNIRS/Concussion (n=20) and Controls (n=17) DTS LE Results.xlsx", sheet = "Long Form Data")

# Specify within_factor and between_factor as variables
Variable1 = readline("Enter the first variable for RM-MANOVA: ")
Variable2 = readline("Enter the second variable for RM-MANOVA: ")
Variable3 = readline("Enter the third variable for RM-MANOVA: ")
Variable4 = readline("Enter the fourth variable for RM-MANOVA: ")
within_factor <- "Task"
between_factor <- "Group"

# Combine the variables into a formula
formula_string <- paste("cbind(", Variable1, ",", Variable2, ",", Variable3, ",", Variable4, ") ~", within_factor, "*", between_factor)

# Create the model using MANOVA
Manova_Model <- manova(as.formula(formula_string), data = Data)

# Get the MANOVA results
Manova_Results <- summary(Manova_Model, test = "Pillai")  

# Generate Summary Statistics for each variable
SummaryStats_Variable1 <- describeBy(Data[[Variable1]], group = list(Data[[within_factor]], Data[[between_factor]]), mat = TRUE)
SummaryStats_Variable2 <- describeBy(Data[[Variable2]], group = list(Data[[within_factor]], Data[[between_factor]]), mat = TRUE)
SummaryStats_Variable3 <- describeBy(Data[[Variable3]], group = list(Data[[within_factor]], Data[[between_factor]]), mat = TRUE)
SummaryStats_Variable4 <- describeBy(Data[[Variable4]], group = list(Data[[within_factor]], Data[[between_factor]]), mat = TRUE)

# If main effect is significant, perform pairwise tests for each variable
Pairwise_Results_Variable1 <- pairwise.t.test(Data[[Variable1]], interaction(Data[[within_factor]], Data[[between_factor]]), p.adj = "none", paired = FALSE)
Pairwise_Results_Variable2 <- pairwise.t.test(Data[[Variable2]], interaction(Data[[within_factor]], Data[[between_factor]]), p.adj = "none", paired = FALSE)
Pairwise_Results_Variable3 <- pairwise.t.test(Data[[Variable3]], interaction(Data[[within_factor]], Data[[between_factor]]), p.adj = "none", paired = FALSE)
Pairwise_Results_Variable4 <- pairwise.t.test(Data[[Variable4]], interaction(Data[[within_factor]], Data[[between_factor]]), p.adj = "none", paired = FALSE)

# # This script has the capability to do pairwise comparisons from a multivariate model, however please ensure these are the correct comparisons you are looking for!
# # Combine p-values from all pairwise tests
# all_pvalues <- c(Pairwise_Results_Variable1$p.value, Pairwise_Results_Variable2$p.value, Pairwise_Results_Variable3$p.value, Pairwise_Results_Variable4$p.value)
# 
# # Adjust p-values using False Discovery Rate (FDR) correction
# adjusted_pvalues <- p.adjust(all_pvalues, method = "fdr")
# 
# # Assign both unadjusted and adjusted p-values back to pairwise results
# Pairwise_Results_Variable1$unadjusted_p.value <- Pairwise_Results_Variable1$p.value
# Pairwise_Results_Variable2$unadjusted_p.value <- Pairwise_Results_Variable2$p.value
# Pairwise_Results_Variable3$unadjusted_p.value <- Pairwise_Results_Variable3$p.value
# Pairwise_Results_Variable4$unadjusted_p.value <- Pairwise_Results_Variable4$p.value
# 
# Pairwise_Results_Variable1$p.value <- adjusted_pvalues[1:length(Pairwise_Results_Variable1$p.value)]
# Pairwise_Results_Variable2$p.value <- adjusted_pvalues[(length(Pairwise_Results_Variable1$p.value)+1):(length(Pairwise_Results_Variable1$p.value)+length(Pairwise_Results_Variable2$p.value))]
# Pairwise_Results_Variable3$p.value <- adjusted_pvalues[(length(Pairwise_Results_Variable1$p.value)+length(Pairwise_Results_Variable2$p.value)+1):(length(Pairwise_Results_Variable1$p.value)+length(Pairwise_Results_Variable2$p.value)+length(Pairwise_Results_Variable3$p.value))]
# Pairwise_Results_Variable4$p.value <- adjusted_pvalues[(length(Pairwise_Results_Variable1$p.value)+length(Pairwise_Results_Variable2$p.value)+length(Pairwise_Results_Variable3$p.value)+1):length(adjusted_pvalues)]

# Export CSVs if desired
export_csv <- readline("Do you want to export the CSV files? (Y/N): ")
if (export_csv == "Y") {
  # Choose export directory
  export_dir <- selectDirectory()
  
  # Write CSVs for summary statistics
  write.csv(SummaryStats_Variable1, file.path(export_dir, paste0("SummaryStats_", Variable1, ".csv")), row.names = TRUE)
  write.csv(SummaryStats_Variable2, file.path(export_dir, paste0("SummaryStats_", Variable2, ".csv")), row.names = TRUE)
  write.csv(SummaryStats_Variable3, file.path(export_dir, paste0("SummaryStats_", Variable3, ".csv")), row.names = TRUE)
  write.csv(SummaryStats_Variable4, file.path(export_dir, paste0("SummaryStats_", Variable4, ".csv")), row.names = TRUE)
  
  # Write CSVs for MANOVA results
  # Convert MANOVA summary to data frame
  Manova_Results_df <- as.data.frame(Manova_Results$stats)
  write.csv(Manova_Results_df, file.path(export_dir, "MANOVA_Table.csv"), row.names = TRUE)
  
  # Write CSVs for pairwise comparisons - make sure to select either the unadjusted or adjusted p-values (this example is unadjusted)
  write.csv(as.data.frame(Pairwise_Results_Variable1$p.value), file.path(export_dir, paste0("Pairwise_", Variable1, ".csv")), row.names = TRUE)
  write.csv(as.data.frame(Pairwise_Results_Variable2$p.value), file.path(export_dir, paste0("Pairwise_", Variable2, ".csv")), row.names = TRUE)
  write.csv(as.data.frame(Pairwise_Results_Variable3$p.value), file.path(export_dir, paste0("Pairwise_", Variable3, ".csv")), row.names = TRUE)
  write.csv(as.data.frame(Pairwise_Results_Variable4$p.value), file.path(export_dir, paste0("Pairwise_", Variable4, ".csv")), row.names = TRUE)
  
} else {
  cat("CSV files will not be exported.\n")
}
