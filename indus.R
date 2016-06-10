# indus.R
# Indus River Sub-Basis Data Analysis
# Web source for yearly data at:
# https://github.com/rickhenderson/indus-river/blob/master/yearly.csv

yearly_data <- read.csv("yearly.csv")

monthly_subtotals <- read.csv("monthly.csv")

# Create a data matrix from the values
dataMatrix <- matrix(monthly_subtotals, nrow=52)

# Create a basic plot for Yearly data
plot(yearly_data, type="l", col="blue", xlab="Year", ylab="Flow Volume m^3/s (cumecs)")
