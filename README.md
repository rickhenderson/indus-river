# Indus River Rainfall over 53 Years
A collection of data and analyses based on the Indus River in Pakistan.

Data Description: 53 years of daily rainfall readings from the sub-basin of the Indus River in Pakistan. Measurements are in m3/s (cumecs).

<a href="yearly.csv">yearly.csv</a> - Yearly totals calculated from the original dataset.

## Description
This project originally started when answering a question on an Internet forum about creating a heatmap in Excel using the original data.

The Visual Basic for Applications (VBA) was the first code written to perform the analysis, and the heat map was generated using conditional formatting in an Excel worksheet.

## Analysis Code
<a href="rainfall53.bas">rainfall53.bas</a> - the original VBA code that generated data on an Excel worksheet that could be used to make a heatmap using Excel's Conditional Formatting feature.

<a href="rainfall53.R">rainfall.R</a> - R code to perform a similar analysis.

## Discussion
The major problem working with this data set was date conversion from the original format to a format that worked correctly with the locale of my copy of Excel. This problem recurs in Excel quite frequently, and special care needs to be taken when data is retrieved from a different country where date formats may be different. For example, with improper date handling, data for January 2nd of a given year may assigned to February 1st of the same year, which would lead to incorrect results which would be hard to find.

A way to approach this problem in Excel may be to use Excel's new <b>Data Model</b> features and include a <i>Calendar table</i>, so that date extraction and filtering can be simplified.

### Further Reading
[Pakistan unveils plan to tackle looming water crisis](http://www.climatechangenews.com/2016/03/23/pakistan-unveils-plan-to-tackle-looming-water-crisis) - 2016/03/23
