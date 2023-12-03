# Retail Reporting Automation

## Overview
# This R script automates the process of sending scheduled emails for various reports
# related to retail operations at BookXchange. The script utilizes the RDCOMClient
# package to interact with Microsoft Outlook, attaching relevant reports and sending
# emails to specified recipients.

## Prerequisites
# Make sure to install the required R packages before running the script:

install.packages(c(
  "RPostgreSQL", "dplyr", "dbplyr", "data.table",
  "lubridate", "reshape2", "stringr", "readxl",
  "writexl", "openxlsx", "tidyverse", "RDCOMClient"
))

## Email Configuration
recipient_internal = "internal@example.com"
recipient_team = c(
  "team_member1@example.com", "team_member2@example.com", "team_member3@example.com"
)

## Tasks

### Inventory Report - Daily (Internal)
# This task sends the daily inventory report to internal recipients.

library(RDCOMClient)

invetory_report_internal <- "path/to/internal/inventory_report.xlsx"
OutApp <- COMCreate("Outlook.Application")
outMail_internal = OutApp$CreateItem(0)

outMail_internal[["To"]] = recipient_internal
outMail_internal[["subject"]] = "Report | Inventory Reports - Internal"
outMail_internal[["attachments"]]$Add(invetory_report_internal)
outMail_internal[["body"]] = "Hi Team, \n\nPlease see the attached internal inventory report.\n\nThank you,\nExample User"

outMail_internal$Send()

### Inventory Report - Daily (Team)
# This task sends the daily inventory report to the team with additional details.

invetory_report_team <- "path/to/team/inventory_report.xlsx"
OutApp <- COMCreate("Outlook.Application")
outMail_team = OutApp$CreateItem(0)

outMail_team[["To"]] = paste(recipient_team, collapse=";")
outMail_team[["subject"]] = "Report | Inventory Reports - Team"
outMail_team[["attachments"]]$Add(invetory_report_team)
outMail_team[["body"]] = "Hi Team, \n\nPlease see the attached team inventory report with additional details.\n\nThank you,\nExample User"

outMail_team$Send()

# Add similar sections for other tasks...

## Note
# Replace sensitive information such as file paths and email addresses before running
# the script in your environment.

Author: Example User
Email: example.user@example.com
