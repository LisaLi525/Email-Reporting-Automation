```markdown
## Overview
This R script automates the process of sending scheduled emails for various reports related to retail operations at BookXchange. The script utilizes the RDCOMClient package to interact with Microsoft Outlook, attaching relevant reports and sending emails to specified recipients.

## Prerequisites
Make sure to install the required R packages before running the script:

```r
install.packages(c(
  "RPostgreSQL", "dplyr", "dbplyr", "data.table",
  "lubridate", "reshape2", "stringr", "readxl",
  "writexl", "openxlsx", "tidyverse", "RDCOMClient"
))
```

## Email Configuration
Configure the recipient email addresses before running the script:

```r
recipient_internal = "internal@example.com"
recipient_team = c(
  "team_member1@example.com", "team_member2@example.com", "team_member3@example.com"
)
```

## Tasks

### Inventory Report - Daily (Internal)
This task sends the daily inventory report to internal recipients.

```r
# ... [Code for Internal Inventory Report Email]
```

### Inventory Report - Daily (Team)
This task sends the daily inventory report to the team with additional details.

```r
# ... [Code for Team Inventory Report Email]
```

<!-- Add similar sections for other tasks... -->

## Note
Replace sensitive information such as file paths and email addresses before running the script in your environment.

Author: Example User
Email: example.user@example.com
```

This README provides an overview of the script, lists prerequisites, guides on email configuration, and includes sections for each task with corresponding code snippets. Ensure to replace example emails, file paths, and any confidential information before running the script.
