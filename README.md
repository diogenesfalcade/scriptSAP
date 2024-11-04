# Python-Routine

### SAP data extraction and cleaning bot, with integration and use of VBScript
### The process is to access SAP with an automation that accesses the RP system locally
### Within the repository we find both the bot (based on OS), as well as the extraction scripts (VBScript), in addition to the cleaning (Pandas & Pure Python)

### You should only run the script through main and the folder structure should be kept the same
### The prerequisite for running the script is to keep the PMO & Planning Team folder synchronized with your PC,
### so that it is possible to access it directly through Windows Explorer

### NEW PROJECTS
### 1. Add the Project ID (6 digits - NNNNNN) followed by a 'tab' (\t) with the project name - otherwise the bot will not identify it
### 2. Classify ALL work packages with the WBS (NNNNNN-XXBR) followed by 'tab'(\t) with your ProjectID (6 digits - NNNNNN) and finally,
### also followed by tab, the name of the cell, which must exactly follow the name entered in the other cells of the other projects
