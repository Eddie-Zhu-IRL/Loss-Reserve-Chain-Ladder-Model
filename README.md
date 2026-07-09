# :rocket: Loss Development Reserving using the Chain Ladder Method
## An actuarial reserving case study implemented independently in Excel, VBA, Python and R


## Overview
This project demonstrates the implementation of a deterministic Chain Ladder reserving model using four independent approaches:

- Microsoft Excel Built-in Functions
- Excel VBA
- Python
- R

The objective is to estimate future claim liabilities by calculating:

- Loss Development Triangles
- Loss Development Factors (LDFs)
- Age-to-Ultimate Factors
- Ultimate Loss
- Outstanding Case Reserves
- Incurred But Not Reported (IBNR)

Each implementation follows the same actuarial methodology and produces consistent reserving estimates, providing validation across multiple analytical platforms.

## Business Problem
General insurers establish reserves to ensure sufficient funds are available to meet future claim payments arising from policies already written.

Since many claims remain unsettled or have not yet been reported, actuaries estimate future liabilities using historical claims development experience.

This project demonstrates one of the most widely used deterministic reserving techniques:

### Chain Ladder Method
The workflow estimates future claim development from historical loss triangles to calculate ultimate losses and IBNR reserves.

## Project Objectives

This project aims to:

- Construct cumulative loss development triangles from historical claims data.
- Calculate Age-to-Age Loss Development Factors.
- Estimate Age-to-Ultimate Factors.
- Project Ultimate Losses using the Chain Ladder method.
- Estimate Outstanding Reserves and IBNR.
- Validate reserving calculations by independently implementing the methodology in Excel, VBA, Python and R.

## Dataset

This project uses the Casualty Actuarial Society (CAS) PP Auto Data Set (.csv), updated December 2025.

- Data Source

The dataset is published by the Casualty Actuarial Society (CAS) as part of its loss reserving research resources. The data are extracted from NAIC Schedule P and are commonly used for actuarial loss reserving analysis and research.

Official source:
https://www.casact.org/publications-research/research/research-resources/loss-reserving-data-pulled-naic-schedule-p

- Dataset Usage

The ppauto_pos98-07 dataset is used as the input data source for Excel-based loss reserving model. The dataset is imported into Excel through Power Query connection, where it is transformed and loaded for subsequent actuarial analysis. Dataset is also analyzed in R and Python. 

The model uses the dataset to develop loss triangles and perform reserving analysis, including:

  - Data validation and preparation
  - Loss development analysis
  - Reserve estimation
  - Model output reconciliation with R and Python implementations

- Data Loading

The dataset file is not included in this repository. Users should download the dataset directly from the official CAS website and load it into the workbook through the existing Power Query connection.

After loading the data, follow steps listed on Cover Page and refresh the workbook queries before running the analysis.

   - Dataset Version
   - Dataset: PP Auto Data Set (.csv)
   - Source: Casualty Actuarial Society (CAS)
   - Version: Updated December 2025
  
### 1. Loss Development Triangle: Excel Built-in Functions
This implementation demonstrates a fully dynamic reserving model without requiring manual recalculation.

Features include:
- Dynamic loss development triangle
- Dynamic accident year expansion
- Named Ranges
- INDEX/MATCH
- SUMIFS
- OFFSET
- Conditional Formatting
- Automatic recalculation of reserving metrics

Triangle
LDF table
Ultimate Loss table

### 2. Loss Development Triangle: VBA
The VBA version automates Loss Triangle and LDF workflow.

Functions include:

- Loss Data: user can use dynamic data filtering from dropdown list, and extract filtered data for further analysis
- Loss Triangle: create triangle
- LDF: Arithmetic AVG & Vol-Weighted AVG: calculate LDF using simple arithmetic average and volume-weighted average
- Summary: consolidate earned premiums, incurred losses, cumulative paid loss, calculted Age-to-Ult LDF, Ult Losses, IBNR

Screenshots:
- Triangle construction
  ![pic1](../Figures/Summary_Excel.png)
- Development factors
- Ultimate losses
- IBNR estimates

### 3. Python Implementation
Implemented using:

- pandas
- numpy
- chainladder

The Python implementation reproduces the reserving calculations using actuarial libraries and provides an independent validation of the Excel and VBA models.

### 4. R Implementation
Implemented using:

- ChainLadder
- tidyverse

The R implementation independently reproduces the reserving calculations and allows comparison with both manual and Python implementations.  

### 5. Validation

One of the objectives of this project is to verify that independent implementations produce consistent actuarial estimates.

| Metric | Excel | VBA | Python | R |
|--------|:-----:|:---:|:------:|:-:|
| Latest Reported Loss | ✓ | ✓ | ✓ | ✓ |
| LDF | ✓ | ✓ | ✓ | ✓ |
| Age-to-Ultimate | ✓ | ✓ | ✓ | ✓ |
| Ultimate Loss | ✓ | ✓ | ✓ | ✓ |
| Outstanding Reserve | ✓ | ✓ | ✓ | ✓ |
| IBNR | ✓ | ✓ | ✓ | ✓ |

The independent implementations produced identical reserving estimates, demonstrating consistency of the methodology across software platforms.

### 6. Business Interpretation
The Chain Ladder method projects future claims development based on historical development patterns.

The resulting estimates help insurers:

- Assess outstanding claim liabilities.
- Estimate future claim payments.
- Support financial reporting.
- Inform capital management.
- Monitor reserve adequacy.

## Contact
Author: Eddie Zhu
