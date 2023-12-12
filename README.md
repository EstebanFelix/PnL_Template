# Automated P&L generator tool

Automate the generation of financial reports using Excel macros.

## Overview

This tool simplifies the process of generating financial reports by automating the data retriving from the company ERP system that's read through a combintation of tables (facts and dimentiontals) that result in a SQL View.

## Features

- **Dynamic Report Generation:** Automatically creates financial reports based on the provided data model.
- **User-Friendly Interface:** Designed with a user-friendly interface for easy navigation and execution.

## Dependencies

- This tool relies on specific Excel configurations, including CubeValue formulas and connections to an SQL view.
- The .bas is configured to work with a particular data model with a pre build excel worksheets that contains cubevalue formulas that reads the data model.
- Ensure that the required data sources, connections, and data model configurations are available in your environment.

## Prerequisites

- Microsoft Excel
- Access to the specified SQL view
- Understanding of the data model and CubeValue formulas used in the tool

## Getting Started

1. Download the repository.
2. Ensure that the necessary data connections are available.
3. Enable macros if prompted.
4. Run the macro `Run_PnLs` to generate financial reports.

## Notes

- The Excel file is configured for a specific data model and SQL view. Running the tool on a different machine/excel file may require adjusting connections and configurations.
