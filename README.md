# Automated P&L generator tool

Automate the generation of financial reports using Excel macros.

## Overview

This tool simplifies the process of generating financial reports by automating repetitive tasks in Microsoft Excel.

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
2. Open the Excel file (`FinancialReportingAutomationTool.xlsm`) in a compatible environment.
3. Ensure that the necessary data connections are available.
4. Enable macros if prompted.
5. Run the macro `Run_PnLs` to generate financial reports.

## Notes

- The Excel file is configured for a specific data model and SQL view. Running the tool on a different machine may require adjusting connections and configurations.
- For assistance with setting up the required dependencies, refer to [Documentation Link] or contact the project contributors.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.