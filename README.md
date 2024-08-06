# VBA Automation Module

This VBA module contains a collection of subroutines and functions designed to automate various tasks in Microsoft Excel, particularly focused on data processing, file management, and interaction with external applications like Sage and Chrome.

## Key Features:

1. **Mouse and Keyboard Control**: Functions for simulating mouse movements, clicks, and keyboard inputs.

2. **Window Management**: Routines for manipulating application windows (e.g., maximizing, focusing).

3. **Chrome Automation**: Functions for opening Chrome, navigating to specific URLs, and verifying page loads.

4. **Excel File Operations**: Subroutines for saving Excel files, formatting sheets, and moving data between sheets.

5. **Data Cleaning and Validation**: Functions for cleaning strings and validating purchase order numbers.

6. **Directory Navigation**: Routines for locating specific directories and files within a folder structure.

7. **Sage Integration**: Functions for interacting with Sage software, including error checking.

8. **Subcontract and PO Processing**: Logic for handling different types of purchase orders and subcontracts.

## Main Components:

- Mouse movement and click simulations
- Excel worksheet formatting and data manipulation
- File system operations
- Chrome browser automation
- Sage software integration
- Purchase Order (PO) validation and processing
- Error handling and user prompts

## Usage Notes:

This module is designed to work with specific Excel workbooks and external applications. It includes hard-coded screen coordinates for mouse movements, which may need adjustment based on screen resolution and application versions.

## Caution:

The code includes automation of mouse movements and keyboard inputs. Use with caution and ensure proper testing before implementation in a production environment.
