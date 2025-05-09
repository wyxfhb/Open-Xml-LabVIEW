# OpenXML LabVIEW

A lightweight wrapper for reading `.xlsx` Excel files using Open XML.

## Prerequisites

Before using this library, ensure you have the following dependencies:

- **LUnit Test Framework** – [GitHub Repository](https://github.com/Astemes/astemes-lunit)
- **Open XML SDK** – Download the following DLLs via NuGet or other sources:
  - `DocumentFormat.OpenXml.dll`
  - `DocumentFormat.OpenXml.Framework.dll`
  
  **Note:** It is recommended to place these DLLs next to each other within your project folder.

## Installation

1. Clone the repository along with submodules:
   ```sh
   git clone --recurse-submodules https://github.com/pettaa123/Open-Xml-LabVIEW

## Public API
This library provides functions for retrieving cell values and named ranges in an Excel worksheet:

# Get Cell Value
Returns the value of a given cell, handling numeric, boolean and string types.

![readcell](sample get cell value.png)

# Get Named Range of Sheet
Retrieves a specified named range from a worksheet.

# Get Named Ranges of Sheet
Lists all named ranges defined within a worksheet.