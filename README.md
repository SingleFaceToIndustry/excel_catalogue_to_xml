# Excel to XML Converter for SFTI Product Catalogue template

## Overview

This Python-based tool converts product catalogues from MS Excel format (using with the SFTI template) into XML files compliant with the "Peppol BIS Catalogue 3 (OASIS UBL)" standard. 
It's designed to facilitate the simple way of creating standardized XML catalogues, enhancing interoperability between buyers and seller within the Peppol network.

## Features

- **Conversion of Excel to XML:** Supports converting an Excel file to XML format following the Peppol BIS Catalogue specifications.
- **Flexible Input Options:** Accepts both file paths and byte arrays of Excel files as input.
- **Maximum Number of Line Configuration:** Allows setting a maximum number of lines to be processed from the Excel file.
- **Enveloping in Peppol SBDH:** Chose by a setting in the Excel spreadsheet template

## Prerequisites

Before using this converter, ensure you have the following installed:
- Python 3.11
- openpyxl

## Installation

Clone this repository to your local machine using:

```bash
git clone https://github.com/SingleFaceToIndustry/excel_catalogue_to_xml.git
```

Navigate into the project directory and install the required dependencies:

```bash
cd excel_catalogue_to_xml
pip install -r requirements.txt
```

## Usage

To convert an Excel file to XML, invoke the `excel_to_xml` function with the path to your Excel file or a byte array of the Excel file. Optionally, you can specify a maximum number of lines to process.

```python
from excel_catalogue_to_xml import excel_to_xml

# Example using a file path
excel_to_xml('/path/to/your/excel-file.xlsx', max_lines=100)

# Example using a byte array
with open('/path/to/your/excel-file.xlsx', 'rb') as file:
    excel_bytes = file.read()
    excel_to_xml(excel_bytes, max_lines=100)
```

## Contributing

SFTI (Single Face To Industry) maintains this code. We welcome contributions and input from the community. If you have suggestions, bug reports, or enhancements, please submit them in the issues or discussions section of this repository.

## License

This project is open-source and available under  Apache License v2, see the LICENSE file for more details.

## Disclaimer

The code is provided "as is", and SFTI takes no responsibility for any errors or unintentional results arising from its use.

## Contact

For further information or assistance, please open an issue in this repository, and a maintainer will get back to you.

---