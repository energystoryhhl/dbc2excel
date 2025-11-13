# DBC2Excel

A CAN bus DBC (Database CAN) to Excel converter tool with GUI interface.

## Description

This tool converts DBC files (commonly used in automotive CAN bus networks) to Excel format for easy viewing and analysis. The DBC format contains CAN message and signal definitions, while the output Excel file provides a structured matrix view of all messages and signals with their properties.

## Features

- **GUI Interface**: User-friendly wxPython-based graphical interface
- **Flexible Export Options**: Configurable output with various optional fields
- **Signal Descriptions**: Extract and include signal descriptions from DBC files
- **Signal Value Descriptions**: Include enumerated signal value descriptions
- **Initial Values**: Export initial/default values for signals
- **Sender/Receiver Information**: Display transmitter and receiver nodes
- **Sorting Options**: Sort CAN IDs in ascending or descending order
- **Customizable Text Length**: Set maximum length for signal value descriptions

## Requirements

- Python 3.x
- wxPython
- xlwt (for Excel file generation)

## Installation

1. Clone or download the project files
2. Install required dependencies:
   ```bash
   pip install wxpython xlwt
   ```

## Usage

### GUI Mode (Recommended)

1. Run the main application:
   ```bash
   python dbc2excel_main.py
   ```

2. Use the GUI interface to:
   - Select a DBC file using "选择dbc文件" (Select DBC File) button
   - Configure export options using checkboxes:
     - **Generate Signal Description**: Include signal descriptions
     - **Generate Signal Value Description**: Include enumerated value descriptions  
     - **Generate Initial Value**: Include initial/default values
     - **Generate Sender and Receiver**: Include transmitter/receiver information
     - **Ascending Sort**: Sort CAN IDs (uncheck for descending order)
   - Set maximum text length for signal value descriptions
   - Click "生成Excel文件" (Generate Excel File) to convert

### Command Line Mode

```python
import dbc2excel as d2e

# Create DBC loader instance
dbc = d2e.DbcLoad('your_file.dbc')

# Convert to Excel with options
dbc.dbc2excel(
    filepath='your_file.dbc',
    if_sig_desc=True,           # Include signal descriptions
    if_sig_val_desc=True,       # Include signal value descriptions
    val_description_max_number=70,  # Max length for value descriptions
    if_start_val=True,          # Include initial values
    if_recv_send=True,          # Include sender/receiver info
    if_asc_sort=True            # Sort ascending (True) or descending (False)
)
```

## Output Format

The generated Excel file contains a matrix with the following columns:

- **Message Information**: Name, Type, ID, Send Type, Cycle Time, Length
- **Signal Information**: Name, Description, Byte Order, Start Byte/Bit, Send Type, Bit Length
- **Data Properties**: Data Type, Resolution, Offset, Min/Max Values (Physical & Hex)
- **Values**: Initial Value, Invalid Value, Inactive Value, Unit
- **Additional Info**: Signal Value Descriptions, Cycle Times, Sender/Receiver mapping

## File Structure

```
dbc2excel/
├── dbc2excel_main.py    # Main GUI application
├── dbc2excel.py         # Core DBC parsing and Excel generation logic
├── README.md            # This file
├── makefile.txt         # Build configuration
└── backup/              # Backup files
    ├── dbc2excel_main.py
    └── dbc2excel.py
```

## Version History

### 2018/11/14
- Added CAN ID sorting functionality
- Fixed start bit calculation bug for Motorola byte order (MSB)

## Technical Details

### Supported DBC Features

- **BO_ (Message Objects)**: CAN message definitions with ID, name, size, and transmitter
- **SG_ (Signal Objects)**: Signal definitions with bit position, length, byte order, scaling
- **CM_ SG_ (Signal Comments)**: Signal descriptions and documentation
- **VAL_ (Value Descriptions)**: Enumerated signal value meanings
- **BA_ "GenSigStartValue"**: Initial/default signal values
- **BA_ "GenMsgCycleTime"**: Message transmission cycle times

### Byte Order Handling

The tool correctly handles both Intel (little-endian) and Motorola (big-endian, MSB) byte ordering, automatically calculating the correct start bit positions for signals.

## License

This project is provided as-is for educational and development purposes.

## Contributing

For issues, questions, or contributions, please refer to the original blog post: 
https://blog.csdn.net/hhlenergystory/article/details/80443454

## Author

Created by Huang Honglei (黄洪磊)
