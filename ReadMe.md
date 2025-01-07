# Pocket Assistant

Pocket Assistant is a Python-based application designed to streamline daily financial management by automating various tasks related to Excel file operations, data entry, and statistical analysis. This application provides a user-friendly interface for creating, modifying, and analyzing Excel files, making it a valuable tool for managing personal or organizational finances efficiently.

## Features

Facilitates the management of daily financial transactions by categorizing them under various tags. Ideal for users unfamiliar with Excel, it offers a straightforward solution to record and organize financial data effectively. 

## Folder Structure

```
Pocket_Assistant/
├── Modify_Excel_File.py          # Script to modify existing Excel files
├── New_Excel_File_Creation.py    # Script to create new Excel files
├── Show_Statistics.py            # Script to compute and display statistics
├── Taking_Tag_Choice.py          # Script to handle tag-based data choices
├── Transaction_Entry.py          # Script to manage transaction entries
├── requirements.txt              # List of dependencies required by the application
└── Pocket_Assistant.exe          # Compiled executable version of the application
```

## Getting Started

### Prerequisites

Make sure you have the following installed:

- Python 3.x
- Pip package manager
- Required libraries (listed in `requirements.txt`)

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/N-Talukder/Pocket_Assistant.git
   cd Pocket_Assistant
   ```
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Usage

You can run the Application using:

```bash
python Pocket_Assistant.py
```

Alternatively, you can use the precompiled executable:

```bash
Pocket_Assistant.exe
```

### Scripts Description

- **Modify\_Excel\_File.py**: This script allows modifications to the Excel file based on user-defined parameters, such as deleting entries, modifying, excluding tags, and including new tags. 
- **New\_Excel\_File\_Creation.py**: Generates new Excel files with specified formats and initial data.
- **Show\_Statistics.py:** Filters and displays data from previous entries in the Excel sheet in separate sheets.
- **Taking\_Tag\_Choice.py**: Handles classifying transactions based on user-specified tags or identifiers while setting up the application.
- **Transaction\_Entry.py**: Facilitates easy and structured transaction data entry to the Excel file.

## Requirements

The application uses the following Python libraries. All required libraries are listed in `requirements.txt`. You can install them using:

```bash
pip install -r requirements.txt
```

## Contributing

Contributions are welcome! Feel free to create an issue or submit a pull request if you have any suggestions or improvements.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.