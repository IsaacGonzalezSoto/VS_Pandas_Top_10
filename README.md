# Faults Analysis and Reporting

This project processes a CSV file containing machine fault data and generates an Excel report with separate sheets for each machine. The report includes tables and bar charts for the top 10 faults by duration and occurrences. Additionally, it includes summary sheets with an index, sorted unique faults per station, and sorted total duration per station.

## Features

- **Data Aggregation**: Aggregates fault data by machine and fault description.
- **Top 10 Faults**: Extracts the top 10 faults by duration and occurrences for each machine.
- **Bar Charts**: Creates bar charts using Matplotlib and inserts them into the Excel report.
- **Summary Sheets**: Includes summary sheets with an index, sorted unique faults, and sorted total duration.
- **PEP8 Compliant**: Follows PEP8 coding standards.
- **Zoom Level**: Sets the zoom level to 50% for all sheets in the Excel report.

## Installation

1. Clone the repository:

    ```sh
    git clone https://github.com/yourusername/faults-analysis.git
    cd faults-analysis
    ```

2. Create a virtual environment and activate it:

    ```sh
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

3. Install the required packages:

    ```sh
    pip install -r requirements.txt
    ```

## Usage

1. Place your `faults.csv` file in the project directory.

2. Run the script:

    ```sh
    python process_faults.py
    ```

3. The script will generate an Excel file named `faults_per_machine.xlsx` in the project directory.

## File Structure

- `process_faults.py`: The main script that processes the CSV file and generates the Excel report.
- `README.md`: This readme file.
- `requirements.txt`: List of required Python packages.

## Sample Data Format

The `faults.csv` file should have the following columns:

- `D_MachineName`: The name of the machine.
- `D_StateDesc`: The state description.
- `D_MsgCode`: The message code.
- `D_MsgDesc`: The message description.
- `T_TotalDuration`: The total duration of the fault.
- `T_TotalOccur`: The total occurrences of the fault.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

## License

This project is licensed under the MIT License.
