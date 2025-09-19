# MultiTypeProcessor

MultiTypeProcessor is a Python utility designed to process mixed data from multiple sources—such as IP addresses, domains, file hashes, and URLs. It cleans and classifies the data, merges new entries with existing records, writes the consolidated data to dedicated Excel files, and can retrieve threat analysis information from the VirusTotal API.
 Features

- Reads mixed data from a main text file and additional TXT files in a specified folder.
- Cleans and normalizes input data (e.g., removing duplicates, converting obfuscated URLs or domains).
- Validates data types:
  - IP address
  - Domain names
  - File hashes (MD5, SHA1, SHA256, SHA512)
  - URLs
- Merges unique data into a single main text file.
- Exports each data type to a separate Excel file with columns for threat analysis statistics.
- Integrates with the VirusTotal API to query threat information and update the Excel files.
 Requirements

- Python 3.7+
- The following Python packages:
  - requests
  - openpyxl
  - (Optional) glob, re, os, time, random, urllib.parse (standard libraries)

Install the required packages using pip:

  pip install requests openpyxl
 Configuration

Before running the script, update the configuration in the MultiTypeProcessor class:

- Set the paths:
  - `data_folder`: Folder containing additional TXT files with data.
  - `main_txt`: Path to the main text file that contains mixed data.
  - Excel file paths for each data type (ip, domain, hash, url).
- Update the API key for VirusTotal:
  - Replace `"YOURAPIKEY"` with your actual VirusTotal API key.
- Adjust any other settings (headers, endpoints, etc.) as needed.
 Usage

1. Clone the repository:

    git clone https://github.com/Metehan-Uluocak/MultiTypeProcessor.git
    cd MultiTypeProcessor

2. Update the configuration in the code (file paths and API key) to suit your environment.

3. Run the script:

    python vt-to-excel.py

   The tool will:
   - Remove duplicate lines in the main TXT file.
   - Merge and classify new data from TXT files in the specified folder.
   - Write unique data to Excel files for each data type.
   - Optionally, query VirusTotal and update threat analysis information in the Excel files.


Feel free to adjust any sections to better fit your project specifics or additional functionalities. Happy coding!
