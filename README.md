# Nmap XML to Word Report Converter

This Python program takes an Nmap XML output and converts it into a formatted Word document with detailed information about open ports, services, and potential vulnerabilities found during the scan. The Word document is generated in a table-based layout, making it easy to review and add additional notes.

## Features
- Parses Nmap XML output.
- Extracts and presents open ports, services, and product details.
- Displays fingerprint and script data for each port, if available.
- Supports host and port-specific notes.
- Generates a clear, structured Word document with the results.

## Prerequisites

- Python 3.x
- The `python-docx` library for generating Word documents. Install it via pip:
  ```bash
  pip install python-docx
  ```
  Or with apt
  ```bash
  apt install python3-docx
  ```

## How to Use

### 1. Run Nmap with the Following Command

To generate the XML report needed for the script, use Nmap with the following command in Kali Linux:

```bash
sudo nmap -sS -p- -sV --script vuln -iL ips.txt -T4 -oA my_report
```

This command performs the following:
- `-sS`: Conducts a TCP SYN scan.
- `-p-`: Scans all 65535 ports.
- `-sV`: Attempts to determine service versions.
- `--script vuln`: Runs vulnerability scripts on each open port to identify potential issues.
- `-iL ips.txt`: Reads a list of IPs or hostnames from the file `ips.txt`.
- `-T4`: Increases scan speed, adjusting timing to make it faster.
- `-oA my_report`: Saves the output in all formats (normal, XML, and Grepable) with the basename `my_report`.

### 2. Creating the IP List File

To scan multiple IPs or hosts, put each IP address or hostname on a new line in a text file called `ips.txt`. For example:

```text
192.168.1.1
192.168.1.2
192.168.2.1
example.com
```

### 3. Run the Python Program

Once the Nmap scan is complete and you have an XML file (e.g., `my_report.xml`), run the Python program to convert the report into a Word document. 

```bash
python parsereport.py my_report.xml my_report.docx
```

This will create a `my_report.docx` file that contains a structured and formatted report of the Nmap scan results.

### Example Usage

```bash
# Step 1: Create a list of IPs in a file called ips.txt
echo -e "192.168.1.1\n192.168.1.2" > ips.txt

# Step 2: Run the Nmap scan
sudo nmap -sS -p- -sV --script vuln -iL ips.txt -T4 -oA my_report

# Step 3: Convert Nmap XML output to a Word document
python parsereport.py my_report.xml my_report.docx
```

### Example Output

The resulting Word document will contain:
- Host information such as IP address, status, and optionally hostname.
- Port details, including service name, product, version, and method.
- Fingerprint and additional info like script output, if available.

### Notes

- Ensure that `ips.txt` is in the same directory as the Nmap command unless you provide an absolute path.
- The Nmap scan could take time depending on the number of IPs, ports, and scripts used. Use the `-T4` option to speed up the scan.

## Contributing

If you want to contribute or suggest improvements, feel free to fork the repository and create a pull request!

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
