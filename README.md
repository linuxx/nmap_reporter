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
  Or install using apt
  ```bash
  apt install python3-docx
