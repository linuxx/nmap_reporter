import xml.etree.ElementTree as ET
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Function to parse Nmap XML and return structured scan data
def parse_nmap_xml(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    scan_data = []

    # Iterate over each host in the Nmap output
    for host in root.findall('host'):
        host_data = {}
        host_data['address'] = host.find('address').attrib.get('addr', 'N/A')
        host_data['status'] = host.find('status').attrib.get('state', 'N/A')
        
        # Safely retrieve the hostname if it exists
        hostname_elem = host.find('hostnames/hostname')
        host_data['hostname'] = hostname_elem.attrib.get('name', 'N/A') if hostname_elem is not None else 'N/A'
        
        ports = []
        # Process each port under the host
        for port in host.findall('./ports/port'):
            service_elem = port.find('./service')
            
            # Build the port data dictionary
            port_data = {
                'portid': port.attrib.get('portid', 'N/A'),
                'protocol': port.attrib.get('protocol', 'N/A'),
                'service': service_elem.attrib.get('name', 'N/A') if service_elem is not None else 'N/A',
                'product': service_elem.attrib.get('product', 'N/A') if service_elem is not None else 'N/A',
                'version': service_elem.attrib.get('version', 'N/A') if service_elem is not None else 'N/A',
                'extrainfo': service_elem.attrib.get('extrainfo', 'N/A') if service_elem is not None else 'N/A',
                'method': service_elem.attrib.get('method', 'N/A') if service_elem is not None else 'N/A',
                'state': port.find('./state').attrib.get('state', 'N/A'),
                'fingerprint': None,
                'info': ''
            }

            # Collect script output for 'Info' and 'Fingerprint'
            script_output = ''
            for script in port.findall('script'):
                script_id = script.attrib.get('id', 'N/A')
                script_text = f"{script_id}: {script.attrib.get('output', 'N/A')}"
                
                if script_id == 'fingerprint-strings':
                    port_data['fingerprint'] = script.attrib.get('output', 'N/A')
                else:
                    script_output += f"{script_text}\n"
            
            if script_output.strip():
                port_data['info'] = script_output.strip()
            
            ports.append(port_data)

        host_data['ports'] = ports
        scan_data.append(host_data)

    return scan_data


# Helper function to set text and formatting in a Word cell
def set_cell_text(cell, text, font_size=10, bold=False, color=RGBColor(0, 0, 0)):
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(text)
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color

# Helper function to add borders to a Word table
def add_borders(table):
    tbl = table._element
    tblBorders = OxmlElement('w:tblBorders')
    
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')  # Thinner borders
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')
        tblBorders.append(border)
    
    tbl.tblPr.append(tblBorders)

# Function to create the Word report from scan data
def create_nmap_report(scan_data, output_doc):
    doc = Document()

    # Add the title to the document
    doc.add_heading('Nmap Security Scan Report', 0)

    # Iterate over each host and create the corresponding tables
    for host in scan_data:
        # Create a table for the host's general information and notes
        table = doc.add_table(rows=2, cols=3)
        table.autofit = True
        add_borders(table)

        # Set the host details (address, status, and hostname)
        set_cell_text(table.cell(0, 0), f"Host: {host['address']}", font_size=12, bold=True, color=RGBColor(0x2E, 0x75, 0xB6))
        set_cell_text(table.cell(0, 1), f"Status: {host['status']}", font_size=12, bold=True, color=RGBColor(0x2E, 0x75, 0xB6))
        set_cell_text(table.cell(0, 2), f"Hostname(PTR): {host['hostname']}", font_size=12, bold=True, color=RGBColor(0x2E, 0x75, 0xB6))

        # Add the 'Notes' section spanning all columns
        notes_cell = table.cell(1, 0)
        notes_cell.merge(table.cell(1, 1))
        notes_cell.merge(table.cell(1, 2))
        set_cell_text(notes_cell, 'Notes:', font_size=10, bold=True)
        notes_cell.paragraphs[0].add_run("\n\n\n")  # Add extra space for notes

        # Apply custom shading to the host table cells
        for row in table.rows:
            for cell in row.cells:
                shading = OxmlElement("w:shd")
                shading.set(qn("w:fill"), "FFFFFF" if cell is notes_cell else "DCE6F1")
                cell._element.get_or_add_tcPr().append(shading)

        doc.add_paragraph("")  # Add a space before the port section

        # Add ports information (if any)
        if not host['ports']:
            doc.add_paragraph("No open ports found.")
            continue

        # For each port, create a table for detailed info
        for port in host['ports']:
            # Create a single row for the port header
            port_row = doc.add_table(rows=1, cols=1)
            add_borders(port_row)
            port_cell = port_row.rows[0].cells[0]
            set_cell_text(port_cell, f"Port {port['portid']} ({port['protocol']})", font_size=12, bold=True, color=RGBColor(0x4F, 0x81, 0xBD))
            
            # Apply shading to the port row
            shading = OxmlElement("w:shd")
            shading.set(qn("w:fill"), "D9E1F2")
            port_cell._element.get_or_add_tcPr().append(shading)

            # Create a table for service and product details under the port
            port_info_table = doc.add_table(rows=2, cols=6)
            add_borders(port_info_table)
            
            # Set the headers
            hdr_cells = port_info_table.rows[0].cells
            headers = ['Service', 'Product', 'Version', 'Info', 'Method', 'Vulnerabilities']
            for idx, header in enumerate(headers):
                set_cell_text(hdr_cells[idx], header, font_size=10, bold=True)

            # Populate the port details
            row_cells = port_info_table.rows[1].cells
            set_cell_text(row_cells[0], port['service'], font_size=10)
            set_cell_text(row_cells[1], port['product'], font_size=10)
            set_cell_text(row_cells[2], port['version'], font_size=10)
            set_cell_text(row_cells[3], port['extrainfo'], font_size=10)
            set_cell_text(row_cells[4], port['method'], font_size=10)
            set_cell_text(row_cells[5], 'Yes' if port['info'] else 'No', font_size=10)

            # Add fingerprint if available
            if port['fingerprint']:
                fingerprint_row = port_info_table.add_row()
                fingerprint_cell = fingerprint_row.cells[0]
                # Merge all cells in this row to create one wide cell for the fingerprint
                fingerprint_cell.merge(fingerprint_row.cells[1])
                fingerprint_cell.merge(fingerprint_row.cells[2])
                fingerprint_cell.merge(fingerprint_row.cells[3])
                fingerprint_cell.merge(fingerprint_row.cells[4])
                fingerprint_cell.merge(fingerprint_row.cells[5])
                set_cell_text(fingerprint_cell, f"Fingerprint: {port['fingerprint']}", font_size=10)
                shading_fingerprint = OxmlElement("w:shd")
                shading_fingerprint.set(qn("w:fill"), "F2F2F2")
                fingerprint_cell._element.get_or_add_tcPr().append(shading_fingerprint)

            # Add additional info row if available and not duplicating fingerprint
            if port['info'] and (port['fingerprint'] is None or port['fingerprint'] not in port['info']):
                info_row = port_info_table.add_row()
                info_cell = info_row.cells[0]
                info_cell.merge(info_row.cells[1])
                info_cell.merge(info_row.cells[2])
                info_cell.merge(info_row.cells[3])
                info_cell.merge(info_row.cells[4])
                info_cell.merge(info_row.cells[5])
                set_cell_text(info_cell, f"Info: {port['info']}", font_size=10)
                shading_info = OxmlElement("w:shd")
                shading_info.set(qn("w:fill"), "F5F5F5")
                info_cell._element.get_or_add_tcPr().append(shading_info)

            doc.add_paragraph("")  # Space between ports

    # Save the final Word report
    doc.save(output_doc)


# Main function to handle CLI arguments
def main():
    import argparse
    parser = argparse.ArgumentParser(description="Nmap XML to Word Report Converter")
    parser.add_argument("xml_file", help="Path to the Nmap XML file")
    parser.add_argument("output_doc", help="Path for the output Word document")

    args = parser.parse_args()

    # Parse the XML file and create the report
    scan_data = parse_nmap_xml(args.xml_file)
    create_nmap_report(scan_data, args.output_doc)
    print(f"Report saved as {args.output_doc}")


if __name__ == "__main__":
    main()
