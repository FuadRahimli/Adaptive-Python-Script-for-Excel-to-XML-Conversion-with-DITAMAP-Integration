Adaptive Python Script for Excel to XML Conversion with DITAMAP IntegrationAdaptive Python Script for Excel to XML Conversion with DITAMAP Integration
Jun 2023 - Jun 2023Jun 2023 - Jun 2023

Associated with NokiaAssociated with Nokia
I have written a Python script in accordance with the requirements outlined by the product owner and senior team members. The script's purpose is to generate XML files and a DITAMAP from any given .XLSX Excel file, with the user manually inputting the Excel file name each time.

The script accommodates Excel files of varying dimensions, handling any number of columns or rows. It dynamically processes information in each row, creating an XML file for each row, even accounting for scenarios with empty rows. XML file names are derived from the unique values in the first column of the Excel file, labeled "AlarmId NetAct." The script expects all Excel files to contain a sheet named "Alarm_list," with "AlarmID NetAct" as the key column for XML file names.

Additionally, the script generates a DITAMAP file that references all the XML files, adhering to standard XML and DITAMAP file beginnings. It has been meticulously crafted to meet the specific requirements for compatibility with the Oxygen XML Editor, ensuring proper execution within the editor's working principles.

The implementation leverages various libraries, including "pandas," "numpy," "os," "sys," "codecs," and "xml.dom.minidom." These libraries collectively streamline processes such as data handling, XML document creation, file operations, and command-line input management in the provided code.

As a result, 3059 XML files and 1 DITAMAP were successfully generated. I have rigorously tested the script to confirm its proper execution in Visual Studio and its seamless functionality within the Oxygen XML Editor, affirming that all specified requirements have been met.

Achievements: 
Thanks to this task, I have improved a range of skills including Python Programming, Data Processing, XML Handling, File Operations, Libraries Integration, Problem Solving, Testing, Documentation.
