📄 IP to Country Lookup in Excel Using IP2Location CSV

Overview
This guide shows you how to use Excel to resolve a list of IP addresses to their corresponding countries using a static CSV file from IP2Location.

🗂️ Folder Structure
├── ip2location.csv ← IP range database (CSV format)
├── ip_list.xlsx ← Your Excel file with a list of IPs
└── README.md ← This documentation

📌 Requirements

Microsoft Excel (2016 or newer recommended)

Basic familiarity with Excel formulas and VBA

IP2Location CSV file (IP ranges in numeric format)

📋 Sample IP2Location CSV Format
The CSV file should have the following structure:

ip_from ip_to country_code country_name
16777216 16777471 US United States
16777472 16778239 CN China
📄 Step-by-Step Setup

Prepare your IP List
In your Excel file, list IP addresses (e.g. 8.8.8.8) in a column, e.g., column A.

Import the IP2Location CSV into Excel

Open the CSV file in Excel.

Format it as a table (Ctrl+T).

Make sure the columns are ordered: ip_from (col 1), ip_to (col 2), country_code (col 3), country_name (col 4).

Name the table (e.g., IPRanges).

Add VBA Functions
Press Alt+F11 → Insert → Module → Paste the following code:

Function to Convert IP to Number:

vba
Copy
Edit
Function IPToLong(IP As String) As Double
Dim bytes() As String
If InStr(IP, ".") = 0 Then
IPToLong = -1
Exit Function
End If
bytes = Split(IP, ".")
If UBound(bytes) <> 3 Then
IPToLong = -1
Exit Function
End If
IPToLong = bytes(0) _ 16777216# + bytes(1) _ 65536# + bytes(2) \* 256# + bytes(3)
End Function
Function to Find Country from IP Range:

vba
Copy
Edit
Function FindCountry(ipNumber As Double, lookupRange As Range) As String
Dim r As Range
For Each r In lookupRange.Rows
If ipNumber >= r.Cells(1, 1).Value And ipNumber <= r.Cells(1, 2).Value Then
FindCountry = r.Cells(1, 4).Value ' Returns country_name; change to (1,3) for code
Exit Function
End If
Next r
FindCountry = "Not found"
End Function
Use the Functions in Excel
Assuming:

Your IP address is in cell A2

IP2Location table is in Sheet2 from A2:D100000

Formula:

excel
Copy
Edit
=FindCountry(IPToLong(A2), Sheet2!A2:D100000)
📎 Notes

Use country code instead of name by changing r.Cells(1, 4) to r.Cells(1, 3).

For large datasets, Excel may become slow — consider using Power Query or a database.

📦 Optional: Use Power Query
If you're using Excel 365/2019+:

Load both the IP list and CSV into Power Query

Convert IPs to numbers using custom column

Merge queries by checking if ip_number falls between ip_from and ip_to (requires a custom join script)

🙌 Credits
Data provided by IP2Location.com
Script written by Muhamad Yusuf
