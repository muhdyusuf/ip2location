# 📄 IP to Country Lookup in Excel Using IP2Location CSV

## Overview

This guide shows you how to use Excel to resolve a list of IP addresses to their corresponding countries using a static CSV file from IP2Location.

---

## 🗂️ Folder Structure

```
.
├── ip2location.csv             ← IP range database (CSV format)
├── ip_list.xlsx                ← Your Excel file with a list of IPs
└── README.md                   ← This documentation
```

---

## 📌 Requirements

- Microsoft Excel (2016 or newer recommended)
- Basic familiarity with Excel formulas and VBA
- IP2Location CSV file (IP ranges in numeric format)

---

## 📋 Sample IP2Location CSV Format

The CSV file should have the following structure:

| ip_from  | ip_to    | country_code | country_name  |
| -------- | -------- | ------------ | ------------- |
| 16777216 | 16777471 | US           | United States |
| 16777472 | 16778239 | CN           | China         |

---

## 🛠️ Step-by-Step Setup

### 1. Prepare Your IP List

In your Excel file, list IP addresses (e.g. `8.8.8.8`) in column A.

### 2. Import the IP2Location CSV into Excel

- Open the CSV file in Excel.
- Format it as a table (`Ctrl + T`)(optional).
- Make sure the columns are in this order: `ip_from`, `ip_to`, `country_code`, `country_name`.
- Name the table (optional), e.g., `IPRanges`.

### 3. Add VBA Functions

Press `Alt + F11` → Insert → Module → Paste the following code:

#### Function to Convert IP to Number

```vba
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
    IPToLong = bytes(0) * 16777216# + bytes(1) * 65536# + bytes(2) * 256# + bytes(3)
End Function
```

#### Function to Find Country from IP Range

```vba
Function FindCountry(ipNumber As Double, lookupRange As Range) As String
    Dim r As Range
    For Each r In lookupRange.Rows
        If ipNumber >= r.Cells(1, 1).Value And ipNumber <= r.Cells(1, 2).Value Then
            FindCountry = r.Cells(1, 4).Value ' Change to (1,3) to return country_code
            Exit Function
        End If
    Next r
    FindCountry = "Not found"
End Function
```

### 4. Use the Functions in Excel

Assuming:

- Your IP address is in cell `A2`
- Your IP2Location data is in `Sheet2!A2:D100000`

Use this formula:

```excel
=FindCountry(IPToLong(A2), Sheet2!A2:D100000)
```

---

## 📎 Notes

- If you want to return the `country_code`, change `r.Cells(1, 4)` to `r.Cells(1, 3)` in the VBA function.
- Excel performance may slow down with very large CSV files.
- Consider using Power Query or a proper database for large-scale processing.

---

## 🙌 Credits

- IP range data: [IP2Location](https://www.ip2location.com/)
- Script & documentation: Muhamad Yusuf
