# XLSXread.jl
Excel file reader for julia based on EzXML and ZipFile. It can read a single sheet or multple sheets at a time. Returns a dictionary with DataFrames with the sheet number as the key

## Usage
```
using XLSXread

df = readxlsx("sample.xlsx", 1)
```
