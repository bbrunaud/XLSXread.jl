# XLSXread.jl
Fast Excel file reader for julia based on [EzXML](https://github.com/bicycle1885/EzXML.jl) and [ZipFile](http://zipfilejl.readthedocs.io/en/latest/). It can read a single sheet or multple sheets at a time. Returns a dictionary with DataFrames with the sheet number as the key

## Usage
```julia
using XLSXread

df = readxlsx("sample.xlsx", 1)
```
```
4×3 DataFrames.DataFrame
│ Row │ Unit │ Process │ Value │
├─────┼──────┼─────────┼───────┤
│ 1   │ "U1" │ "P1"    │ 1.11  │
│ 2   │ "U2" │ "P2"    │ 2.22  │
│ 3   │ "U3" │ "P3"    │ 3.33  │
│ 4   │ "U4" │ "P4"    │ 4.44  │
```
