module XLSXread

using EzXML
using DataFrames
using ZipFile

export readxlsx

function readxlsx(file::String,sheetlist::Array{Int64,1})

  # Unzip file
  r = ZipFile.Reader(file)
  filearray = [f.name for f in r.files]

  # Get Strings
  filestr = "xl/sharedStrings.xml"
  fidx = indexin([filestr],filearray)[1]
  file = r.files[fidx]
  xfile = parsexml(readstring(file))

  sharedstring = String[]

  xroot = root(xfile)
  for a in eachnode(xroot)
    for b in eachnode(a)
      push!(sharedstring,content(b))
    end
  end

  # Get Values
  dfdict = Dict()
  for sheetnumber in sheetlist

    sheetvalues = Dict()
    filestr = "xl/worksheets/sheet$sheetnumber.xml"
    fidx = indexin([filestr],filearray)[1]
    file = r.files[fidx]
    xfile = parsexml(readstring(file))

    SD = nothing

    xroot = root(xfile)
    for sd in eachnode(xroot)
      if name(sd) == "sheetData"
        SD = sd
        for row in eachnode(sd)
          if name(row) == "row"
            for cell in eachnode(row)
              if name(cell) == "c"
                celltype = cell["t"]
                val = nothing
                if celltype == "n"
                  val = parse(Float64,content(cell))
                elseif celltype == "s"
                  val = sharedstring[parse(Int64,content(cell))+1]
                end
                sheetvalues[cell["r"]] = val
              end
            end
          end
        end
      end
    end

    ## TODO support more than 27 columns
    # Get range info
    abc = "ABCDEFGHIJKLMNOPQRSTUVXYZ"
    mincol = 1000
    maxcol = 0
    minrow = 1000000
    maxrow = 0
    for k in keys(sheetvalues)
      cidx = search(abc,k[1])
      mincol = min(mincol,cidx)
      maxcol = max(maxcol,cidx)
      ridx = parse(Int64,k[2:end])
      minrow = min(minrow,ridx)
      maxrow = max(maxrow,ridx)
    end

    # Build output
    # Assume the first row is the header
    df = DataFrame()
    header = String[]
    for c in mincol:maxcol
      hname = get(sheetvalues,"$(abc[c])$minrow","Header$c")
      push!(header,hname)
      column = []
      for r in (minrow+1):maxrow
        push!(column,get(sheetvalues,"$(abc[c])$r",NA))
      end
      df[Symbol(hname)] = @data([c for c in column])
    end
    dfdict[sheetnumber] = df
  end
  dfdict
end

function readxlsx(file::String,sheetnumber::Int64)
  d = readxlsx(file,[sheetnumber])
  return d[sheetnumber]
end

function readxlsx(file::String,sheetlist::UnitRange{Int64})
  return readxlsx(file,[i for i in sheetlist])
end

end
