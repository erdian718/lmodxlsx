# lmodxlsx

[plandem/xlsx](https://github.com/plandem/xlsx) bindings for [Lua](https://github.com/ofunc/lua).

## Usage

```go
package main

import (
	"ofunc/lmodxlsx"
	"ofunc/lua/util"
)

func main() {
	l := util.NewState()
	l.Preload("xlsx", lmodxlsx.Open)
	util.Run(l, "main.lua")
}
```

```lua
local xlsx = require 'xlsx'

local xl, err = xlsx.open('test.xlsx')
assert(err == nil, err)

print(xl:sheet(1):cell('A1').value)

xl:close()
```

## Dependencies

* [ofunc/lua](https://github.com/ofunc/lua)
* [plandem/xlsx](https://github.com/plandem/xlsx)

## Documentation

### xlsx.open(x)

Opens a XLSX file with name or io.reader.

### xlsx.fromtime(x)

Converts Lua time value to XLSX number value.

### xlsx.totime(x)

Converts XLSX number value to Lua time value.

### spreadsheet:sheet(x)

Returns a sheet by index or name.

### spreadsheet:sheets()

Returns iterator for all sheets of spreadsheet.

### spreadsheet:save([name])

Saves current XLSX file with `name`.

### spreadsheet:close()

Closes current XLSX file.

### sheet:range(ref)

Returns a range for `ref`.

### sheet:cell(c[, r])

Returns a cell for indexes or ref.

### sheet:col(i)

Returns a col for index.

### sheet:row(i)

Returns a row for index.

### sheet:cols()

Returns iterator for all cols of sheet.

### sheet:rows()

Returns iterator for all rows of sheet.

### sheet:inscol(i)

Inserts a col at index and returns it.

### sheet:insrow(i)

Inserts a row at index and returns it.

### sheet:delcol(i)

Deletes a col at index.

### sheet:delrow(i)

Deletes a row at index.

### sheet:dim()

Returns total number of cols and rows in sheet.

### sheet:setdim(ncol, nrow)

Sets total number of cols and rows in sheet.

### sheet:set(options)

Sets options for sheet.
```
active: boolean
visibility: 'visible' | 'hidden' | 'veryhidden'
```

### sheet.name

The name of the sheet.
It's readable and writable.

### range:bounds()

Returns bounds of range: `fromcol`, `fromrow`, `tocol`, `torow`.

### range:cells()

Returns iterator for all cells in range.

### range:clear()

Clears each cell value in range.

### range:reset()

Resets each cell data into zero state.

### range:copyto(cidx, ridx)

Copies range cells into another range starting indexes cidx and ridx.
Merged cells are not supported.

### range:merge()

Merges range.

### range:split()

Splits cells in range.

### range:format(s)

Sets style format to all cells in range.

### range:link(l)

Sets hyperlink for range.

### col:cell(ridx)

Returns cell of col at row with `ridx`.

### col:cells()

Returns iterator for all cells in col.

### col:clear()

Clears each cell value in col.

### col:reset()

Resets each cell data into zero state.

### col:copyto(cidx[, withoptions])

Copies col cells into another col with `cidx`.
Merged cells are not supported.

### col:set(options)

Sets options for column.
```
level: integer (1 - 8)
collapsed: boolean
phonetic: boolean
hidden: boolean
width: number
```

### col.format

Default style for the column.
It's readable and writable.

### col.index

Col index of the column.
It's read only.

### row:cell(cidx)

Returns cell of col at col with `cidx`.

### row:cells()

Returns iterator for all cells in row.

### row:clear()

Clears each cell value in row.

### row:reset()

Resets each cell data into zero state.

### row:copyto(ridx[, withoptions])

Copies col cells into another row with `ridx`.
Merged cells are not supported.

### row:set(options)

Sets options for row.
```
level: integer (1 - 8)
collapsed: boolean
phonetic: boolean
hidden: boolean
height: number
```

### row.format

Default style for the row.
It's readable and writable.

### row.index

Col index of the row.
It's read only.

### cell:clear()

Clears cell's value.

### cell:reset()

Resets current current cell information.

### cell.value

The value of the cell.
It's readable and writable.

### cell.format

Default style for the cell.
It's readable and writable.

### cell.link

The hyperlink for the cell.
It's readable and writable.
