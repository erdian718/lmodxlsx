/*
Copyright 2019 by ofunc

This software is provided 'as-is', without any express or implied warranty. In
no event will the authors be held liable for any damages arising from the use of
this software.

Permission is granted to anyone to use this software for any purpose, including
commercial applications, and to alter it and redistribute it freely, subject to
the following restrictions:

1. The origin of this software must not be misrepresented; you must not claim
that you wrote the original software. If you use this software in a product, an
acknowledgment in the product documentation would be appreciated but is not
required.

2. Altered source versions must be plainly marked as such, and must not be
misrepresented as being the original software.

3. This notice may not be removed or altered from any source distribution.
*/

package lmodxlsx

import (
	"ofunc/lua"

	"github.com/plandem/xlsx"
	"github.com/plandem/xlsx/options"
	"github.com/plandem/xlsx/types"
)

func metaSheet(l *lua.State, mcell, mrow, mcol, mrange int) int {
	l.NewTable(0, 16)
	idx := l.AbsIndex(-1)

	l.Push("range")
	l.PushClosure(lSheetRange, mrange)
	l.SetTableRaw(idx)

	l.Push("cell")
	l.PushClosure(lSheetCell, mcell)
	l.SetTableRaw(idx)

	l.Push("col")
	l.PushClosure(lSheetCol, mcol)
	l.SetTableRaw(idx)

	l.Push("row")
	l.PushClosure(lSheetRow, mrow)
	l.SetTableRaw(idx)

	l.Push("cols")
	l.PushClosure(lSheetCols, mcol)
	l.SetTableRaw(idx)

	l.Push("rows")
	l.PushClosure(lSheetRows, mrow)
	l.SetTableRaw(idx)

	l.Push("inscol")
	l.PushClosure(lSheetInsCol, mcol)
	l.SetTableRaw(idx)

	l.Push("insrow")
	l.PushClosure(lSheetInsRow, mrow)
	l.SetTableRaw(idx)

	l.Push("delcol")
	l.Push(lSheetDelCol)
	l.SetTableRaw(idx)

	l.Push("delrow")
	l.Push(lSheetDelRow)
	l.SetTableRaw(idx)

	l.Push("dim")
	l.Push(lSheetDim)
	l.SetTableRaw(idx)

	l.Push("setdim")
	l.Push(lSheetSetDim)
	l.SetTableRaw(idx)

	l.Push("set")
	l.Push(lSheetSet)
	l.SetTableRaw(idx)

	l.Push("close")
	l.Push(lSheetClose)
	l.SetTableRaw(idx)

	l.Push("__index")
	l.PushClosure(lSheetIndex, idx)
	l.SetTableRaw(idx)

	l.Push("__newindex")
	l.Push(lSheetNewIndex)
	l.SetTableRaw(idx)

	return idx
}

func lSheetRange(l *lua.State) int {
	l.Push(toSheet(l, 1).Range(types.Ref(l.ToString(2))))
	l.PushIndex(lua.FirstUpVal - 1)
	l.SetMetaTable(-2)
	return 1
}

func lSheetCell(l *lua.State) int {
	sheet := toSheet(l, 1)
	var cell *xlsx.Cell
	if l.AbsIndex(-1) > 2 {
		c := int(l.ToInteger(2)) - 1
		r := int(l.ToInteger(3)) - 1
		cell = sheet.Cell(c, r)
	} else {
		cell = sheet.CellByRef(types.CellRef(l.ToString(2)))
	}
	l.Push(cell)
	l.PushIndex(lua.FirstUpVal - 1)
	l.SetMetaTable(-2)
	return 1
}

func lSheetCol(l *lua.State) int {
	i := int(l.ToInteger(2)) - 1
	l.Push(toSheet(l, 1).Col(i))
	l.PushIndex(lua.FirstUpVal - 1)
	l.SetMetaTable(-2)
	return 1
}

func lSheetRow(l *lua.State) int {
	i := int(l.ToInteger(2)) - 1
	l.Push(toSheet(l, 1).Row(i))
	l.PushIndex(lua.FirstUpVal - 1)
	l.SetMetaTable(-2)
	return 1
}

func lSheetCols(l *lua.State) int {
	iter := toSheet(l, 1).Cols()
	l.PushClosure(func(l *lua.State) int {
		if !iter.HasNext() {
			return 0
		}
		i, c := iter.Next()
		l.Push(i + 1)
		l.Push(c)
		l.PushIndex(lua.FirstUpVal - 1)
		l.SetMetaTable(-2)
		return 2
	}, lua.FirstUpVal-1)
	return 1
}

func lSheetRows(l *lua.State) int {
	iter := toSheet(l, 1).Rows()
	l.PushClosure(func(l *lua.State) int {
		if !iter.HasNext() {
			return 0
		}
		i, r := iter.Next()
		l.Push(i + 1)
		l.Push(r)
		l.PushIndex(lua.FirstUpVal - 1)
		l.SetMetaTable(-2)
		return 2
	}, lua.FirstUpVal-1)
	return 1
}

func lSheetInsCol(l *lua.State) int {
	i := int(l.ToInteger(2)) - 1
	l.Push(toSheet(l, 1).InsertCol(i))
	l.PushIndex(lua.FirstUpVal - 1)
	l.SetMetaTable(-2)
	return 1
}

func lSheetInsRow(l *lua.State) int {
	i := int(l.ToInteger(2)) - 1
	l.Push(toSheet(l, 1).InsertRow(i))
	l.PushIndex(lua.FirstUpVal - 1)
	l.SetMetaTable(-2)
	return 1
}

func lSheetDelCol(l *lua.State) int {
	i := int(l.ToInteger(2)) - 1
	toSheet(l, 1).DeleteCol(i)
	return 0
}

func lSheetDelRow(l *lua.State) int {
	i := int(l.ToInteger(2)) - 1
	toSheet(l, 1).DeleteRow(i)
	return 0
}

func lSheetDim(l *lua.State) int {
	c, r := toSheet(l, 1).Dimension()
	l.Push(c)
	l.Push(r)
	return 2
}

func lSheetSetDim(l *lua.State) int {
	c := int(l.ToInteger(2)) - 1
	r := int(l.ToInteger(3)) - 1
	toSheet(l, 1).SetDimension(c, r)
	return 0
}

func lSheetSet(l *lua.State) int {
	sheet := toSheet(l, 1)

	l.Push("active")
	l.GetTable(2)
	if l.ToBoolean(-1) {
		sheet.SetActive()
	}

	l.Push("visibility")
	l.GetTable(2)
	switch l.ToString(-1) {
	case "visible":
		sheet.Set(&options.SheetOptions{
			Visibility: options.VisibilityTypeVisible,
		})
	case "hidden":
		sheet.Set(&options.SheetOptions{
			Visibility: options.VisibilityTypeHidden,
		})
	case "veryhidden":
		sheet.Set(&options.SheetOptions{
			Visibility: options.VisibilityTypeVeryHidden,
		})
	}
	return 0
}

func lSheetClose(l *lua.State) int {
	toSheet(l, 1).Close()
	return 0
}

func lSheetIndex(l *lua.State) int {
	sheet := toSheet(l, 1)
	switch key := l.ToString(2); key {
	case "name":
		l.Push(sheet.Name())
	default:
		l.Push(key)
		l.GetTableRaw(lua.FirstUpVal - 1)
	}
	return 1
}

func lSheetNewIndex(l *lua.State) int {
	sheet := toSheet(l, 1)
	switch key := l.ToString(2); key {
	case "name":
		sheet.SetName(l.ToString(3))
	default:
		panic("xlsx.sheet: invalid field: " + key)
	}
	return 0
}
