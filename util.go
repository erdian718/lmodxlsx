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
	"github.com/plandem/xlsx/format"
)

func toSpreadsheet(l *lua.State, i int) *xlsx.Spreadsheet {
	if v, ok := l.GetRaw(1).(*xlsx.Spreadsheet); ok {
		return v
	} else {
		panic("xlsx: invalid spreadsheet: " + l.ToString(i))
	}
}

func toSheet(l *lua.State, i int) xlsx.Sheet {
	if v, ok := l.GetRaw(i).(xlsx.Sheet); ok {
		return v
	} else {
		panic("xlsx: invalid sheet: " + l.ToString(i))
	}
}

func toRange(l *lua.State, i int) *xlsx.Range {
	if v, ok := l.GetRaw(i).(*xlsx.Range); ok {
		return v
	} else {
		panic("xlsx: invalid range: " + l.ToString(i))
	}
}

func toCol(l *lua.State, i int) *xlsx.Col {
	if v, ok := l.GetRaw(i).(*xlsx.Col); ok {
		return v
	} else {
		panic("xlsx: invalid col: " + l.ToString(i))
	}
}

func toRow(l *lua.State, i int) *xlsx.Row {
	if v, ok := l.GetRaw(i).(*xlsx.Row); ok {
		return v
	} else {
		panic("xlsx: invalid row: " + l.ToString(i))
	}
}

func toCell(l *lua.State, i int) *xlsx.Cell {
	if v, ok := l.GetRaw(i).(*xlsx.Cell); ok {
		return v
	} else {
		panic("xlsx: invalid cell: " + l.ToString(i))
	}
}

func toStyle(l *lua.State, i int) format.DirectStyleID {
	if v, ok := l.GetRaw(i).(format.DirectStyleID); ok {
		return v
	} else {
		panic("xlsx: invalid style: " + l.ToString(i))
	}
}
