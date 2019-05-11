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
)

func metaSpreadsheet(l *lua.State, msheet int) int {
	l.NewTable(0, 8)
	idx := l.AbsIndex(-1)

	l.Push("sheet")
	l.PushClosure(lSpreadsheetSheet, msheet)
	l.SetTableRaw(idx)

	l.Push("sheets")
	l.PushClosure(lSpreadsheetSheets, msheet)
	l.SetTableRaw(idx)

	l.Push("save")
	l.Push(lSpreadsheetSave)
	l.SetTableRaw(idx)

	l.Push("close")
	l.Push(lSpreadsheetClose)
	l.SetTableRaw(idx)

	l.Push("__index")
	l.PushIndex(idx)
	l.SetTableRaw(idx)

	return idx
}

func lSpreadsheetSheet(l *lua.State) int {
	spreadsheet := toSpreadsheet(l, 1)
	index := -1
	var sheet xlsx.Sheet
	if l.TypeOf(2) == lua.TypeNumber {
		index = int(l.ToInteger(2)) - 1
	} else {
		name := l.ToString(2)
		for i, v := range spreadsheet.GetSheetNames() {
			if v == name {
				index = i
			}
		}
	}
	if index >= 0 {
		if l.ToBoolean(3) {
			sheet = spreadsheet.Sheet(index, xlsx.SheetModeStream)
		} else {
			sheet = spreadsheet.Sheet(index)
		}
	}
	if sheet == nil {
		return 0
	} else {
		l.Push(sheet)
		l.PushIndex(lua.FirstUpVal - 1)
		l.SetMetaTable(-2)
		return 1
	}
}

func lSpreadsheetSheets(l *lua.State) int {
	iter := toSpreadsheet(l, 1).Sheets()
	l.PushClosure(func(l *lua.State) int {
		if !iter.HasNext() {
			return 0
		}
		i, s := iter.Next()
		l.Push(i + 1)
		l.Push(s)
		l.PushIndex(lua.FirstUpVal - 1)
		l.SetMetaTable(-2)
		return 2
	}, lua.FirstUpVal-1)
	return 1
}

func lSpreadsheetSave(l *lua.State) int {
	spreadsheet := toSpreadsheet(l, 1)
	if l.IsNil(2) {
		if err := spreadsheet.Save(); err == nil {
			l.Push(nil)
		} else {
			l.Push(err.Error())
		}
	} else {
		if err := spreadsheet.SaveAs(l.ToString(2)); err == nil {
			l.Push(nil)
		} else {
			l.Push(err.Error())
		}
	}
	return 1
}

func lSpreadsheetClose(l *lua.State) int {
	if err := toSpreadsheet(l, 1).Close(); err == nil {
		l.Push(nil)
	} else {
		l.Push(err.Error())
	}
	return 1
}
