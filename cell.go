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

	"github.com/plandem/xlsx/types"
)

func metaCell(l *lua.State) int {
	l.NewTable(0, 4)
	idx := l.AbsIndex(-1)

	l.Push("clear")
	l.Push(lCellClear)
	l.SetTableRaw(idx)

	l.Push("reset")
	l.Push(lCellReset)
	l.SetTableRaw(idx)

	l.Push("__index")
	l.PushClosure(lCellIndex, idx)
	l.SetTableRaw(idx)

	l.Push("__newindex")
	l.Push(lCellNewIndex)
	l.SetTableRaw(idx)

	return idx
}

func lCellClear(l *lua.State) int {
	toCell(l, 1).Clear()
	return 0
}

func lCellReset(l *lua.State) int {
	toCell(l, 1).Reset()
	return 0
}

func lCellIndex(l *lua.State) int {
	cell := toCell(l, 1)
	switch key := l.ToString(2); key {
	case "value":
		switch cell.Type() {
		case types.CellTypeNumber:
			v, _ := cell.Float()
			l.Push(v)
		case types.CellTypeBool:
			v, _ := cell.Bool()
			l.Push(v)
		case types.CellTypeSharedString, types.CellTypeInlineString:
			l.Push(cell.String())
		default:
			if v, err := cell.Float(); err == nil {
				l.Push(v)
			} else {
				l.Push(cell.String())
			}
		}
	case "format":
		l.Push(cell.Formatting())
	case "link":
		if link := cell.Hyperlink(); link == nil {
			l.Push(nil)
		} else {
			l.Push(link.String())
		}
	default:
		l.Push(key)
		l.GetTableRaw(lua.FirstUpVal - 1)
	}
	return 1
}

func lCellNewIndex(l *lua.State) int {
	cell := toCell(l, 1)
	switch key := l.ToString(2); key {
	case "value":
		switch l.TypeOf(3) {
		case lua.TypeNumber:
			v, _ := l.TryFloat(3)
			cell.SetFloat(v)
		case lua.TypeBoolean:
			cell.SetBool(l.ToBoolean(3))
		case lua.TypeNil:
			cell.Clear()
		default:
			cell.SetString(l.ToString(3))
		}
	case "format":
		cell.SetFormatting(toStyle(l, 3))
	case "link":
		if l.IsNil(3) {
			cell.RemoveHyperlink()
		} else {
			cell.SetHyperlink(l.ToString(3))
		}
	default:
		panic("xlsx.cell: invalid field: " + key)
	}
	return 0
}
