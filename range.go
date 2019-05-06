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

func metaRange(l *lua.State, mcell int) int {
	l.NewTable(0, 16)
	idx := l.AbsIndex(-1)

	l.Push("bounds")
	l.Push(lRangeBounds)
	l.SetTableRaw(idx)

	l.Push("cells")
	l.PushClosure(lRangeCells, mcell)
	l.SetTableRaw(idx)

	l.Push("clear")
	l.Push(lRangeClear)
	l.SetTableRaw(idx)

	l.Push("reset")
	l.Push(lRangeReset)
	l.SetTableRaw(idx)

	l.Push("copyto")
	l.Push(lRangeCopyTo)
	l.SetTableRaw(idx)

	l.Push("merge")
	l.Push(lRangeMerge)
	l.SetTableRaw(idx)

	l.Push("split")
	l.Push(lRangeSplit)
	l.SetTableRaw(idx)

	l.Push("format")
	l.Push(lRangeFormat)
	l.SetTableRaw(idx)

	l.Push("link")
	l.Push(lRangeLink)
	l.SetTableRaw(idx)

	l.Push("__index")
	l.PushIndex(idx)
	l.SetTableRaw(idx)

	return idx
}

func lRangeBounds(l *lua.State) int {
	bs := toRange(l, 1).Bounds()
	l.Push(bs.FromCol)
	l.Push(bs.FromRow)
	l.Push(bs.ToCol)
	l.Push(bs.ToRow)
	return 4
}

func lRangeCells(l *lua.State) int {
	iter := toRange(l, 1).Cells()
	l.PushClosure(func(l *lua.State) int {
		if !iter.HasNext() {
			return 0
		}
		cidx, ridx, cell := iter.Next()
		l.Push(cidx + 1)
		l.Push(ridx + 1)
		l.Push(cell)
		l.PushIndex(lua.FirstUpVal - 1)
		l.SetMetaTable(-2)
		return 3
	}, lua.FirstUpVal-1)
	return 1
}

func lRangeClear(l *lua.State) int {
	toRange(l, 1).Clear()
	return 0
}

func lRangeReset(l *lua.State) int {
	toRange(l, 1).Reset()
	return 0
}

func lRangeCopyTo(l *lua.State) int {
	if l.AbsIndex(-1) == 2 {
		toRange(l, 1).CopyToRef(types.Ref(l.ToString(2)))
	} else {
		cidx := int(l.ToInteger(2))
		ridx := int(l.ToInteger(3))
		toRange(l, 1).CopyTo(cidx, ridx)
	}
	return 0
}

func lRangeMerge(l *lua.State) int {
	if err := toRange(l, 1).Merge(); err == nil {
		return 0
	} else {
		l.Push(err.Error())
		return 1
	}
}

func lRangeSplit(l *lua.State) int {
	toRange(l, 1).Split()
	return 0
}

func lRangeFormat(l *lua.State) int {
	toRange(l, 1).SetFormatting(toStyle(l, 2))
	return 0
}

func lRangeLink(l *lua.State) int {
	rng := toRange(l, 1)
	if l.IsNil(2) {
		rng.RemoveHyperlink()
	} else {
		rng.SetHyperlink(l.ToString(2))
	}
	return 0
}
