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

// plandem/xlsx bindings for Lua.
package lmodxlsx

import (
	"io"
	"time"

	"ofunc/lua"

	"github.com/plandem/xlsx"
)

// Open opens the module.
func Open(l *lua.State) int {
	mcell := metaCell(l)
	mrow := metaRow(l, mcell)
	mcol := metaCol(l, mcell)
	mrange := metaRange(l, mcell)
	msheet := metaSheet(l, mcell, mrow, mcol, mrange)
	mspreadsheet := metaSpreadsheet(l, msheet)
	l.NewTable(0, 4)

	l.Push("version")
	l.Push("0.0.1")
	l.SetTableRaw(-3)

	l.Push("open")
	l.PushClosure(lOpen, mspreadsheet)
	l.SetTableRaw(-3)

	l.Push("fromtime")
	l.Push(lFromTime)
	l.SetTableRaw(-3)

	l.Push("totime")
	l.Push(lToTime)
	l.SetTableRaw(-3)

	return 1
}

func lOpen(l *lua.State) int {
	var arg interface{}
	if l.TypeOf(1) == lua.TypeUserData {
		if r, ok := l.GetRaw(1).(io.Reader); ok {
			arg = r
		}
	} else {
		arg = l.ToString(1)
	}

	xl, err := xlsx.Open(arg)
	if err == nil {
		l.Push(xl)
		l.PushIndex(lua.FirstUpVal - 1)
		l.SetMetaTable(-2)
		return 1
	} else {
		l.Push(nil)
		l.Push(err.Error())
		return 2
	}
}

func lFromTime(l *lua.State) int {
	v := l.GetRaw(1).(time.Time).Unix()
	l.Push(float64(v)/86400 + 25569)
	return 1
}

func lToTime(l *lua.State) int {
	if v, err := l.TryFloat(1); err == nil {
		l.Push(time.Unix(int64(86400*(v-25569)), 0))
	} else {
		t, _ := time.Parse("2006-01-02T15:04:05", l.ToString(1))
		l.Push(t)
	}
	return 1
}
