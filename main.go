package main

import (
	"errors"
	"os"
	"strconv"
	"time"

	"github.com/nguyenthenguyen/docx"
	"github.com/okonma-violet/confdecoder"
	"github.com/okonma-violet/logs/logger"
	"github.com/okonma-violet/spec/logs/encode"
	"github.com/tealeg/xlsx"
)

// DOCX FILE TAG EXAMPLE: [[_tag1_]]
// TEMPLATE CONFIG's ROW EXAMPLE: tag1 B13
// TEMPLATE CONFIG's SHEET SPECIFYING EXAMPLE (ZERO BASED): sheet 0
// IF NO SHEET SPECIFIED - IT TAKES FIRST (0)

const (
	tagopen  = "[[_"
	tagclose = "_]]"
)

type replaceData struct {
	tag string

	value string
}

func main() {
	args := os.Args
	if len(args) < 5 || args[1] == "--help" {
		println("format: [template config file path] [docx template file path] [new docx file path] [xlsx file path]")
		return
	}

	tfd, err := confdecoder.ParseFile(args[1])
	if err != nil {
		println("Parsing template config file err: " + err.Error())
		return
	}
	df, err := docx.ReadDocxFile(args[2])
	if err != nil {
		println("Reading docx file err: " + err.Error())
		return
	}
	defer df.Close()

	xf, err := xlsx.OpenFile(args[4])
	if err != nil {
		println("Reading xlsx file err: " + err.Error())
		return
	}

	dx := df.Editable()

	flsh := logger.NewFlusher(encode.DebugLevel)
	l := flsh.NewLogsContainer("docscoitus")

	tmpl_rows := tfd.Rows()
	var sht *xlsx.Sheet
	var replaces []replaceData

	for i := 0; i < len(tmpl_rows); i++ {
		if tmpl_rows[i].Name == "sheet" {
			shti, err := strconv.Atoi(tmpl_rows[i].Value)
			if err != nil {
				l.Error("Template", errors.New("sheet specified, but not by num, atoi err: "+err.Error()))
				goto divorce
			}
			if len(xf.Sheets) < (shti + 1) {
				l.Error("Template", errors.New("specified sheet num is bigger than num of sheets in readed xlsx: "+strconv.Itoa(shti)+" you want (zero based), "+strconv.Itoa(len(xf.Sheets))+" xlsx has"))
				goto divorce
			}
			sht = xf.Sheets[shti]
			tmpl_rows = tmpl_rows[:i+copy(tmpl_rows[i:], tmpl_rows[i+1:])]
		}
	}

	l.Debug("Template", "rows readed: "+strconv.Itoa(len(tmpl_rows)))
	replaces = make([]replaceData, len(tmpl_rows))
	for i := 0; i < len(tmpl_rows); i++ {
		if tmpl_rows[i].Name == "" || tmpl_rows[i].Value == "" {
			l.Error("Template", errors.New("bad row - name or value is empty"))
			goto divorce
		}
		icol, irow, err := xlsx.GetCoordsFromCellIDString(tmpl_rows[i].Value)
		if err != nil {
			l.Error("Template", errors.New("getting coords from row value err: "+err.Error()))
			goto divorce
		}
		replaces[i].tag = tagopen + tmpl_rows[i].Name + tagclose
		l.Debug("Template", "readed row: tag="+replaces[i].tag+", coords x="+strconv.Itoa(icol)+" y="+strconv.Itoa(irow)+" ("+tmpl_rows[i].Value+")")

		if irow >= sht.MaxRow {
			l.Error("Template", errors.New("specified row num in coords is bigger than num of rows in sheet, "+strconv.Itoa(irow)+" you want (zero based), "+strconv.Itoa(sht.MaxRow)+" sheet has"))
			goto divorce
		}
		if icol >= sht.MaxCol {
			l.Error("Template", errors.New("specified column num in coords is bigger than num of columns in sheet, "+strconv.Itoa(icol)+" you want (zero based), "+strconv.Itoa(sht.MaxCol)+" sheet has"))
			goto divorce
		}
		cell, err := sht.Cell(irow, icol)
		if err != nil {
			l.Error("Template", errors.New("sheet.Cell() err: "+err.Error()))
			goto divorce
		}
		replaces[i].value = cell.String()
		l.Debug("Xlsx", "got value for tag "+replaces[i].tag+": "+cell.Value)
	}
	l.Debug("Template", "all rows are readed, all data from xlsx is gotten")

	for i := 0; i < len(replaces); i++ {
		if err = dx.Replace(replaces[i].tag, replaces[i].value, -1); err != nil {
			l.Error("Replace", errors.New("replacing tag \""+replaces[i].tag+"\" with \""+replaces[i].value+"\" err: "+err.Error()))
			goto divorce
		}
		l.Debug("Replace", "replaced tag \""+replaces[i].tag+"\" with \""+replaces[i].value+"\"")
	}
	l.Debug("Replace", "all done")

	if err = dx.WriteToFile(args[3]); err != nil {
		l.Error("Writing result file", err)
	}
	l.Debug("Writing result file", "result written to "+args[3])

divorce:
	flsh.Close()
	flsh.DoneWithTimeout(time.Second * 5)
}
