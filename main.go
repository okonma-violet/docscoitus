package main

import (
	"os"
	"strconv"

	"github.com/nguyenthenguyen/docx"
	"github.com/okonma-violet/confdecoder"
	"github.com/tealeg/xlsx"
)

// DOCX FILE TAG EXAMPLE: [[_tag1_]]
// TAG NAMING [[_sheet_]] IS FORBIDDEN
// TEMPLATE CONFIG's ROW EXAMPLE: tag1 B13
//
// TEMPLATE CONFIG's SHEET SPECIFYING EXAMPLE: "sheet 0 (ZERO BASED)" OR "sheet [sheetname]"
// SHEET MUST BE SPECIFIED BEFORE ROWS THIS SHEET IS APPLIED ON
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
	args := os.Args[1:]
	if len(args) < 4 || args[0] == "--help" {
		println("format: [docx template file path] [new docx file path] [template config file path] [xlsx file path] (...[template config file path] [xlsx file path])")
		return
	}

	df, err := docx.ReadDocxFile(args[0])
	if err != nil {
		println("Reading docx file err: " + err.Error())
		return
	}
	defer df.Close()

	if (len(args))%2 > 0 {
		println("Xlsx's args must be paired (template's config filepath + xlsx filepath) (at least one pair)")
		return
	}

	dx := df.Editable()
	for k := len(args) - 1; k > 1; k -= 2 {
		xf, err := xlsx.OpenFile(args[k])
		if err != nil {
			println("Reading xlsx file err: " + err.Error())
			return
		}
		if len(xf.Sheets) == 0 {
			println("Reading xlsx file err: no sheets found in file") // яхз возможно ли это
			return
		}

		tfd, err := confdecoder.ParseFile(args[k-1])
		if err != nil {
			println("Parsing template config file err: " + err.Error())
			return
		}

		sht := xf.Sheets[0]
		for j := 0; j < len(tfd.Rows); j++ {
			if tfd.Rows[j].Key == "sheet" {
				shti, err := strconv.Atoi(tfd.Rows[j].Value)
				if err != nil {
					var ok bool
					if sht, ok = xf.Sheet[tfd.Rows[j].Value]; !ok {
						println("Xlsx template's config err: sheet specified neither by num nor by existing sheetname (" + args[k-1] + ")")
						return
					}
				} else {
					if len(xf.Sheets) < (shti + 1) {
						println("Xlsx templates  config'err: specified sheet num is bigger than num of sheets in readed xlsx: " + strconv.Itoa(shti) + " you want (zero based), " + strconv.Itoa(len(xf.Sheets)) + " xlsx has")
						return
					}
					sht = xf.Sheets[shti]
				}
				continue
			}

			if tfd.Rows[j].Key == "" || tfd.Rows[j].Value == "" {
				println("Xlsx template's config err: bad row - name or value is empty")
				return
			}
			icol, irow, err := xlsx.GetCoordsFromCellIDString(tfd.Rows[j].Value)
			if err != nil {
				println("Xlsx template's config err: getting coords from row value err: " + err.Error())
				return
			}
			tag := tagopen + tfd.Rows[j].Key + tagclose
			if irow >= sht.MaxRow {
				println("Xlsx template's config err: specified row num in coords is bigger than num of rows in sheet, " + strconv.Itoa(irow) + " you want (zero based), " + strconv.Itoa(sht.MaxRow) + " sheet has")
				return
			}
			if icol >= sht.MaxCol {
				println("Xlsx template's config err: specified column num in coords is bigger than num of columns in sheet, " + strconv.Itoa(icol) + " you want (zero based), " + strconv.Itoa(sht.MaxCol) + " sheet has")
				return
			}
			cell, err := sht.Cell(irow, icol)
			if err != nil {
				println("Xlsx template's config err: sheet.Cell() err: " + err.Error())
				return
			}
			if err = dx.Replace(tag, cell.String(), -1); err != nil {
				println("Replace err: replacing tag \"" + tag + "\" with \"" + cell.String() + "\" err: " + err.Error())
				return
			}
			println("Replace: replaced tag \"" + tag + "\" with \"" + cell.String() + "\"")
		}

	}

	if err = dx.WriteToFile(args[1]); err != nil {
		println("Writing result file err: " + err.Error())
	}
	println("Done! Result written to " + args[1])
}
