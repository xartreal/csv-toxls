package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
	"golang.org/x/text/encoding/charmap"
)

func parseCL() ([]string, []string) {
	var args, flags = []string{}, []string{}
	for i := 1; i < len(os.Args); i++ {
		if strings.HasPrefix(os.Args[i], "-") { //flags
			flags = append(flags, strings.TrimPrefix(os.Args[i], "-"))
		} else {
			args = append(args, os.Args[i])
		}
	}
	return args, flags
}

func parseMultiFlg(flag string) (string, []string) { //flag, string values array
	if !strings.Contains(flag, ":") {
		return flag, []string{}
	}
	rf := strings.Split(flag, ":")
	return rf[0], rf[1:]
}

func errprintf(format string, a ...interface{}) {
	fmt.Printf(format, a...)
	os.Exit(1)
}

func toutf(text string) string {
	ftext, _ := charmap.Windows1251.NewDecoder().String(text)
	return ftext
}

func atoi(in string) int {
	x, err := strconv.Atoi(in)
	if err != nil {
		return 0
	}
	return x
}

func inarr(n int, nidx []int) bool {
	for _, v := range nidx {
		if v == n {
			return true
		}
	}
	return false
}

func mkQuotedXlsx(csvname, xlsname string, nidx []int) {
	var reader *csv.Reader
	csvFile, err := os.Open(csvname)
	if err != nil {
		errprintf("ERROR: Can't open file '%s'\n", csvname)
	}
	defer csvFile.Close()

	if win {
		reader = csv.NewReader(charmap.Windows1251.NewDecoder().Reader(csvFile))
	} else {
		reader = csv.NewReader(csvFile)
	}
	if !comma {
		reader.Comma = ';'
	}

	xlsxFile := xlsx.NewFile()
	sheet, err := xlsxFile.AddSheet("1")
	if err != nil {
		errprintf("ERROR: Can't create file '%s'\n", xlsname)
	}
	fields, err := reader.Read()
	if err != nil {
		fmt.Printf("ERR: %q\n", err)
	}
	cnt := 1
	for err == nil {
		row := sheet.AddRow()
		for j := 0; j < len(fields); j++ {
			cell := row.AddCell()
			if cnt == 1 && !dataonly {
				cell.Value = fields[j]
				continue
			}
			if inarr(j, nidx) {
				cell.SetInt(atoi(fields[j]))
			} else {
				cell.Value = fields[j]
			}
		}
		fields, err = reader.Read()
		if err == nil {
			cnt++
		} else if err != io.EOF {
			fmt.Printf("ERR: %q\n", err)
		}
	}
	xlsxFile.Save(xlsname)
	fmt.Printf("Written: %d rows\n", cnt)
}

func mkDefaultXlsx(csvname, xlsname string, nidx []int) {
	csvbin, err := ioutil.ReadFile(csvname)
	if err != nil {
		errprintf("ERROR: Can't open file '%s'\n", csvname)
	}

	xlsxFile := xlsx.NewFile()
	sheet, _ := xlsxFile.AddSheet("1")

	cnt, skip := 0, 0
	delimiter := ";"
	if comma {
		delimiter = ","
	}

	csvarr := strings.Split(string(csvbin), "\n")
	for i := 0; i < len(csvarr); i++ {
		if !empty && (len(csvarr[i]) < 1) {
			skip++
			continue
		}
		xli := strings.TrimSpace(csvarr[i])
		if win {
			xli = toutf(xli)
		}
		cline := strings.Split(xli, delimiter)
		row := sheet.AddRow()
		for j := 0; j < len(cline); j++ {
			cell := row.AddCell()
			if i == 0 && !dataonly {
				cell.Value = cline[j]
				continue
			}
			if inarr(j, nidx) {
				cell.SetInt(atoi(cline[j]))
			} else {
				cell.Value = cline[j]
			}
		}
		cnt++
	}
	xlsxFile.Save(xlsname)
	fmt.Printf("Written: %d rows, skipped: %d\n", cnt, skip)
}

func about() {
	t := `usage:
	csv-toxls [flags] filename
	flags:
	 -w : windows encoding
	 -q : quoted fields
	 -c : comma delimited
	 -e : enable empty rows
	 -n : numeric fields (ex: -n:1:3:6)
	 -d : data only (no header)
	`
	errprintf(t)
}

var win bool       //win-1251 encoded
var quoted bool    //quoted fields
var comma bool     // comma (,) instead semicolon (;)
var empty bool     // enable empty rows
var nidx = []int{} //numeric fields
var dataonly bool  //data only

func main() {
	if len(os.Args) < 2 {
		about()
	}
	args, flags := parseCL()
	for _, v := range flags {
		f, narr := parseMultiFlg(v)
		switch f {
		case "w":
			win = true
		case "q":
			quoted = true
		case "c":
			comma = true
		case "e":
			empty = true
		case "d":
			dataonly = true
		case "n":
			for _, n := range narr {
				if x := atoi(n); x != 0 {
					nidx = append(nidx, x)
				}
			}

		}
	}

	if len(args) < 1 {
		errprintf("Incorrect command line syntax\n")
	}
	csvname := args[0]
	if filepath.Ext(csvname) != ".csv" {
		errprintf("ERROR: Unknown format (no .csv)\n")
	}
	xlsname := strings.Replace(csvname, ".csv", ".xlsx", -1)

	if quoted {
		mkQuotedXlsx(csvname, xlsname, nidx)
	} else {
		mkDefaultXlsx(csvname, xlsname, nidx)
	}

}
