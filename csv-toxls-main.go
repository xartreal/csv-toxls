package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
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

func errprintf(format string, a ...interface{}) {
	fmt.Printf(format, a...)
	os.Exit(1)
}

func toutf(text string) string {
	ftext, _ := charmap.Windows1251.NewDecoder().String(text)
	return ftext
}

func mkQuotedXlsx(csvname, xlsname string) {
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
		for _, field := range fields {
			cell := row.AddCell()
			cell.Value = field
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

func mkDefaultXlsx(csvname, xlsname string) {
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
			cell.Value = cline[j]
		}
		cnt++
	}
	xlsxFile.Save(xlsname)
	fmt.Printf("Written: %d rows, skipped: %d\n", cnt, skip)
}

var win bool    //win-1251 encoded
var quoted bool //quoted fields
var comma bool  // comma (,) instead semicolon (;)
var empty bool  // enable empty rows

func main() {
	if len(os.Args) < 2 {
		errprintf("usage:\ncsv-toxls [flags] filename\nflags:\n -w : windows encoding\n" +
			" -q : quoted fields\n -c : comma delimited\n -e : enable empty rows\n")
	}
	args, flags := parseCL()
	for _, v := range flags {
		switch v {
		case "w":
			win = true
		case "q":
			quoted = true
		case "c":
			comma = true
		case "e":
			empty = true
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
		mkQuotedXlsx(csvname, xlsname)
	} else {
		mkDefaultXlsx(csvname, xlsname)
	}

}
