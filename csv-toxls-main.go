package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"regexp"
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

func atof(in string) float64 {
	x, err := strconv.ParseFloat(in, 64)
	if err != nil {
		return 0.0
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
	reader.Comma = delimiter

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
			if cnt == 1 && !dataonly.MatchString(fields[j]) { //header recognition
				cell.Value = fields[j]
				continue
			}
			if inarr(j, nidx) {
				if strings.Contains(fields[j], ".") {
					cell.SetFloat(atof(fields[j]))
				} else {
					cell.SetInt(atoi(fields[j]))
				}
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

	cnt, skipped := 0, 0

	csvarr := strings.Split(string(csvbin), "\n")
	for i := 0; i < len(csvarr); i++ {
		if !empty && (len(csvarr[i]) < 1) {
			skipped++
			continue
		}
		xli := strings.TrimSpace(csvarr[i])
		if win {
			xli = toutf(xli)
		}
		fields := strings.Split(xli, string(delimiter))
		row := sheet.AddRow()
		for j := 0; j < len(fields); j++ {
			cell := row.AddCell()
			if i == 0 && !dataonly.MatchString(fields[j]) { //header recognition
				cell.Value = fields[j]
				continue
			}
			if inarr(j, nidx) {
				if strings.Contains(fields[j], ".") {
					cell.SetFloat(atof(fields[j]))
				} else {
					cell.SetInt(atoi(fields[j]))
				}
			} else {
				cell.Value = fields[j]
			}
		}
		cnt++
	}
	xlsxFile.Save(xlsname)
	fmt.Printf("Written: %d rows, skipped: %d\n", cnt, skipped)
}

func printHeader(csvname string) {
	var reader *csv.Reader

	file, err := os.Open(csvname)
	if err != nil {
		errprintf("ERROR: File '%s' not found\n", csvname)
	}
	defer file.Close()

	if win {
		reader = csv.NewReader(charmap.Windows1251.NewDecoder().Reader(file))
	} else {
		reader = csv.NewReader(file)
	}
	reader.Comma = delimiter
	reader.LazyQuotes = true

	record, e := reader.Read()
	if e != nil {
		errprintf("ERR: %s\n", e)
	}
	if len(record) < 2 {
		errprintf("ERROR: No header\n")
	}
	fmt.Printf("--------\n")
	for i := 0; i < len(record); i++ {
		fmt.Printf("%d: %s\n", i, record[i])
	}
	fmt.Printf("--------\n")
}

func about() {
	t := `usage:
	csv-toxls [flags] filename
	flags:
	 -w : windows-1251 encoding
	 -q : quoted fields
	 -c : comma delimited
	 -e : enable empty rows
	 -n : numeric fields (ex: -n:1:3:6)
	 -h : show header only
	`
	errprintf(t)
}

var (
	win       bool      //win-1251 encoded
	quoted    bool      //quoted fields
	empty     bool      // enable empty rows
	header    bool      //show header
	nidx      = []int{} //numeric fields
	dataonly  = regexp.MustCompile(`^\d`)
	delimiter = ';' //default: semicolon (;) instead comma (,)
)

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
			delimiter = ','
		case "e":
			empty = true
		case "h":
			header = true
		case "n":
			for _, n := range narr {
				if x := atoi(n); x != 0 {
					nidx = append(nidx, x)
				}
			}
		default:
			errprintf("ERROR: Invalid option: '%s'\n", f)
		}
	}

	if len(args) < 1 {
		errprintf("ERROR: No input file\n")
	}
	csvname := args[0]
	if filepath.Ext(csvname) != ".csv" {
		errprintf("ERROR: Unknown format (no .csv)\n")
	}
	xlsname := strings.Replace(csvname, ".csv", ".xlsx", -1)

	switch {
	case header:
		printHeader(csvname)
	case quoted:
		mkQuotedXlsx(csvname, xlsname, nidx)
	default:
		mkDefaultXlsx(csvname, xlsname, nidx)
	}

}
