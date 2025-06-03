package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func usage() {
	fmt.Fprintln(os.Stderr, `xlcsv - convert Excel to csv

USAGE:
  xlcsv [OPTIONS] [ARGS]

ARGS:
  <file name>
        Excel file to read, if "-" or not provided stdin will be used.

OPTIONS:
  -h, --help
      Prints help information

  -p, --password <password>
      Password, if the file is password protected

  -ls, --list-sheets
      Lists available sheets

  -s, --sheet <name>
      Sheet name that should be converted, defaults to "Sheet1"
  -c, --columns <x[,y]>
      Column indexes to include, zero based, can be used to change column order, defaults to all columns
  -o, --output <name>
      Output file, if empty stdout will be used`)
}

func main() {
	opts := struct {
		listSheets bool

		password string

		sheet   string
		columns columns
		output  string
	}{}

	flag.StringVar(&opts.password, "p", "", "")
	flag.StringVar(&opts.password, "password", "", "")

	flag.BoolVar(&opts.listSheets, "ls", false, "")
	flag.BoolVar(&opts.listSheets, "list-sheets", false, "")

	flag.StringVar(&opts.sheet, "s", "Sheet1", "")
	flag.StringVar(&opts.sheet, "sheet", "Sheet1", "")

	flag.Var(&opts.columns, "c", "")
	flag.Var(&opts.columns, "columns", "")

	flag.StringVar(&opts.output, "o", "", "")
	flag.StringVar(&opts.output, "output", "", "")

	flag.Usage = usage
	flag.Parse()

	input := flag.Arg(0)

	var in io.Reader
	if input == "" || input == "-" {
		in = os.Stdin
	} else {
		f, err := os.Open(input)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error: opening file: %v\n", err)
			os.Exit(1)
		}
		defer f.Close()

		in = f
	}

	file, err := excelize.OpenReader(in, excelize.Options{
		Password: opts.password,
	})
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: opening spreadsheet (is it password protected?): %v\n", err)
		os.Exit(1)
	}
	defer file.Close()

	var out io.Writer
	if opts.output == "" {
		out = os.Stdout
	} else {
		f, err := os.Create(opts.output)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error: creating output file: %v\n", err)
			os.Exit(1)
		}
		defer f.Close()

		out = f
	}

	if opts.listSheets {
		err = dumpSheets(file, os.Stdout)
	} else {
		err = dumpData(file, out, opts.sheet, opts.columns)
	}

	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
}

func dumpSheets(file *excelize.File, out io.Writer) error {
	_, err := out.Write([]byte(strings.Join(file.GetSheetList(), "\n") + "\n"))
	return err
}

func dumpData(file *excelize.File, out io.Writer, sheet string, columns []int) error {
	rows, err := file.GetRows(sheet)
	if err != nil {
		return fmt.Errorf("reading rows: %v", err)
	}

	w := csv.NewWriter(out)

	if len(columns) == 0 {
		w.WriteAll(rows)
	} else {
		for _, row := range rows {
			var rec []string
			for _, col := range columns {
				if col < len(row) {
					rec = append(rec, row[col])
				}
			}
			w.Write(rec)
		}
	}

	if err := w.Error(); err != nil {
		return fmt.Errorf("writing csv: %v", err)
	}

	return nil
}

// columns is a custom cli flag that contains a comma separated list of ints.
type columns []int

var _ flag.Value = (*columns)(nil)

func (c *columns) Set(v string) error {
	for a := range strings.SplitSeq(v, ",") {
		i, err := strconv.Atoi(a)
		if err != nil {
			return err
		}

		*c = append(*c, i)
	}
	return nil
}

func (c *columns) String() string {
	var buf strings.Builder
	for _, i := range *c {
		if buf.Len() > 0 {
			buf.WriteRune(',')
		}

		buf.WriteString(strconv.Itoa(i))
	}
	return buf.String()
}
