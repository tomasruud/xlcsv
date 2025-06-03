package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

func usage() {
	fmt.Fprint(os.Stderr, `from-xlsx - parse CSV from a Excel sheet

USAGE:
  from-xlsx [OPTIONS] [ARGS]

ARGS:
  <file name>
        Excel file to read, if "-" or not provided stdin will be used.

OPTIONS:
  -h, --help
      Prints this help information

  -p, --password <value>
      Password, if the file is password protected

  -ls, --list-sheets
      Lists available sheets

  -d, --delimiter <value>
      CSV delimiter, defaults to ",", can be set using the CSV_DELIMITER env var

  -s, --sheet <value>
      Sheet name that should be converted, defaults to "Sheet1"
`)
}

func main() {
	opts := struct {
		password string

		listSheets bool

		delimiter string
		sheet     string
	}{}

	flag.StringVar(&opts.password, "p", "", "")
	flag.StringVar(&opts.password, "password", "", "")

	flag.BoolVar(&opts.listSheets, "ls", false, "")
	flag.BoolVar(&opts.listSheets, "list-sheets", false, "")

	flag.StringVar(&opts.delimiter, "d", os.Getenv("CSV_DELIMITER"), "")
	flag.StringVar(&opts.delimiter, "delimiter", os.Getenv("CSV_DELIMITER"), "")

	flag.StringVar(&opts.sheet, "s", "Sheet1", "")
	flag.StringVar(&opts.sheet, "sheet", "Sheet1", "")

	flag.Usage = usage
	flag.Parse()

	if len(opts.delimiter) != 1 {
		if d := os.Getenv("CSV_DELIMITER"); len(d) == 1 {
			opts.delimiter = d
		} else {
			if opts.delimiter != "" {
				fmt.Fprintln(os.Stderr, "Invalid CSV delimiter provided, falling back to ','.")
			}
			opts.delimiter = ","
		}
	}

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

	if opts.listSheets {
		fmt.Println(strings.Join(file.GetSheetList(), "\n"))
	} else {
		rows, err := file.GetRows(opts.sheet)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error: reading rows: %v\n", err)
			os.Exit(1)
		}

		w := csv.NewWriter(os.Stdout)
		w.Comma = rune(opts.delimiter[0])

		if err := w.WriteAll(rows); err != nil {
			fmt.Fprintf(os.Stderr, "Error: writing csv: %v\n", err)
			os.Exit(1)
		}
	}
}
