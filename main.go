package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"os"

	"github.com/xuri/excelize/v2"
)

func usage() {
	fmt.Fprintln(flag.CommandLine.Output(), `xlcsv - convert Excel to csv

USAGE:
  xlcsv [ARGS]

ARGS:
  <file name>
        Excel file to read, if "-" or not provided stdin will be used

OPTIONS:`)
	flag.PrintDefaults()
}

func main() {
	output := flag.String("o", "", "Output file, if empty it will output to stdout")

	var opts options
	flag.StringVar(&opts.sheet, "sheet", "Sheet1", "Sheet name")
	flag.StringVar(&opts.password, "pw", "", "Password")

	flag.Usage = usage
	flag.Parse()

	var in io.Reader
	if flag.Arg(0) == "" || flag.Arg(0) == "-" {
		in = os.Stdin
	} else {
		f, err := os.Open(flag.Arg(0))
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error: opening file: %v\n", err)
			os.Exit(1)
		}
		defer f.Close()

		in = f
	}

	var out io.Writer
	if *output == "" {
		out = os.Stdout
	} else {
		f, err := os.Create(*output)
		if err != nil {
			fmt.Fprintf(os.Stderr, "Error: creating output file: %v\n", err)
			os.Exit(1)
		}
		defer f.Close()

		out = f
	}

	if err := run(in, out, opts); err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
}

type options struct {
	sheet    string
	password string
}

func run(input io.Reader, out io.Writer, opts options) error {
	file, err := excelize.OpenReader(input, excelize.Options{
		Password: opts.password,
	})
	if err != nil {
		return fmt.Errorf("opening spreadsheet (is it password protected?): %v", err)
	}
	defer file.Close()

	rows, err := file.GetRows(opts.sheet)
	if err != nil {
		return fmt.Errorf("reading rows: %v", err)
	}

	w := csv.NewWriter(out)
	w.WriteAll(rows)

	if err := w.Error(); err != nil {
		return fmt.Errorf("writing csv: %v", err)
	}

	return nil
}
