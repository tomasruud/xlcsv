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
	fmt.Fprintln(flag.CommandLine.Output(), `xlcsv - convert Excel to csv

USAGE:
  xlcsv [ARGS]

ARGS:
  <file name>
        Excel file to read, if "-" or not provided stdin will be used.

OPTIONS:`)
	flag.PrintDefaults()
}

func main() {
	output := flag.String("o", "", "Output file, if empty it will output to stdout.")

	var opts options
	flag.StringVar(&opts.sheet, "sheet", "Sheet1", "Sheet name.")
	flag.StringVar(&opts.password, "password", "", "File password, if any.")
	flag.Var(&opts.columns, "pick", "Comma separated list of column indexes to include (zero based). Can be used to reorder columns.")

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
	columns  columns
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

	if len(opts.columns) == 0 {
		w.WriteAll(rows)
	} else {
		for _, row := range rows {
			var rec []string
			for _, col := range opts.columns {
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
