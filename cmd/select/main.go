package main

import (
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"io"
	"os"
	"os/signal"
	"slices"
	"strconv"
	"strings"
	"syscall"
)

func usage() {
	fmt.Fprint(os.Stderr, `select - select columns from CSV

USAGE:
  select [OPTIONS] [ARGS]

ARGS:
  [<column>...]
      Column names

OPTIONS:
  -h, --help
      Prints this help information

  -s, --strict
      Toggle strict mode, the script will fail if rows are missing columns

  -t, --headers
      List header row

  -d, --delimiter <value>
      CSV delimiter, defaults to ",", can be set using the CSV_DELIMITER env var
`)
}

func main() {
	signal.Ignore(syscall.SIGPIPE)

	opts := struct {
		listHeaders bool
		strict      bool
		delimiter   string
	}{}

	flag.BoolVar(&opts.listHeaders, "t", false, "")
	flag.BoolVar(&opts.listHeaders, "headers", false, "")

	flag.BoolVar(&opts.strict, "s", false, "")
	flag.BoolVar(&opts.strict, "strict", false, "")

	flag.StringVar(&opts.delimiter, "d", "", "")
	flag.StringVar(&opts.delimiter, "delimiter", "", "")

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

	reader := csv.NewReader(os.Stdin)
	reader.Comma = rune(opts.delimiter[0])
	reader.ReuseRecord = true

	writer := csv.NewWriter(os.Stdout)
	writer.Comma = rune(opts.delimiter[0])

	var headers []string
	var indexes []int
	for {
		record, err := reader.Read()
		if errors.Is(err, io.EOF) {
			break
		} else if errors.Is(err, csv.ErrFieldCount) {
			if opts.strict {
				fmt.Fprintf(os.Stderr, "Error: row is missing column(s): %v\n", err)
				os.Exit(1)
			}
		} else if err != nil {
			fmt.Fprintf(os.Stderr, "Error: reading csv: %v\n", err)
			os.Exit(1)
		}

		if headers == nil {
			headers = record

			for _, arg := range flag.Args() {
				if i := slices.Index(headers, arg); i >= 0 {
					indexes = append(indexes, i)
				} else {
					i, err := strconv.Atoi(arg)
					if err != nil || i > reader.FieldsPerRecord {
						fmt.Fprintf(os.Stderr, "Error: invalid column %q\n", arg)
						os.Exit(1)
					}
					indexes = append(indexes, i)
				}
			}
		}

		if opts.listHeaders {
			_, _ = fmt.Println(strings.Join(headers, "\n"))
			return
		}

		var out []string
		if len(indexes) == 0 {
			out = record
		} else {
			for _, i := range indexes {
				if i < len(record) {
					out = append(out, record[i])
				}
			}
		}

		err = writer.Write(out)
		if errors.Is(err, syscall.EPIPE) {
			return
		} else if err != nil {
			fmt.Fprintf(os.Stderr, "Error: write row: %v\n", err)
			os.Exit(1)
		}
	}

	writer.Flush()

	if err := writer.Error(); err != nil && !errors.Is(err, syscall.EPIPE) {
		fmt.Fprintf(os.Stderr, "Error: write: %v\n", err)
		os.Exit(1)
	}
}
