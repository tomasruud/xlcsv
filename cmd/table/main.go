package main

import (
	"bufio"
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"io"
	"os"
	"os/signal"
	"strings"
	"syscall"
	"unicode/utf8"
)

func usage() {
	fmt.Fprint(os.Stderr, `table - pretty print CSV as table

USAGE:
  table [OPTIONS]

OPTIONS:
  -h, --help
      Prints this help information

  -d, --delimiter <value>
      CSV delimiter, defaults to ",", can be set using the CSV_DELIMITER env var
`)
}

func main() {
	stop := make(chan os.Signal, 1)
	signal.Notify(stop, os.Interrupt, syscall.SIGTERM)

	signal.Ignore(syscall.SIGPIPE)

	var opts options

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

	if err := run(os.Stdin, os.Stdout, stop, opts); err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
}

type options struct {
	delimiter string
}

func run(in io.Reader, out io.Writer, stop chan os.Signal, opts options) error {
	tmp, err := os.CreateTemp("", "csv-table-*.csv")
	if err != nil {
		return fmt.Errorf("create temp file: %v", err)
	}
	defer tmp.Close()
	defer os.Remove(tmp.Name())

	go func() {
		<-stop
		fmt.Fprintln(os.Stderr, "Removing temp files for shutdown")
		_ = tmp.Close()
		_ = os.Remove(tmp.Name())
		os.Exit(1)
	}()

	if _, err := io.Copy(tmp, in); err != nil {
		return fmt.Errorf("copy data: %v", err)
	}

	reader := csv.NewReader(tmp)
	reader.Comma = rune(opts.delimiter[0])
	reader.ReuseRecord = true
	reader.Comment = '#'

	if _, err := tmp.Seek(0, io.SeekStart); err != nil {
		return fmt.Errorf("rewinding buffer: %v", err)
	}

	var widths []int
	for {
		record, err := reader.Read()
		if errors.Is(err, io.EOF) {
			break
		} else if err != nil && !errors.Is(err, csv.ErrFieldCount) {
			return fmt.Errorf("reading csv: %v", err)
		}

		for i, col := range record {
			w := utf8.RuneCountInString(col)
			if i >= len(widths) {
				widths = append(widths, w)
			} else if w > widths[i] {
				widths[i] = w
			}
		}
	}

	writer := bufio.NewWriter(out)

	var fmts []string
	var ends string
	for _, width := range widths {
		fmts = append(fmts, fmt.Sprintf("| %%-%ds ", width))
		ends += "+" + strings.Repeat("-", width+2)
	}
	ends += "+"

	_, _ = fmt.Fprintln(writer, ends)

	if _, err := tmp.Seek(0, io.SeekStart); err != nil {
		return fmt.Errorf("rewinding buffer: %v", err)
	}

	for {
		record, err := reader.Read()
		if errors.Is(err, io.EOF) {
			break
		} else if err != nil && !errors.Is(err, csv.ErrFieldCount) {
			return fmt.Errorf("reading csv: %v", err)
		}

		for i, f := range fmts {
			var v string
			if i < len(record) {
				v = record[i]
			} else {
				v = ""
			}

			_, err = fmt.Fprintf(writer, f, v)
			if errors.Is(err, syscall.EPIPE) {
				return nil
			} else if err != nil {
				return fmt.Errorf("writing row: %v", err)
			}
		}

		_, err = fmt.Fprintln(writer, "|")
		if errors.Is(err, syscall.EPIPE) {
			return nil
		} else if err != nil {
			return fmt.Errorf("writing row: %v", err)
		}
	}

	_, _ = fmt.Fprintln(writer, ends)

	err = writer.Flush()
	if errors.Is(err, syscall.EPIPE) {
		return nil
	} else if err != nil {
		return fmt.Errorf("flushing output: %v", err)
	}

	return nil
}
