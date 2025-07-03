package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

const (
	// progressReportInterval defines how often (in number of rows) to report progress.
	progressReportInterval = 1000
)

func main() {
	// --- 1. Argument Parsing ---
	// Define command-line flags.
	outputFile := flag.String("o", "", "Path to the output CSV file. Use '-' for stdout. Defaults to <input_filename>.csv")
	sheetIndex := flag.Int("sheet-index", 0, "0-based index of the Excel sheet to convert")
	flag.Parse()

	// The input file is a required positional argument.
	if flag.NArg() != 1 {
		fmt.Fprintln(os.Stderr, "Usage: go run main.go [-o <output.csv>] [-sheet-index <index>] <input.xlsx | ->")
		flag.Usage()
		return
	}
	inputFile := flag.Arg(0)

	// Determine the final output path.
	actualOutputFile := *outputFile
	if actualOutputFile == "" && inputFile != "-" {
		// If output is not specified and input is not stdin, derive from input filename.
		base := filepath.Base(inputFile)
		actualOutputFile = strings.TrimSuffix(base, filepath.Ext(base)) + ".csv"
	} else if actualOutputFile == "" && inputFile == "-" {
		// If output is not specified and input is stdin, default to stdout.
		actualOutputFile = "-"
	}

	// --- 2. Open Input Stream ---
	var xlsxFile *excelize.File
	var err error

	if inputFile == "-" {
		fmt.Fprintln(os.Stderr, "Reading Excel data from stdin...")
		xlsxFile, err = excelize.OpenReader(os.Stdin)
	} else {
		xlsxFile, err = excelize.OpenFile(inputFile)
	}

	if err != nil {
		log.Fatalf("Failed to open Excel input: %v", err)
	}
	defer func() {
		if err := xlsxFile.Close(); err != nil {
			log.Printf("Error closing Excel file: %v\n", err)
		}
	}()

	// --- 3. Create Output Writer ---
	var csvWriter *csv.Writer
	var output io.WriteCloser

	if actualOutputFile == "-" {
		output = os.Stdout
		csvWriter = csv.NewWriter(output)
	} else {
		file, err := os.Create(actualOutputFile)
		if err != nil {
			log.Fatalf("Failed to create CSV file %s: %v", actualOutputFile, err)
		}
		output = file // We'll use this to defer the close.
		csvWriter = csv.NewWriter(output)
	}
	defer func() {
		// Flush any buffered data to the writer.
		csvWriter.Flush()
		// If the output is a file (not stdout), close it.
		if closer, ok := output.(io.Closer); ok && output != os.Stdout {
			if err := closer.Close(); err != nil {
				log.Printf("Error closing output file: %v\n", err)
			}
		}
	}()

	// --- 4. Stream and Convert Data ---
	// Get the list of sheets and select the one specified by the sheet-index flag.
	sheetList := xlsxFile.GetSheetList()
	if *sheetIndex < 0 || *sheetIndex >= len(sheetList) {
		log.Fatalf("Error: sheet-index %d is out of bounds. The workbook has %d sheets (indices 0 to %d).", *sheetIndex, len(sheetList), len(sheetList)-1)
	}
	sheetName := sheetList[*sheetIndex]

	rows, err := xlsxFile.Rows(sheetName)
	if err != nil {
		log.Fatalf("Failed to get rows iterator for sheet '%s': %v", sheetName, err)
	}

	inputDesc := fmt.Sprintf("file '%s'", inputFile)
	if inputFile == "-" {
		inputDesc = "stdin"
	}
	outputDesc := fmt.Sprintf("file '%s'", actualOutputFile)
	if actualOutputFile == "-" {
		outputDesc = "stdout"
	}
	fmt.Fprintf(os.Stderr, "Converting sheet '%s' (index %d) from %s to %s...\n", sheetName, *sheetIndex, inputDesc, outputDesc)

	var header []string
	var columnCount int
	rowCount := 0

	// --- Read Header and Determine Column Count ---
	if rows.Next() {
		header, err = rows.Columns()
		if err != nil {
			log.Fatalf("Failed to read header row: %v", err)
		}
		columnCount = len(header)
		if err := csvWriter.Write(header); err != nil {
			log.Fatalf("Error writing header to CSV: %v", err)
		}
		rowCount++
	} else {
		fmt.Fprintf(os.Stderr, "Warning: Sheet '%s' is empty.\n", sheetName)
		return
	}

	// --- Process Remaining Rows ---
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			log.Printf("Error reading row %d: %v", rowCount+1, err)
			continue
		}

		// Normalize row length to match the header.
		if len(row) < columnCount {
			paddedRow := make([]string, columnCount)
			copy(paddedRow, row)
			row = paddedRow
		} else if len(row) > columnCount {
			log.Printf("Warning: Row %d has %d columns, more than header's %d. Truncating.", rowCount+1, len(row), columnCount)
			row = row[:columnCount]
		}

		if err := csvWriter.Write(row); err != nil {
			log.Printf("Error writing row %d to CSV: %v", rowCount+1, err)
		}
		rowCount++

		// Report progress to stderr.
		if rowCount%progressReportInterval == 0 {
			fmt.Fprintf(os.Stderr, "... processed %d rows\n", rowCount)
			csvWriter.Flush()
		}
	}

	if err := rows.Close(); err != nil {
		log.Printf("Error closing row iterator: %v\n", err)
	}

	// --- 5. Finalization ---
	if err := csvWriter.Error(); err != nil {
		log.Fatalf("An error occurred during CSV writing: %v", err)
	}

	fmt.Fprintf(os.Stderr, "Conversion complete! Successfully wrote %d rows.\n", rowCount)
}
