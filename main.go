package main

import (
	"flag"
	"log"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	start := time.Now()

	var inputFileFlag = flag.String("input", "", "Enter the name of the file you would like to read from")
	var outputFileFlag = flag.String("output", "", "Enter the name of the file you would like to create and output the data")

	flag.Parse()

	// Open input file
	inputf, err := excelize.OpenFile(*inputFileFlag)
	if err != nil {
		log.Printf("Error reading file %s: %v", *inputFileFlag, err)
		return
	}

	// Create new file
	outputf := excelize.NewFile()

	// Get list of sheets from the input file
	sheets := inputf.GetSheetList()

	// Loop through each sheet and perform various operations
	for sheetIndex, sheet := range sheets {

		// Create the corresponding sheet in the output file
		outputf.NewSheet(sheet)

		// Get Rows for this sheet
		rows, err := inputf.GetRows(sheet)
		if err != nil {
			log.Printf("Error reading rows %v", err)
			return
		}

		// Loop through rows to perform operations
		for rowIndex, row := range rows {

			// Set width of all columns
			lastColName, _ := excelize.ColumnNumberToName(len(row))
			outputf.SetColWidth(sheet, "A", lastColName, 45)

			// Loop through each column to copy values
			for colIndex, cellValue := range row {

				colName, _ := excelize.ColumnNumberToName(colIndex + 1)
				rowNumString := strconv.Itoa(rowIndex + 1)
				axis := colName + rowNumString

				// Make the first row cells text Bold apart from the first sheet
				if rowIndex == 0 && sheetIndex != 0 {
					style, _ := outputf.NewStyle(`{"font":{"bold": true}}`)
					outputf.SetCellStyle(sheet, axis, axis, style)
				}

				outputf.SetCellValue(sheet, axis, cellValue)
			}
		}

	}

	// Delete auto-generated sheet Sheet1
	outputf.DeleteSheet("Sheet1")

	// Set first sheet as active
	outputf.SetActiveSheet(0)

	// Save file to output file name specified by user
	if err := outputf.SaveAs(*outputFileFlag); err != nil {
		log.Println(err)
	}

	log.Println("Executed in: ", time.Since(start))
}
