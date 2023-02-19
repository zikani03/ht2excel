package main

import (
	"flag"
	"fmt"
	"os"

	"github.com/PuerkitoBio/goquery"
	"github.com/sirupsen/logrus"
	"github.com/xuri/excelize/v2"
)

var htmlFilePath string
var outputFile string

func init() {
	flag.StringVar(&htmlFilePath, "f", "", "HTML Data file")
	flag.StringVar(&outputFile, "o", "test.xlsx", "Output file name")
}

type state struct {
	TableNameInferred string
	IsFirstRow        bool
	CurrentRowNum     int
	NumTablesFound    int
	NumRowsProcessed  int
	HasTHead          bool
	HasTBody          bool
	HasTFooter        bool
}

func main() {
	flag.Parse()

	htmlFile, err := os.Open(htmlFilePath)
	if err != nil {
		logrus.WithError(err)
		return
	}

	doc, err := goquery.NewDocumentFromReader(htmlFile)
	if err != nil {
		logrus.WithError(err)
		return
	}

	st := &state{}

	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			logrus.WithError(err)
			return
		}
	}()

	doc.Find("table").Each(func(i int, tbl *goquery.Selection) {
		st.NumTablesFound = st.NumTablesFound + 1
		if st.CurrentRowNum == 0 {
			st.IsFirstRow = true
		}

		sheetName := tbl.AttrOr("data-sheet-name", fmt.Sprintf("Sheet %d", st.NumTablesFound))
		index, err := f.NewSheet(sheetName)
		if err != nil {
			logrus.WithError(err)
			return
		}

		for idx, row := range toExcelSheetData(tbl) {
			cell, err := excelize.CoordinatesToCellName(1, idx+1)
			if err != nil {
				logrus.WithError(err)
				return
			}
			f.SetSheetRow(sheetName, cell, &row)
		}

		// Set active sheet of the workbook.
		f.SetActiveSheet(index)
		// Save spreadsheet by the given path.
		tmpFile, err := os.CreateTemp(os.TempDir(), "_generated_payroll.xlsx")
		if err != nil {
			logrus.WithError(err)
			return
		}
		fullName := tmpFile.Name() + ".xlsx"
		err = f.SaveAs(fullName)
		if err != nil {
			logrus.WithError(err)
			return
		}

		st.CurrentRowNum++
	})

	// Set active sheet of the workbook.
	f.SetActiveSheet(st.NumTablesFound)
	err = f.SaveAs(outputFile)
	if err != nil {
		logrus.WithError(err)
		os.Exit(127)
	}
}

func toExcelSheetData(tbl *goquery.Selection) [][]any {
	sheetData := make([][]any, 0)
	// TODO: get the thead > tr > th : headers and make them excel headers
	// TODO: handle situation where the headers are not in a thead. Default/fallback to using first row as headers
	headerRows := tbl.Find("thead > tr")
	if headerRows != nil {
		if headerRows.Length() > 1 {
			// TODO: handle multiple header rows...

		} else {
			firstRow := headerRows.First()
			tds := firstRow.Children()
			tdsExcelRow := make([]any, 0)

			tds.Each(func(i int, td *goquery.Selection) {
				tdsExcelRow = append(tdsExcelRow, td.Text())
			})

			sheetData = append(sheetData, tdsExcelRow)
		}
	}

	// TODO: get the tbody > tr > td : data and make them excel rows
	// TODO: how do we handle colspan columns...
	// TODO: how do we handle caption elements in the table
	// TODO: how to handle inner HTML in the table data. (support a ".excel-data" field)
	bodyRows := tbl.Find("tbody > tr")
	if bodyRows != nil {
		if bodyRows.Length() > 0 {
			bodyRows.Each(func(i int, tr *goquery.Selection) {
				tds := tr.Children()
				tdsExcelRow := make([]any, 0)

				tds.Each(func(i int, td *goquery.Selection) {
					tdsExcelRow = append(tdsExcelRow, td.Text())
				})

				sheetData = append(sheetData, tdsExcelRow)
			})
		}
	}
	return sheetData
}
