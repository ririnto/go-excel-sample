package main

import (
	"fmt"
	"github.com/mattn/go-runewidth"
	"github.com/tealeg/xlsx"
	"gopkg.in/yaml.v3"
	"log"
	"math"
	"slices"
	"strings"
)

type Record struct {
	Name        string
	Description string
	Tags        []string
	Note        string
}

type ColumnDefinition struct {
	Header string
	Title  string
}

func sanitizeCellValue(value string) string {
	if strings.HasPrefix(value, "-") {
		return "'" + value
	}
	return value
}

func calculateDisplayWidth(content string) float64 {
	maxWidth := 0.0
	for line := range slices.Values(strings.Split(content, "\n")) {
		maxWidth = math.Max(maxWidth, float64(runewidth.StringWidth(line)))
	}
	return maxWidth
}

func mergeCells(row *xlsx.Row, startCol, endCol int, value string, style *xlsx.Style) {
	cell := row.AddCell()
	cell.Value = value
	cell.SetStyle(style)

	columnsToMerge := endCol - startCol
	for i := 0; i < columnsToMerge; i++ {
		row.AddCell().SetStyle(style)
	}

	cell.Merge(columnsToMerge, 0)
}

func addCell(row *xlsx.Row, value string, style *xlsx.Style, maxColumnWidths []float64, columnIndex int) {
	cell := row.AddCell()
	cell.Value = value
	cell.SetStyle(style)

	maxColumnWidths[columnIndex] = math.Max(maxColumnWidths[columnIndex], calculateDisplayWidth(strings.ReplaceAll(value, "'", "")))
}

func createTitleRow(sheet *xlsx.Sheet, style *xlsx.Style, columns []ColumnDefinition) error {
	row := sheet.AddRow()

	for i := 0; i < len(columns); i++ {
		if i == 0 {
			mergeCells(row, 0, 1, columns[0].Title, style)
			i++
		} else if i == 2 {
			mergeCells(row, 2, 3, columns[2].Title, style)
			i++
		} else {
			row.AddCell().SetStyle(style)
		}
	}

	return nil
}

func createHeaderRow(sheet *xlsx.Sheet, style *xlsx.Style, columns []ColumnDefinition, maxColumnWidths []float64) *xlsx.Row {
	row := sheet.AddRow()
	for i, col := range columns {
		addCell(row, sanitizeCellValue(col.Header), style, maxColumnWidths, i)
	}
	return row
}

func populateDataRows(sheet *xlsx.Sheet, style *xlsx.Style, records []Record, maxColumnWidths []float64) error {
	for record := range slices.Values(records) {
		row := sheet.AddRow()
		addCell(row, sanitizeCellValue(record.Name), style, maxColumnWidths, 0)
		addCell(row, sanitizeCellValue(record.Description), style, maxColumnWidths, 1)

		var tags string
		if len(record.Tags) == 1 {
			tags = record.Tags[0]
		} else if 1 < len(record.Tags) {
			if tagYaml, err := yaml.Marshal(record.Tags); err != nil {
				log.Printf("Failed to marshal tags: %v", err)
				return fmt.Errorf("failed to marshal tags: %w", err)
			} else {
				tags = strings.TrimSpace(string(tagYaml))
			}
		}
		addCell(row, sanitizeCellValue(tags), style, maxColumnWidths, 2)
		addCell(row, sanitizeCellValue(record.Note), style, maxColumnWidths, 3)
	}
	return nil
}

func adjustColumnWidths(sheet *xlsx.Sheet, maxColumnWidths []float64) error {
	for i, width := range maxColumnWidths {
		adjustedWidth := math.Min(50, math.Max(10, width*1.2))
		if err := sheet.SetColWidth(i, i, adjustedWidth); err != nil {
			return fmt.Errorf("failed to set column width for column %d: %w", i, err)
		}
	}
	return nil
}

func WriteRecordsToExcel(records []Record, filePath string) error {
	workbook := xlsx.NewFile()
	sheet, err := workbook.AddSheet("Sheet1")
	if err != nil {
		log.Printf("Failed to add sheet: %v", err)
		return fmt.Errorf("failed to add sheet: %w", err)
	}

	standardStyle := xlsx.NewStyle()
	standardStyle.Alignment.WrapText = true
	standardStyle.Alignment.Vertical = "center"

	titleStyle := xlsx.NewStyle()
	titleStyle.Alignment.WrapText = true
	titleStyle.Alignment.Vertical = "center"
	titleStyle.Alignment.Horizontal = "center"
	titleStyle.Font.Bold = true

	columns := []ColumnDefinition{
		{Header: "Name", Title: "Name, Description"},
		{Header: "Description", Title: "Name, Description"},
		{Header: "Tags", Title: "Tags, Note"},
		{Header: "Note", Title: "Tags, Note"},
	}

	maxColumnWidths := []float64{
		calculateDisplayWidth(columns[0].Header),
		calculateDisplayWidth(columns[1].Header),
		calculateDisplayWidth(columns[2].Header),
		calculateDisplayWidth(columns[3].Header),
	}

	if err := createTitleRow(sheet, titleStyle, columns); err != nil {
		log.Printf("Failed to create title row: %v", err)
		return fmt.Errorf("failed to create title row: %w", err)
	}

	createHeaderRow(sheet, standardStyle, columns, maxColumnWidths)

	if err := populateDataRows(sheet, standardStyle, records, maxColumnWidths); err != nil {
		log.Printf("Failed to populate data rows: %v", err)
		return fmt.Errorf("failed to populate data rows: %w", err)
	}

	if err := adjustColumnWidths(sheet, maxColumnWidths); err != nil {
		log.Printf("Failed to adjust column widths: %v", err)
		return fmt.Errorf("failed to adjust column widths: %w", err)
	}

	if err := workbook.Save(filePath); err != nil {
		log.Printf("Failed to save Excel file: %v", err)
		return fmt.Errorf("failed to save Excel file: %w", err)
	}

	return nil
}

func main() {
	records := []Record{
		{Name: "-Name1", Description: "First line\nSecond line", Tags: []string{"go", "excel", "example1"}, Note: "-Note1"},
		{Name: "NormalName", Description: "Single line description", Tags: []string{"yaml", "example2"}, Note: "Note2"},
		{Name: "-Name3", Description: "Multi-line\ntest\nexample", Tags: []string{"multi", "line"}, Note: "Additional note"},
		{Name: "SingleTagName", Description: "Single tag description", Tags: []string{}, Note: "Single tag note"},
	}
	if err := WriteRecordsToExcel(records, "example_records.xlsx"); err != nil {
		log.Fatalf("Failed to create Excel: %v", err)
	}
	log.Println("Excel file created successfully: example_records.xlsx")
}
