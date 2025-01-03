# Excel Records Writer

이 Go 프로그램은 구조화된 레코드를 Excel 파일로 작성합니다. 셀 너비, 스타일 및 데이터 직렬화를 관리하기 위해 여러 라이브러리를 사용합니다.

## 목차

- [코드 구조](#코드-구조)
    - [Imports](#imports)
    - [데이터 구조](#데이터-구조)
    - [유틸리티 함수](#유틸리티-함수)
        - [sanitizeCellValue](#sanitizecellvalue)
        - [calculateDisplayWidth](#calculatedisplaywidth)
        - [mergeCells](#mergecells)
        - [addCell](#addcell)
    - [행 생성 함수](#행-생성-함수)
        - [createTitleRow](#createtitlerow)
        - [createHeaderRow](#createheaderrow)
        - [populateDataRows](#populatedatarows)
        - [adjustColumnWidths](#adjustcolumnwidths)
    - [주요 기능](#주요-기능)
- [예제](#예제)
- [종속성](#종속성)

## 코드 구조

### Imports

<!-- @formatter:off -->
```go
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
```
<!-- @formatter:on -->

- **fmt**: 포맷된 입출력용.
- **runewidth**: 유니코드 문자를 고려한 문자열의 표시 너비 계산.
- **xlsx**: Excel 파일 생성 및 조작.
- **yaml.v3**: 태그 슬라이스를 YAML 형식으로 직렬화.
- **log**: 오류 및 정보 로깅.
- **math**: 수학적 연산.
- **slices**: 슬라이스 유틸리티 함수 제공.
- **strings**: 문자열 조작.

### 데이터 구조

<!-- @formatter:off -->
```go
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
```
<!-- @formatter:on -->

- **Record**: 이름, 설명, 태그, 노트 필드를 가진 단일 데이터 레코드를 나타냅니다.
- **ColumnDefinition**: Excel 시트의 각 열에 대한 헤더와 제목을 정의합니다.

### 유틸리티 함수

#### sanitizeCellValue

<!-- @formatter:off -->
```go
func sanitizeCellValue(value string) string {
    if strings.HasPrefix(value, "-") {
        return "'" + value
    }
    return value
}
```
<!-- @formatter:on -->

- **목적**: 하이픈(`-`)으로 시작하는 셀 값을 Excel에서 문자열로 처리하도록 단일 인용부호를 접두사로 추가합니다. 이는 Excel이 이를 수식이나 숫자로 해석하지 않도록 방지합니다.

#### calculateDisplayWidth

<!-- @formatter:off -->
```go
func calculateDisplayWidth(content string) float64 {
    maxWidth := 0.0
    for line := range slices.Values(strings.Split(content, "\n")) {
        maxWidth = math.Max(maxWidth, float64(runewidth.StringWidth(line)))
    }
    return maxWidth
}
```
<!-- @formatter:on -->

- **목적**: 다중 줄 텍스트의 모든 줄을 고려하여 셀 내용의 최대 표시 너비를 계산합니다. 유니코드 문자를 고려하기 위해 `runewidth`를 사용합니다.

#### mergeCells

<!-- @formatter:off -->
```go
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
```
<!-- @formatter:on -->

- **목적**: 주어진 행에서 `startCol`부터 `endCol`까지 수평으로 셀을 병합합니다. 병합된 셀에 값을 설정하고 지정된 스타일을 적용합니다.

#### addCell

<!-- @formatter:off -->
```go
func addCell(row *xlsx.Row, value string, style *xlsx.Style, maxColumnWidths []float64, columnIndex int) {
    cell := row.AddCell()
    cell.Value = value
    cell.SetStyle(style)

    maxColumnWidths[columnIndex] = math.Max(maxColumnWidths[columnIndex], calculateDisplayWidth(strings.ReplaceAll(value, "'", "")))
}
```
<!-- @formatter:on -->

- **목적**: 주어진 값과 스타일로 행에 셀을 추가합니다. 각 열에 필요한 최대 너비를 추적하기 위해 `maxColumnWidths`를 업데이트합니다.

### 행 생성 함수

#### createTitleRow

<!-- @formatter:off -->
```go
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
```
<!-- @formatter:on -->

- **목적**: `ColumnDefinition`을 기반으로 특정 열을 병합하여 제목 행을 생성합니다. 예를 들어, 첫 번째 제목은 0번과 1번 열을 병합하고, 두 번째 제목은 2번과 3번 열을 병합합니다.

#### createHeaderRow

<!-- @formatter:off -->
```go
func createHeaderRow(sheet *xlsx.Sheet, style *xlsx.Style, columns []ColumnDefinition, maxColumnWidths []float64) *xlsx.Row {
    row := sheet.AddRow()
    for i, col := range columns {
        addCell(row, sanitizeCellValue(col.Header), style, maxColumnWidths, i)
    }
    return row
}
```
<!-- @formatter:on -->

- **목적**: 열 헤더로 헤더 행을 생성합니다. 헤더 값을 정리하고 최대 열 너비를 업데이트합니다.

#### populateDataRows

<!-- @formatter:off -->
```go
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
```
<!-- @formatter:on -->

- **목적**: `records`의 데이터를 시트에 채웁니다. `Tags` 필드는 태그가 여러 개인 경우 YAML로 직렬화하여 처리합니다. 각 셀을 정리하고 행에 추가합니다.

#### adjustColumnWidths

<!-- @formatter:off -->
```go
func adjustColumnWidths(sheet *xlsx.Sheet, maxColumnWidths []float64) error {
    for i, width := range maxColumnWidths {
        adjustedWidth := math.Min(50, math.Max(10, width*1.2))
        if err := sheet.SetColWidth(i, i, adjustedWidth); err != nil {
            return fmt.Errorf("failed to set column width for column %d: %w", i, err)
        }
    }
    return nil
}
```
<!-- @formatter:on -->

- **목적**: 각 열의 너비를 콘텐츠의 최대 너비에 따라 조정하며, 너비가 10에서 50 단위 사이에 유지되도록 합니다.

### 주요 기능

#### WriteRecordsToExcel

<!-- @formatter:off -->
```go
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
```
<!-- @formatter:on -->

- **목적**: Excel 파일 생성을 주도하는 핵심 함수입니다. 워크북과 시트를 초기화하고, 스타일을 정의하며, 열 정의를 설정하고, 제목 및 헤더 행을 생성한 후 데이터를 채우고, 열 너비를 조정하고, 지정된
  파일 경로에 워크북을 저장합니다.

#### main

<!-- @formatter:off -->
```go
func main() {
    records := []Record{
        {Name: "-Name1", Description: "First line\nSecond line", Tags: []string{"go", "excel", "example1"}, Note: "-Note1"},
        {Name: "NormalName", Description: "Single line description", Tags: []string{"yaml"}, Note: "Note2"},
        {Name: "-Name3", Description: "Multi-line\ntest\nexample", Tags: []string{"multi", "line"}, Note: "Additional note"},
        {Name: "SingleTagName", Description: "Single tag description", Tags: []string{}, Note: "Single tag note"},
    }
    if err := WriteRecordsToExcel(records, "example_records.xlsx"); err != nil {
        log.Fatalf("Failed to create Excel: %v", err)
    }
    log.Println("Excel file created successfully: example_records.xlsx")
}
```
<!-- @formatter:on -->

- **목적**: 샘플 레코드를 정의하고 `WriteRecordsToExcel`을 호출하여 Excel 파일을 생성합니다. 성공 또는 오류를 로그로 기록합니다.

## 예제

프로그램을 실행하면 `example_records.xlsx` 파일이 다음과 같은 구조로 생성됩니다:

| Name          | Description                   | Tags                          | Note            |
|---------------|-------------------------------|-------------------------------|-----------------|
| -Name1        | First line<br>Second line     | - go<br>- excel<br>- example1 | -Note1          |
| NormalName    | Single line description       | yaml                          | Note2           |
| -Name3        | Multi-line<br>test<br>example | - multi<br>- line             | Additional note |
| SingleTagName | Single tag description        |                               | Single tag note |

*참고: `-`으로 시작하는 셀은 문자열로 처리됩니다.*

## 종속성

- [go-runewidth](https://github.com/mattn/go-runewidth): 유니코드 문자를 고려한 문자열의 표시 너비 계산.
- [tealeg/xlsx](https://github.com/tealeg/xlsx): Excel 파일 생성 및 조작.
- [yaml.v3](https://pkg.go.dev/gopkg.in/yaml.v3): 데이터 구조를 YAML 형식으로 직렬화.
- [golang.org/x/exp/slices](https://pkg.go.dev/golang.org/x/exp/slices): 일반적인 슬라이스 함수 제공.
