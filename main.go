package main

import (
	"errors"
	"flag"
	"fmt"
	"path/filepath"
	"strings"

	"github.com/wyattis/z/zflag"
	"github.com/xuri/excelize/v2"
)

var filesA = zflag.StringSlice()
var filesB = zflag.StringSlice()
var sheet = "Sheet1"
var shouldPrintDiff = false
var shouldColorSheet = false
var color = "E0EBF5"

type CellDiff struct {
	Cell string
	Row  int
	Col  int
	ValA string
	ValB string
}

func getCellVal(row, col int) string {
	// TODO: this doesn't work if we get into the double letter rows
	return fmt.Sprintf("%s%d", string('A'+col), row+1)
}

func expandGlobs(patterns []string) (res []string, err error) {
	for _, pattern := range patterns {
		matches, err := filepath.Glob(pattern)
		if err != nil {
			return nil, err
		}
		res = append(res, matches...)
	}
	return
}

func run() (err error) {
	if filesA.Len() == 0 || filesB.Len() == 0 {
		return errors.New("Must specify at least one file for each side")
	}
	if filesA.Len() != filesB.Len() {
		return errors.New("Length of files to compare don't match")
	}
	filesA, filesB := filesA.Val(), filesB.Val()

	filesA, err = expandGlobs(filesA)
	if err != nil {
		return
	}
	filesB, err = expandGlobs(filesB)
	if err != nil {
		return
	}

	for i := 0; i < len(filesA); i++ {
		fmt.Printf("comparing %s with %s\n", filesA[i], filesB[i])
		diff, err := compareFiles(filesA[i], filesB[i], sheet, sheet)
		if err != nil {
			return err
		}
		if shouldPrintDiff {
			for _, d := range diff {
				fmt.Printf("%s (%d, %d): '%s' != '%s'\n", d.Cell, d.Row, d.Col, d.ValA, d.ValB)
			}
		} else if len(diff) > 0 {
			fmt.Printf("found %d differences in '%s'\n", len(diff), sheet)
		} else {
			fmt.Println("Files were the same")
		}
		// fmt.Println(diff)
	}
	return
}

func compareFiles(fileA, fileB string, sheetA, sheetB string) (diff []CellDiff, err error) {
	a, err := excelize.OpenFile(fileA)
	if err != nil {
		return
	}
	b, err := excelize.OpenFile(fileB)
	if err != nil {
		return
	}
	diff, err = compareSheets(a, b, sheetA, sheetB)
	if err != nil {
		return
	}
	if shouldColorSheet && len(diff) > 0 {
		newName := strings.ReplaceAll(fileB, ".xlsx", ".diff.xlsx")
		if err = b.SaveAs(newName); err != nil {
			return nil, err
		}
	}
	return
}

func compareSheets(a, b *excelize.File, sheetA, sheetB string) (diff []CellDiff, err error) {
	aDim, err := a.GetSheetDimension(sheetA)
	if err != nil {
		return
	}
	bDim, err := b.GetSheetDimension(sheetB)
	if err != nil {
		return
	}
	// fmt.Println("sheet dimensions", aDim, bDim)
	if aDim != bDim {
		err = errors.New("dimensions don't match")
		return
	}
	aRows, err := a.Rows(sheetA)
	if err != nil {
		return
	}
	defer aRows.Close()
	bRows, err := b.Rows(sheetB)
	if err != nil {
		return
	}
	defer bRows.Close()
	row := 0
	// TODO: We should inherit the style of the cell from the original sheet
	style := &excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{color}, Pattern: 1},
	}
	if err != nil {
		return
	}
	for aRows.Next() && bRows.Next() {
		aCols, err := aRows.Columns()
		if err != nil {
			return nil, err
		}
		bCols, err := bRows.Columns()
		if err != nil {
			return nil, err
		}
		if len(aCols) != len(bCols) {
			return nil, errors.New("Found row size that doesn't match")
		}
		for col := 0; col < len(aCols); col++ {
			aVal, bVal := aCols[col], bCols[col]
			if aCols[col] != bCols[col] {
				diff = append(diff, CellDiff{
					Cell: getCellVal(row, col),
					Row:  row,
					Col:  col,
					ValA: aVal,
					ValB: bVal,
				})
				if shouldColorSheet {
					s, err := b.NewStyle(style)
					if err != nil {
						return nil, err
					}
					err = b.SetCellStyle(sheetB, getCellVal(row, col), getCellVal(row, col), s)
					if err != nil {
						return nil, err
					}
				}
			}
		}
		row++
	}

	return
}

var Example = `
Example usage:

	Different file combinations to compare:
		excel-compare.exe -a a.xlsx -b b.xlsx -sheet Sheet1
		excel-compare.exe -a a.xlsx -b b.xlsx -a a2.xlsx -b b2.xlsx
		excel-compare.exe -a folder_a/*.xlsx -b folder_b/*.xlsx

	See the diff in the console:
		excel-compare.exe -a a.xlsx -b b.xlsx -print-diff

	Modify the color of cells where changes occurred in the B group:
		excel-compare.exe -a a.xlsx -b b.xlsx -color-sheet
		excel-compare.exe -a a.xlsx -b b.xlsx -color-sheet -color FFE0E0
`

func main() {
	flag.Var(filesA, "a", "left input files")
	flag.Var(filesB, "b", "right input files")
	flag.BoolVar(&shouldPrintDiff, "print-diff", false, "print the differences")
	flag.BoolVar(&shouldColorSheet, "color-sheet", false, "change the color of cells where changes occurred in the B group (saved to new file)")
	flag.StringVar(&color, "color", "E0EBF5", "color to use for coloring the differences in the sheet")
	flag.StringVar(&sheet, "sheet", "Sheet1", "sheet to compare")
	flag.Parse()
	if err := run(); err != nil {
		fmt.Println(err)
		fmt.Println()
		flag.Usage()
		fmt.Println(Example)
	}
}
