package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"

	"github.com/xuri/excelize/v2"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Println("使用方法: go run main.go <Excel文件路径>")
		fmt.Println("示例: go run main.go example.xlsx")
		return
	}

	// 获取输入文件路径
	inputFile := os.Args[1]

	// 打开Excel文件
	f, err := excelize.OpenFile(inputFile)
	if err != nil {
		log.Fatalf("无法打开Excel文件: %s", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			log.Fatalf("关闭文件时出错: %s", err)
		}
	}()

	// 获取所有sheet名
	sheets := f.GetSheetList()
	if len(sheets) <= 1 {
		fmt.Println("文件中只有一个sheet，无需合并")
		return
	}

	// 第一个sheet (目标sheet)
	targetSheet := sheets[0]
	fmt.Printf("将合并到目标sheet: %s\n", targetSheet)

	// 获取目标sheet中最后一行的行号
	rows, err := f.GetRows(targetSheet)
	if err != nil {
		log.Fatalf("获取目标sheet行时出错: %s", err)
	}
	nextRow := len(rows) + 1
	if nextRow == 1 {
		// 如果目标sheet是空的，从第1行开始
		nextRow = 1
	}

	// 处理每个源sheet
	for i := 1; i < len(sheets); i++ {
		sourceSheet := sheets[i]
		fmt.Printf("正在处理sheet: %s\n", sourceSheet)

		// 获取源sheet的所有行
		sourceRows, err := f.GetRows(sourceSheet)
		if err != nil {
			log.Printf("获取sheet %s 的行时出错: %s, 跳过此sheet", sourceSheet, err)
			continue
		}

		// 如果目标sheet是空的，需要复制表头
		if nextRow == 1 && len(sourceRows) > 0 {
			for colIdx, cellValue := range sourceRows[0] {
				colName, err := excelize.ColumnNumberToName(colIdx + 1)
				if err != nil {
					log.Printf("转换列索引时出错: %s", err)
					continue
				}
				cellRef := colName + "1"
				if err := f.SetCellValue(targetSheet, cellRef, cellValue); err != nil {
					log.Printf("设置单元格 %s 值时出错: %s", cellRef, err)
				}
			}
			nextRow = 2 // 表头已复制，从第2行开始
		}

		// 从源sheet的第一行开始复制(因为只有第一个sheet有表头，其他sheet没有表头)
		startRowIdx := 0

		// 复制行到目标sheet
		for rowIdx := startRowIdx; rowIdx < len(sourceRows); rowIdx++ {
			for colIdx, cellValue := range sourceRows[rowIdx] {
				colName, err := excelize.ColumnNumberToName(colIdx + 1)
				if err != nil {
					log.Printf("转换列索引时出错: %s", err)
					continue
				}
				cellRef := colName + fmt.Sprintf("%d", nextRow)
				if err := f.SetCellValue(targetSheet, cellRef, cellValue); err != nil {
					log.Printf("设置单元格 %s 值时出错: %s", cellRef, err)
				}
			}
			nextRow++
		}
	}

	// 生成输出文件名
	ext := filepath.Ext(inputFile)
	base := inputFile[:len(inputFile)-len(ext)]
	outputFile := base + "_merged" + ext

	// 删除其他sheet，只保留第一个sheet
	for i := len(sheets) - 1; i > 0; i-- {
		if err := f.DeleteSheet(sheets[i]); err != nil {
			log.Printf("删除sheet '%s' 时出错: %s", sheets[i], err)
		} else {
			fmt.Printf("已删除sheet: %s\n", sheets[i])
		}
	}

	// 保存结果
	if err := f.SaveAs(outputFile); err != nil {
		log.Fatalf("保存文件时出错: %s", err)
	}

	fmt.Printf("成功将所有sheet合并到 '%s'，结果保存在: %s\n", targetSheet, outputFile)
}
