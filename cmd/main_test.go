package main

import (
	"testing"

	"github.com/sirupsen/logrus"
	"github.com/xuri/excelize/v2"
)

func Test_initProblemCellNum(t *testing.T) {
	fileResult := "../result_tmp.xlsx"
	fResult, err := excelize.OpenFile(fileResult)
	if err != nil {
		logrus.Fatalf("打开 Excel 文件异常，原因: %v", err)
	}
	defer func() {
		// Close the spreadsheet.
		if err := fResult.Close(); err != nil {
			logrus.Errorln(err)
			return
		}
	}()

	initScoreAnalysisSheetCellNum(fResult)

}
