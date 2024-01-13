package main

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/sirupsen/logrus"
	"github.com/spf13/pflag"
	"github.com/xuri/excelize/v2"

	logging "github.com/DesistDaydream/logging/pkg/logrus_init"
)

var (
	logFlags logging.LogrusFlags
)

type questionCellNum struct {
	classRateColStart int
	gradeRateColStart int
}

// TODO: 读取 result_tmp.xlsx 文件，从文件中获取单元格的坐标位置
var pcm = map[string]questionCellNum{
	// "计算题": {
	// 	classRateColStart: 7,
	// 	gradeRateColStart: 4,
	// },
	// "填空题": {
	// 	classRateColStart: 13,
	// 	gradeRateColStart: 10,
	// },
	// "判断题": {
	// 	classRateColStart: 39,
	// 	gradeRateColStart: 36,
	// },
	// "选择题": {
	// 	classRateColStart: 49,
	// 	gradeRateColStart: 46,
	// },
	// "解决问题": {
	// 	classRateColStart: 59,
	// 	gradeRateColStart: 56,
	// },
}

// TODO: 获取成绩分析表特定单元格坐标
func initScoreAnalysisSheetCellNum(f *excelize.File, questionTypes []string) {
	// sheetList := []string{"年级各题得分率","班级各题得分率"}
	// sheetList := f.GetSheetList()

	// pcm1 := make(map[string]questionCellNum)

	var (
		rowsGrade, rowsClass [][]string
		err                  error
	)

	if rowsGrade, err = f.GetRows("年级各题得分率"); err != nil {
		logrus.Fatalf("获取年级各题得分率数据异常: %v", err)
	}
	if rowsClass, err = f.GetRows("班级各题得分率"); err != nil {
		logrus.Fatalf("获取班级各题得分率数据异常: %v", err)
	}

	for _, qt := range questionTypes {
		var (
			classRateColStart, gradeRateColStart int
		)

		for colIndex, row := range rowsGrade[1] {
			if row == fmt.Sprintf("%v总分", qt) {
				gradeRateColStart = colIndex + 1
				break
			}
		}

		for colIndex, row := range rowsClass[1] {
			if row == fmt.Sprintf("%v总分", qt) {
				classRateColStart = colIndex + 1
				break
			}
		}

		pcm[qt] = questionCellNum{
			classRateColStart: classRateColStart,
			gradeRateColStart: gradeRateColStart,
		}
	}

	logrus.Debugf("检查获取的单元格列号: %v", pcm)
}

type examination struct {
	questionTypes []string                // 题目类型列表
	questions     map[string]questionInfo // 题目类型对应的该题目的相关信息
	classTotal    int                     // 班级总数
	classSize     map[string]int          // 每个班级的人数
}

type questionInfo struct {
	questionNumbers []int       // 某题型的题号
	questionScore   map[int]int // 题目号对应的题目分值
}

func NewExamination(f *excelize.File) *examination {
	var questionTypes []string
	questions := make(map[string]questionInfo)
	rows, err := f.GetRows("info")
	if err != nil {
		logrus.Errorf("get cols failed, err:%v", err)
	}

	for i, row := range rows {
		var qi questionInfo

		if i == 0 {
			continue
		}

		if row == nil {
			break
		}

		// 获取题号
		qNums, _ := f.GetCellValue("info", "B"+strconv.Itoa(i+1))
		qNum := strings.Fields(qNums)
		for _, n := range qNum {
			ni, err := strconv.Atoi(n)
			if err != nil {
				logrus.Fatalf("将题号 %v 转为数字异常: %v", n, err)
			}
			qi.questionNumbers = append(qi.questionNumbers, ni)
		}

		// 获取题号对应的分值
		qScores, _ := f.GetCellValue("info", "C"+strconv.Itoa(i+1))
		qScore := strings.Fields(qScores)
		qi.questionScore = make(map[int]int)
		for i, s := range qScore {
			si, err := strconv.Atoi(s)
			if err != nil {
				logrus.Fatalf("将题号对应分值 %v 转为数字异常: %v", s, err)
			}
			qi.questionScore[qi.questionNumbers[i]] = si
		}

		// 题型名称
		questionTypeName := row[0]
		questionTypes = append(questionTypes, questionTypeName)
		// 为每个题型添加题型信息
		questions[questionTypeName] = qi
	}

	if logrus.GetLevel() == logrus.DebugLevel {
		logrus.Debugf("共有 %v 种题型", len(questions))
		for k, v := range questions {
			logrus.WithFields(logrus.Fields{}).Debugf("%v 有 %v 道: %v", k, len(v.questionNumbers), v.questionNumbers)
			for k1, v1 := range v.questionScore {
				logrus.Debugf("第 %v 题分值为: %v", k1, v1)
			}
		}
	}

	// 获取班级总数
	classTotal, err := f.GetCellValue("info", "D2")
	if err != nil {
		logrus.Errorf("获取班级数异常: %v", err)
	}
	c, _ := strconv.Atoi(classTotal)

	// 获取每班人数
	sheets := f.GetSheetList()
	css := make(map[string]int)
	classSizeCellValue, _ := f.GetCellValue("info", "E2")
	classSizes := strings.Split(classSizeCellValue, " ")
	for i, classSize := range classSizes {
		cs, _ := strconv.Atoi(classSize)
		css[sheets[i+1]] = cs
	}

	return &examination{
		questionTypes: questionTypes,
		questions:     questions,
		classTotal:    c,
		classSize:     css,
	}
}

// 计算班级各题型的得分情况
func (e *examination) calculateOfQuestionTypeWithClass(questionType string, sheetName string, classRowNum int, f *excelize.File, fResult *excelize.File) {
	var (
		questionAllTotalScore   float64
		questionTypeActualScore float64
	)

	questionColStart := pcm[questionType].classRateColStart + 2
	pt := e.questions[questionType]
	for _, pn := range pt.questionNumbers {
		col, _ := excelize.ColumnNumberToName(pn + 3)
		sum := float64(0)
		for i := 0; i < e.classSize[sheetName]; i++ {
			v, err := f.GetCellValue(sheetName, col+strconv.Itoa(i+2))
			if err != nil {
				logrus.Errorf("获取单元格值异常: %v", err)
			}
			vf, _ := strconv.ParseFloat(v, 64)
			sum += float64(vf)
		}

		totalScore := (float64(pt.questionScore[pn]) * float64(e.classSize[sheetName]))      // 总分数
		actualScore := (float64(pt.questionScore[pn])*float64(e.classSize[sheetName]) - sum) // 实得分
		rate := (actualScore / totalScore) * 100                                             // 得分率

		logrus.Printf("第 %v 题 %v 总得分 %v，得分率为 %0.2f%%", pn, questionType, actualScore, rate)

		resultScoreCol, _ := excelize.ColumnNumberToName(questionColStart)
		resultRateCol, _ := excelize.ColumnNumberToName(questionColStart + 1)
		fResult.SetCellStr("班级各题得分率", resultScoreCol+strconv.Itoa(classRowNum), fmt.Sprintf("%0.2f", actualScore)) // 第 X 题实得分
		fResult.SetCellStr("班级各题得分率", resultRateCol+strconv.Itoa(classRowNum), fmt.Sprintf("%0.2f%%", rate))       // 第 X 题得分率
		questionColStart += 2

		questionAllTotalScore += totalScore
		questionTypeActualScore += actualScore
	}

	questionScoreCol, _ := excelize.ColumnNumberToName(pcm[questionType].classRateColStart)
	questionRateCol, _ := excelize.ColumnNumberToName(pcm[questionType].classRateColStart + 1)
	questionTypeRate := (questionTypeActualScore / questionAllTotalScore) * 100
	fResult.SetCellStr("班级各题得分率", questionScoreCol+strconv.Itoa(classRowNum), fmt.Sprintf("%0.2f", questionTypeActualScore)) // 某题型实得分
	fResult.SetCellStr("班级各题得分率", questionRateCol+strconv.Itoa(classRowNum), fmt.Sprintf("%0.2f%%", questionTypeRate))       // 某题型得分率
}

// 计算年级各题型的得分情况
func (e *examination) calculateOfQuestionTypeWithGrade(questionType string, sheets []string, f *excelize.File, fResult *excelize.File) {
	var (
		questionAllTotalScore  float64
		questionAllActualScore float64
		calculationRate        float64
	)

	pt := e.questions[questionType]
	questionColStart := pcm[questionType].gradeRateColStart + 2
	// 遍历某题型下的所有题目编号
	for _, pn := range pt.questionNumbers {
		var (
			totalScoreWithAllClass  float64 // 某题所有班级的总分数之和
			actualScoreWithAllClass float64 // 某题所有班级的实得分之和
		)
		col, _ := excelize.ColumnNumberToName(pn + 3) // 该题所在的列号
		// 计算该题目在每个班级的扣分总和
		for _, sheetName := range sheets[1:] {
			deductedPointSum := float64(0) // 所有学生在该题的扣分总和
			for i := 0; i < e.classSize[sheetName]; i++ {
				v, err := f.GetCellValue(sheetName, col+strconv.Itoa(i+2))
				if err != nil {
					logrus.Errorf("获取单元格值异常: %v", err)
				}
				vf, _ := strconv.ParseFloat(v, 64)
				deductedPointSum += float64(vf)
			}
			logrus.Debugf("%v 扣了 %v 分", sheetName, deductedPointSum)
			totalScoreWithClass := (float64(pt.questionScore[pn]) * float64(e.classSize[sheetName])) // 某题全班总分数
			actualScoreWithClass := (totalScoreWithClass - deductedPointSum)                         // 某题全班实得分

			totalScoreWithAllClass += totalScoreWithClass   // 某题全年级总分数
			actualScoreWithAllClass += actualScoreWithClass // 某题全年级实得分
		}

		rate := (actualScoreWithAllClass / totalScoreWithAllClass) * 100 // 得分率
		logrus.Printf("第 %v 题 %v 实际得分 %v，得分率为 %0.2f%%", pn, questionType, actualScoreWithAllClass, rate)

		resultScoreCol, _ := excelize.ColumnNumberToName(questionColStart)
		resultRateCol, _ := excelize.ColumnNumberToName(questionColStart + 1)
		fResult.SetCellStr("年级各题得分率", resultScoreCol+"3", fmt.Sprintf("%0.2f", actualScoreWithAllClass)) // 第 X 题实得分
		fResult.SetCellStr("年级各题得分率", resultRateCol+"3", fmt.Sprintf("%0.2f%%", rate))                   // 第 X 题得分率
		questionColStart += 2

		questionAllTotalScore += totalScoreWithAllClass   // 某题型下的全部题的全年级总分数
		questionAllActualScore += actualScoreWithAllClass // 某题型下的全部题的全年级实得分
	}

	calculationRate = (questionAllActualScore / questionAllTotalScore) * 100

	questionScoreCol, _ := excelize.ColumnNumberToName(pcm[questionType].gradeRateColStart)
	questionRateCol, _ := excelize.ColumnNumberToName(pcm[questionType].gradeRateColStart + 1)
	fResult.SetCellStr("年级各题得分率", questionScoreCol+"3", fmt.Sprintf("%0.2f", questionAllActualScore)) // 某题型实得分
	fResult.SetCellStr("年级各题得分率", questionRateCol+"3", fmt.Sprintf("%0.2f%%", calculationRate))       // 某题型总得分率
}

func main() {
	// 初始化日志
	logging.AddFlags(&logFlags)
	logFlags.LogOutput = "./stdout.log"
	pflag.Parse()
	if err := logging.LogrusInit(&logFlags); err != nil {
		logrus.Fatal("初始化日志失败", err)
	}

	fileOrginInfo := "info.xlsx"
	f, err := excelize.OpenFile(fileOrginInfo)
	if err != nil {
		logrus.Fatalf("打开 Excel 文件异常，原因: %v", err)
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			logrus.Errorln(err)
			return
		}
	}()

	// 打开需要填充结果的文件
	fileResult := "result_tmp.xlsx"
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

	e := NewExamination(f)
	logrus.Infof("共 %v 个班，每班人数分别为 %v", e.classTotal, e.classSize)

	initScoreAnalysisSheetCellNum(fResult, e.questionTypes)

	sheets := f.GetSheetList()

	// for i, sheetName := range sheets[1:2] {
	for i, sheetName := range sheets[1:] {
		for _, qt := range e.questionTypes {
			e.calculateOfQuestionTypeWithClass(qt, sheetName, i+3, f, fResult)
		}
	}

	for _, qt := range e.questionTypes {
		e.calculateOfQuestionTypeWithGrade(qt, sheets, f, fResult)
	}

	err = fResult.Save()
	if err != nil {
		logrus.Fatalf("保存文件异常，原因: %v", err)
	}
}
