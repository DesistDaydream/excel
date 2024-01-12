package main

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/sirupsen/logrus"
	"github.com/xuri/excelize/v2"
)

type problemCellNum struct {
	classRateColStart int
	gradeRateColStart int
}

var pcm = map[string]problemCellNum{
	"计算题": problemCellNum{
		classRateColStart: 7,
		gradeRateColStart: 4,
	},
	"填空题": problemCellNum{
		classRateColStart: 13,
		gradeRateColStart: 10,
	},
	"判断题": problemCellNum{
		classRateColStart: 39,
		gradeRateColStart: 36,
	},
	"选择题": problemCellNum{
		classRateColStart: 49,
		gradeRateColStart: 46,
	},
	"解决问题": problemCellNum{
		classRateColStart: 59,
		gradeRateColStart: 56,
	},
}

type problem struct {
	problems   map[string]problemType
	classTotal int            // 班级总数
	classSize  map[string]int // 每个班级的人数
}

type problemType struct {
	// typeName       string
	problemNumbers []int       // 题目号
	problemScore   map[int]int // 题目分值
}

func NewProblems(f *excelize.File) *problem {
	ps := make(map[string]problemType)
	rows, err := f.GetRows("info")
	if err != nil {
		logrus.Errorf("get cols failed, err:%v", err)
	}
	// fmt.Println(rows)
	for i, row := range rows {
		var q problemType

		if i == 0 {
			continue
		}

		if row == nil {
			break
		}

		// 获取题号
		qNums, _ := f.GetCellValue("info", "B"+strconv.Itoa(i+1))
		qNum := strings.Split(qNums, " ")
		for _, n := range qNum {
			ni, err := strconv.Atoi(n)
			if err != nil {
				logrus.Errorf("strconv.Atoi failed, err:%v", err)
			}
			q.problemNumbers = append(q.problemNumbers, ni)
		}

		// 获取题号对应的分值
		qScores, _ := f.GetCellValue("info", "C"+strconv.Itoa(i+1))
		qScore := strings.Split(qScores, " ")
		q.problemScore = make(map[int]int)
		for i, s := range qScore {
			si, _ := strconv.Atoi(s)
			q.problemScore[q.problemNumbers[i]] = si
		}

		// 获取题型名称
		// q.typeName = row[0]

		ps[row[0]] = q
	}

	fmt.Println(ps)

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

	return &problem{
		problems:   ps,
		classTotal: c,
		classSize:  css,
	}
}

// 计算得分率
func (p *problem) calculateScoreRate(sheetName string, f *excelize.File) float64 {
	// 计算得分率
	for problemType, problem := range p.problems {
		logrus.Infof("开始计算 %v 得分率", problemType)
		for _, pn := range problem.problemNumbers {
			// fmt.Println(qn)
			col, _ := excelize.ColumnNumberToName(pn + 3)
			sum := float64(0)
			for i := 0; i < p.classSize[sheetName]; i++ {
				v, err := f.GetCellValue(sheetName, col+strconv.Itoa(i+2))
				if err != nil {
					logrus.Errorf("获取单元格值异常: %v", err)
				}
				vf, _ := strconv.ParseFloat(v, 64)
				sum += float64(vf)
			}
			rate := (float64(problem.problemScore[pn])*float64(p.classSize[sheetName]) - sum) / (float64(problem.problemScore[pn]) * float64(p.classSize[sheetName])) * 100
			fmt.Printf("第 %v 题 %v 总得分 %v，得分率为 %0.2f%%\n", pn, problemType, (float64(problem.problemScore[pn])*float64(p.classSize[sheetName]) - sum), rate)
		}
	}

	return 0
}

// 计算班级各题型的得分情况
func (p *problem) calculateOfProblemTypeWithClass(problemType string, sheetName string, classRowNum int, f *excelize.File, fResult *excelize.File) {
	var (
		problemAllTotalScore  float64
		problemAllActualScore float64
		calculationRate       float64
	)

	problemColStart := pcm[problemType].classRateColStart + 2
	pt := p.problems[problemType]
	for _, pn := range pt.problemNumbers {
		// fmt.Println(qn)
		col, _ := excelize.ColumnNumberToName(pn + 3)
		sum := float64(0)
		for i := 0; i < p.classSize[sheetName]; i++ {
			v, err := f.GetCellValue(sheetName, col+strconv.Itoa(i+2))
			if err != nil {
				logrus.Errorf("获取单元格值异常: %v", err)
			}
			vf, _ := strconv.ParseFloat(v, 64)
			sum += float64(vf)
		}

		totalScore := (float64(pt.problemScore[pn]) * float64(p.classSize[sheetName]))      // 总分数
		actualScore := (float64(pt.problemScore[pn])*float64(p.classSize[sheetName]) - sum) // 实得分
		rate := (actualScore / totalScore) * 100                                            // 得分率

		fmt.Printf("第 %v 题 %v 总得分 %v，得分率为 %0.2f%%\n", pn, problemType, actualScore, rate)

		resultScoreCol, _ := excelize.ColumnNumberToName(problemColStart)
		fResult.SetCellStr("班级各题得分率", resultScoreCol+strconv.Itoa(classRowNum), fmt.Sprintf("%0.2f", actualScore)) // 第 X 题实得分
		resultRateCol, _ := excelize.ColumnNumberToName(problemColStart + 1)
		fResult.SetCellStr("班级各题得分率", resultRateCol+strconv.Itoa(classRowNum), fmt.Sprintf("%0.2f%%", rate)) // 第 X 题得分率
		problemColStart += 2

		problemAllTotalScore += totalScore
		problemAllActualScore += actualScore
	}

	problemScoreCol, _ := excelize.ColumnNumberToName(pcm[problemType].classRateColStart)
	problemRateCol, _ := excelize.ColumnNumberToName(pcm[problemType].classRateColStart + 1)
	calculationRate = (problemAllActualScore / problemAllTotalScore) * 100
	fResult.SetCellStr("班级各题得分率", problemScoreCol+strconv.Itoa(classRowNum), fmt.Sprintf("%0.2f", problemAllActualScore)) // 某题型实得分
	fResult.SetCellStr("班级各题得分率", problemRateCol+strconv.Itoa(classRowNum), fmt.Sprintf("%0.2f%%", calculationRate))      // 某题型总得分率
}

// 计算年级各题型的得分情况
func (p *problem) calculateOfProblemTypeWithGrade(problemType string, sheets []string, f *excelize.File, fResult *excelize.File) {
	var (
		problemAllTotalScore  float64
		problemAllActualScore float64
		calculationRate       float64
	)

	pt := p.problems[problemType]
	problemColStart := pcm[problemType].gradeRateColStart + 2
	// 遍历某题型下的所有题目编号
	for _, pn := range pt.problemNumbers {
		var (
			totalScoreWithAllClass  float64 // 某题所有班级的总分数之和
			actualScoreWithAllClass float64 // 某题所有班级的实得分之和
		)
		col, _ := excelize.ColumnNumberToName(pn + 3) // 该题所在的列号
		// 计算该题目在每个班级的扣分总和
		for _, sheetName := range sheets[1:] {
			deductedPointSum := float64(0) // 所有学生在该题的扣分总和
			for i := 0; i < p.classSize[sheetName]; i++ {
				v, err := f.GetCellValue(sheetName, col+strconv.Itoa(i+2))
				if err != nil {
					logrus.Errorf("获取单元格值异常: %v", err)
				}
				vf, _ := strconv.ParseFloat(v, 64)
				deductedPointSum += float64(vf)
			}
			// fmt.Printf("%v 扣了 %v 分\n", sheetName, deductedPointSum)
			totalScoreWithClass := (float64(pt.problemScore[pn]) * float64(p.classSize[sheetName])) // 某题全班总分数
			actualScoreWithClass := (totalScoreWithClass - deductedPointSum)                        // 某题全班实得分

			totalScoreWithAllClass += totalScoreWithClass   // 某题全年级总分数
			actualScoreWithAllClass += actualScoreWithClass // 某题全年级实得分
		}

		rate := (actualScoreWithAllClass / totalScoreWithAllClass) * 100 // 得分率
		fmt.Printf("第 %v 题 %v 实际得分 %v，得分率为 %0.2f%%\n", pn, problemType, actualScoreWithAllClass, rate)

		resultScoreCol, _ := excelize.ColumnNumberToName(problemColStart)
		resultRateCol, _ := excelize.ColumnNumberToName(problemColStart + 1)
		fResult.SetCellStr("年级各题得分率", resultScoreCol+"3", fmt.Sprintf("%0.2f", actualScoreWithAllClass)) // 第 X 题实得分
		fResult.SetCellStr("年级各题得分率", resultRateCol+"3", fmt.Sprintf("%0.2f%%", rate))                   // 第 X 题得分率
		problemColStart += 2

		problemAllTotalScore += totalScoreWithAllClass   // 某题型下的全部题的全年级总分数
		problemAllActualScore += actualScoreWithAllClass // 某题型下的全部题的全年级实得分
	}

	calculationRate = (problemAllActualScore / problemAllTotalScore) * 100

	problemScoreCol, _ := excelize.ColumnNumberToName(pcm[problemType].gradeRateColStart)
	problemRateCol, _ := excelize.ColumnNumberToName(pcm[problemType].gradeRateColStart + 1)
	fResult.SetCellStr("年级各题得分率", problemScoreCol+"3", fmt.Sprintf("%0.2f", problemAllActualScore)) // 某题型实得分
	fResult.SetCellStr("年级各题得分率", problemRateCol+"3", fmt.Sprintf("%0.2f%%", calculationRate))      // 某题型总得分率
}

func main() {
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

	p := NewProblems(f)
	logrus.Infof("共 %v 个班，每班人数分别为 %v", p.classTotal, p.classSize)

	fmt.Println(p)

	sheets := f.GetSheetList()

	// for i, sheetName := range sheets[1:2] {
	for i, sheetName := range sheets[1:] {
		// 计算每道题的得分率
		// q.calculateScoreRate(sheetName, f)
		// 计算计算题的得分率
		p.calculateOfProblemTypeWithClass("计算题", sheetName, i+3, f, fResult)
		p.calculateOfProblemTypeWithClass("填空题", sheetName, i+3, f, fResult)
		p.calculateOfProblemTypeWithClass("判断题", sheetName, i+3, f, fResult)
		p.calculateOfProblemTypeWithClass("选择题", sheetName, i+3, f, fResult)
		p.calculateOfProblemTypeWithClass("解决问题", sheetName, i+3, f, fResult)
	}

	p.calculateOfProblemTypeWithGrade("计算题", sheets, f, fResult)
	p.calculateOfProblemTypeWithGrade("填空题", sheets, f, fResult)
	p.calculateOfProblemTypeWithGrade("判断题", sheets, f, fResult)
	p.calculateOfProblemTypeWithGrade("选择题", sheets, f, fResult)
	p.calculateOfProblemTypeWithGrade("解决问题", sheets, f, fResult)

	err = fResult.Save()
	if err != nil {
		logrus.Fatalf("保存文件异常，原因: %v", err)
	}

}