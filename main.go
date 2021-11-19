package main

import (
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
)

func help(arg []string) {
	fmt.Println("Example:\n")
	fmt.Printf("\t%v fileName.xlsx\n\n", arg[0])
}

// 输出excel文件内容
func exportExcel(fileName string) {

	f, err := excelize.OpenFile(fileName)
	if err != nil {
		println(err.Error())
		return
	}

	for _, name := range f.GetSheetMap() {
		// 获取所有单元格
		rows, err := f.GetRows(name)
		if err != nil {
			fmt.Println(err)
			return
		}

		for _, row := range rows {
			for _, colCell := range row {
				// 将内容输出到标准输出
				fmt.Fprintf(os.Stdout, colCell)
			}
		}
	}
}

func existsFile(path string) bool {
	_, err := os.Stat(path)
	if err != nil {
		if os.IsExist(err) {
			return true
		}
		return false
	}
	return true
}

func main() {

	arg := os.Args

	if len(arg) != 2 {
		fmt.Println("Args Error")
		help(arg)
		os.Exit(1)
	}

	if arg[1] == "-help" || arg[1] == "-h" {
		help(arg)
		os.Exit(0)
	}

	if !existsFile(arg[1]) {
		fmt.Println("Excel File not Found")
		os.Exit(1)
	}

	exportExcel(arg[1])
}
