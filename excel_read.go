package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/golang/freetype/truetype"
	"github.com/wcharczuk/go-chart/v2"
)

func main() {
	// if file name wront need to back to the begining
readfile:
	var sourceName string
	fmt.Println("Please Enter File Name:(Default:OT.xlsx)")
	fmt.Scanln(&sourceName)
	if sourceName == "" {
		sourceName = "OT.xlsx"
	} else if !strings.Contains(sourceName, ".xlsx") {
		sourceName += ".xlsx"
	}
	path, err := os.Getwd()
	if err != nil {
		fmt.Println(err)
	}

	// formatted := fmt.Sprintf("%d/%d/%d", t.Month(), t.Day(), t.Year())
	f, err := excelize.OpenFile(sourceName)
	if err != nil {
		fmt.Println(err, "\nRetry to enter the correct file name.")
		goto readfile
	}

	// start counting time
	t := time.Now()
	fmt.Println("Start Parsing File From: \n" + path + "\\" + sourceName)
	rows, err := f.GetRows("Form1")
	var count int
	employee := make(map[string]float64) // pieChart Count
	if err == nil {
		fmt.Scanln("") // becuz goto function need to init param
		for i, row := range rows {
			if i == 0 || i == 1 {
				continue
			}
			var s []string
			for _, colCell := range row {
				s = append(s, colCell) // append the column to a ary called s
			}
			date, error := time.Parse("01-02-06", s[8]) // parse the xlsx string to time var
			// check if it is the right value including ot, last month, same year, without errors
			if s[9] != "OT時數" || date.Month() != t.Month()-1 || date.Year() != t.Year() || error != nil {
				m := i + 1                                // i startfrom 0 but xlsx start from 1
				err := f.SetRowVisible("Form1", m, false) // set the useless row unvisible
				if err != nil {
					fmt.Println(err)
				}
				count = m
			} else if s[9] == "OT時數" || date.Month() == t.Month()-1 || date.Year() == t.Year() || error == nil {
				employee[s[7]] = employee[s[7]] + 1
			}
		}
	}
	// stop counting time
	elapsed := time.Since(t)
	// draw the chart
	chartMaker(employee)

	// deal with create a new file
	var filename string
	filename = path + "\\OT時數表.xlsx"
	f.SaveAs(filename)

	fmt.Println("\n\nSUCCEEDED. \nTime Spent:", elapsed)
	fmt.Println("\nDocument Info:")
	fmt.Println(" ------------------------------------------------------------------------------")
	fmt.Println("\tRow Count:", count, "\t Filter name: OT時數", "\t Filter Month:", t.Month()-1)
	fmt.Println("\tNew File Path:", filename)
	fmt.Println("\tChart File Path:", path+"\\pieChart.png")
	fmt.Println(" ------------------------------------------------------------------------------")
	fmt.Println("Press Enter to exit.")
	fmt.Scanln()

}

func chartMaker(e map[string]float64) {
	// default font cannot support chinese need other font
	font := getZWFont()
	a := []chart.Value{}
	for key, value := range e { // Order not specified
		// fmt.Println(key, value)
		a = append(a, chart.Value{Label: key, Value: value})
	}
	// chart setting
	pie := chart.PieChart{
		Width:  512,
		Height: 512,
		Values: a,
		Font:   font,
	}
	f, _ := os.Create("pieChart.png")
	defer f.Close()
	pie.Render(chart.PNG, f)
}

func getZWFont() *truetype.Font {

	fontFile := "chinese.msyh.ttf"

	fontBytes, err := ioutil.ReadFile(fontFile)
	if err != nil {
		fmt.Println(err)
		return nil
	}
	font, err := truetype.Parse(fontBytes)
	if err != nil {
		fmt.Println(err)
		return nil
	}
	return font
}
