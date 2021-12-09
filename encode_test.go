package eexcel

import (
	"bytes"
	"fmt"
	"io/ioutil"
	"strconv"
	"strings"
	"testing"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type testStruct struct {
	AAA string
	BBB int
	CCC string
	ddd string
	EEE string `eexcel:"eEe"`
}

func TestMarshal(t *testing.T) {
	input := testStruct{
		AAA: "aaa",
		BBB: 222,
		CCC: "ccc",
		ddd: "dDd",
		EEE: "eee",
	}

	sh := DefaultSheetName
	b, _ := Marshal(input)
	xlsx, err := excelize.OpenReader(bytes.NewReader(b))
	if err != nil {
		t.Fatal(err)
	}

	if g, e := xlsx.GetCellValue(sh, "A1"), "key"; g != e {
		t.Errorf("got %q want %q", g, e)
	}
	if g, e := xlsx.GetCellValue(sh, "B1"), "value"; g != e {
		t.Errorf("got %q want %q", g, e)
	}

	if g, e := xlsx.GetCellValue(sh, "A2"), "AAA"; g != e {
		t.Errorf("got %q want %q", g, e)
	}
	if g, e := xlsx.GetCellValue(sh, "B2"), "aaa"; g != e {
		t.Errorf("got %q want %q", g, e)
	}

	if g, e := xlsx.GetCellValue(sh, "A3"), "BBB"; g != e {
		t.Errorf("got %q want %q", g, e)
	}
	if g, e := getCellValueInt64(xlsx, sh, "B3"), int64(222); g != e {
		t.Errorf("got %d want %d", g, e)
	}

	if g, e := xlsx.GetCellValue(sh, "A4"), "CCC"; g != e {
		t.Errorf("got %q want %q", g, e)
	}
	if g, e := xlsx.GetCellValue(sh, "B4"), "ccc"; g != e {
		t.Errorf("got %q want %q", g, e)
	}

	if g, e := xlsx.GetCellValue(sh, "A5"), "eEe"; g != e {
		t.Errorf("got %q want %q", g, e)
	}
	if g, e := xlsx.GetCellValue(sh, "B5"), "eee"; g != e {
		t.Errorf("got %q want %q", g, e)
	}

	//err = xlsx.SaveAs("testdata/test.xlsx")
	//if err != nil {
	//	t.Fatal(err)
	//}
}

func getCellValueInt64(xlsx *excelize.File, sheet, axis string) int64 {
	v := xlsx.GetCellValue(sheet, axis)

	ret, err := strconv.ParseInt(v, 0, 0)
	if err != nil {
		fmt.Printf("err : %s\n", v)
		return 0
	}

	return ret
}

func TestUnarshal(t *testing.T) {
	input := testStruct{}

	b, err := ioutil.ReadFile("testdata/test.xlsx")
	if err != nil {
		t.Fatal(err)
	}

	err = Unmarshal(b, &input)
	if err != nil {
		t.Fatal(err)
	}

	if g, e := input.AAA, "aaa"; g != e {
		t.Errorf("got %q want %q", g, e)
	}

	if g, e := input.BBB, 222; g != e {
		t.Errorf("got %d want %d", g, e)
	}

	if g, e := input.CCC, "ccc"; g != e {
		t.Errorf("got %q want %q", g, e)
	}

	if g, e := input.EEE, "eee"; g != e {
		t.Errorf("got %q want %q", g, e)
	}
}

func ExampleMarshal() {
	type st struct {
		Name   string `eexcel:"name"`
		Number int    `eexcel:"number"`
	}

	input := st{
		Name:   "go-eexcel",
		Number: 123456,
	}
	b, _ := Marshal(input)

	if false {
		// Save to file
		ioutil.WriteFile("out.xlsx", b, 0666)
	}

	xlsx, _ := excelize.OpenReader(bytes.NewReader(b))
	rows := xlsx.GetRows(DefaultSheetName)

	for _, row := range rows {
		fmt.Printf("%s\n", strings.Join(row, "	"))
	}

	// Output:
	// key	value
	// name	go-eexcel
	// number	123456
}

func ExampleUnmarshal() {
	type testStruct struct {
		A string `eexcel:"AAA"`
		B int    `eexcel:"BBB"`
	}
	output := testStruct{}

	b, _ := ioutil.ReadFile("testdata/test.xlsx")
	Unmarshal(b, &output)

	fmt.Printf("A : %q\n", output.A)
	fmt.Printf("B : %d\n", output.B)

	// Output:
	// A : "aaa"
	// B : 222
}
