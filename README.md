# go-eexcel

go-eexcel implements encoding and decoding of XLSX like `encoding/json`

## Usage

```go
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
```

```go
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
```

## License

MIT

## Author

sago35 - <sago35@gmail.com>
