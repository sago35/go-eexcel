# go-eexcel

go-eexcel implements encoding and decoding of XLSX like `encoding/json`

## Usage

```go
func ExampleEexcel() {
	type st struct {
		Name   string `eexcel:"name"`
		Number int    `eexcel:"number"`
	}

	// marshal
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
	fmt.Printf("\n")

	// unmarshal
	output := st{}
	Unmarshal(b, &output)

	fmt.Printf("Name   : %q\n", output.Name)
	fmt.Printf("Number : %d\n", output.Number)

	// Output:
	// key	value
	// name	go-eexcel
	// number	123456
	//
	// Name   : "go-eexcel"
	// Number : 123456
}
```

## License

MIT

## Author

sago35 - <sago35@gmail.com>
