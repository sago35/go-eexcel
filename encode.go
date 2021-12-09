package eexcel

import (
	"bytes"
	"fmt"
	"reflect"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var (
	DefaultSheetName = "defines"
)

func Marshal(v interface{}) ([]byte, error) {
	sh := DefaultSheetName
	xlsx := excelize.NewFile()

	xlsx.NewSheet(sh)
	xlsx.SetCellStr(sh, "A1", "key")
	xlsx.SetCellStr(sh, "B1", "value")

	rv := reflect.ValueOf(v)
	rt := reflect.TypeOf(v)

	row := 2
	for i := 0; i < rv.Type().NumField(); i++ {
		k := rt.Field(i)
		v := rv.Field(i)

		key := k.Name
		if strings.ToLower(key) == key {
			// skip private key-value
			continue
		}

		tag := k.Tag.Get("eexcel")
		if tag != "" {
			key = tag
		}
		xlsx.SetCellStr(sh, fmt.Sprintf("A%d", row), key)

		switch v.Kind() {
		case reflect.String:
			xlsx.SetCellStr(sh, fmt.Sprintf("B%d", row), v.String())
		case reflect.Int:
			xlsx.SetCellInt(sh, fmt.Sprintf("B%d", row), int(v.Int()))
		default:
			xlsx.SetCellStr(sh, fmt.Sprintf("B%d", row), v.Kind().String())
		}

		row++
	}

	buf := bytes.Buffer{}
	xlsx.Write(&buf)
	return buf.Bytes(), nil
}

func Unmarshal(data []byte, v interface{}) error {
	sh := DefaultSheetName
	xlsx, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		return err
	}

	rv := reflect.ValueOf(v)
	rt := reflect.TypeOf(v)
	if rt.Kind() == reflect.Ptr {
		rv = rv.Elem()
		rt = rt.Elem()
	}

	row := 2
	for {
		key := xlsx.GetCellValue(sh, fmt.Sprintf("A%d", row))
		if key == "" {
			break
		}

		for i := 0; i < rv.Type().NumField(); i++ {
			k := rt.Field(i)
			v := rv.Field(i)

			kk := k.Name
			tag := k.Tag.Get("eexcel")
			if tag != "" {
				kk = tag
			}
			if kk != key {
				continue
			}

			val := xlsx.GetCellValue(sh, fmt.Sprintf("B%d", row))
			switch v.Kind() {
			case reflect.String:
				v.SetString(val)
			case reflect.Int:
				vv, err := strconv.ParseInt(val, 0, 0)
				if err != nil {
					return err
				}
				v.SetInt(vv)
			default:
			}
			break
		}

		row++
	}

	return nil
}
