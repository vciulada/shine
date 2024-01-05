package shine

import (
	"fmt"
	"log"
	"reflect"
	"strconv"
	"time"
)

type DataProvider interface {
	GetData() [][]string
}

type marshel struct {
	data     [][]string
	rowIndex int
}

func newMarchel(provider DataProvider) marshel {
	result := marshel{}
	result.data = provider.GetData()
	return result
}

func (m marshel) unMarshal(field interface{}) {
	rv := reflect.ValueOf(field)
	m.changerv(rv)
}

func (m marshel) changerv(rv reflect.Value) {
	if rv.Kind() == reflect.Ptr {
		rv = rv.Elem()
	}

	if rv.Kind() == reflect.Struct {
		m.changeStruct(rv)
	}
	if rv.Kind() == reflect.Slice {
		m.changeSlice(rv)
	}
}

func (m marshel) changeStruct(rv reflect.Value) {
	if !rv.CanAddr() {
		return
	}
	for i := 0; i < rv.NumField(); i++ {
		field := rv.Field(i)
		tag := rv.Type().Field(i).Tag.Get("excel")
		if tag == "" {
			//skip if excel tag is not provided
			continue
		}
		index, err := strconv.Atoi(tag)
		if err != nil {
			log.Fatal("excel tag is not a number")
		}
		if len(m.data[m.rowIndex]) <= index {
			//skip if data cell is empty
			continue
		}
		switch field.Kind() {
		case reflect.String:
			field.SetString(m.data[m.rowIndex][index])
		case reflect.Uint8:
			intValue, err := strconv.ParseUint(m.data[m.rowIndex][index], 10, 8)
			if err != nil {
				log.Fatal("cell is not an int value")
			}
			field.SetUint(uint64(intValue))
		case reflect.Uint16:
			intValue, err := strconv.ParseUint(m.data[m.rowIndex][index], 10, 16)
			if err != nil {
				log.Fatal("cell is not an int value")
			}
			field.SetUint(uint64(intValue))
		case reflect.Uint32:
			intValue, err := strconv.ParseUint(m.data[m.rowIndex][index], 10, 32)
			if err != nil {
				log.Fatal("cell is not an int value")
			}
			field.SetUint(uint64(intValue))
		case reflect.Uint64:
			intValue, err := strconv.ParseUint(m.data[m.rowIndex][index], 10, 64)
			if err != nil {
				log.Fatal("cell is not an int value")
			}
			field.SetUint(uint64(intValue))
		case reflect.Uint:
			intValue, err := strconv.ParseUint(m.data[m.rowIndex][index], 10, 64)
			if err != nil {
				log.Fatal("cell is not an int value")
			}
			field.SetUint(uint64(intValue))
		case reflect.Int8:
			intValue, err := strconv.ParseInt(m.data[m.rowIndex][index], 10, 8)
			if err != nil {
				log.Fatal("cell is not an int value")
			}
			field.SetInt(int64(intValue))
		case reflect.Int16:
			intValue, err := strconv.ParseInt(m.data[m.rowIndex][index], 10, 16)
			if err != nil {
				log.Fatal("cell is not an int value")
			}
			field.SetInt(int64(intValue))
		case reflect.Int32:
			intValue, err := strconv.ParseInt(m.data[m.rowIndex][index], 10, 32)
			if err != nil {
				log.Fatal("cell is not an int value")
			}
			field.SetInt(int64(intValue))
		case reflect.Int64:
			intValue, err := strconv.ParseInt(m.data[m.rowIndex][index], 10, 64)
			if err != nil {
				log.Fatal("cell is not an int value")
			}
			field.SetInt(int64(intValue))
		case reflect.Int:
			intValue, err := strconv.Atoi(m.data[m.rowIndex][index])
			if err != nil {
				log.Fatal("cell is not an int value")
			}
			field.SetInt(int64(intValue))
		case reflect.Float64:
			intValue, err := strconv.ParseFloat(m.data[m.rowIndex][index], 64)
			if err != nil {
				log.Fatal("cell is not a float64")
			}
			field.SetFloat(intValue)
		case reflect.Float32:
			intValue, err := strconv.ParseFloat(m.data[m.rowIndex][index], 32)
			if err != nil {
				log.Fatal("cell is not a float32")
			}
			field.SetFloat(intValue)
		case reflect.Struct:
			switch field.Type().String() {
			case "time.Time":
				date, err := time.Parse("01-02-06", m.data[m.rowIndex][index])
				if err != nil {
					log.Fatalf("cannot parse %s to date on line %d:%s", date, m.rowIndex, err)
				}
				field.Set(reflect.ValueOf(date))
			default:
				fmt.Println("Unsupported struct")
			}

		default:
			fmt.Println(field.Kind())
			fmt.Println("unknown field")
		}
	}
}

func (m *marshel) changeSlice(rv reflect.Value) {
	ln := rv.Len()
	if ln == 0 {
		for i := 0; i < len(m.data); i++ {
			var elem reflect.Value

			typ := rv.Type().Elem()
			if typ.Kind() == reflect.Ptr {
				elem = reflect.New(typ.Elem())
			}
			if typ.Kind() == reflect.Struct {
				elem = reflect.New(typ).Elem()
			}

			rv.Set(reflect.Append(rv, elem))
		}
	}

	ln = rv.Len()
	for i := 0; i < ln; i++ {
		m.rowIndex = i
		m.changerv(rv.Index(i))
	}
}

func UnMarshal(data DataProvider, field interface{}) {
	m := newMarchel(data)
	m.unMarshal(field)
}
