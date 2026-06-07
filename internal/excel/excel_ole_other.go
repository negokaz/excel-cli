//go:build !windows

package excel

import "fmt"

func NewExcelOle(_ string) (Excel, func(), error) {
	return nil, func() {}, fmt.Errorf("OLE automation is only supported on Windows")
}

func NewExcelOleOpen(_ string) (Excel, func(), error) {
	return nil, func() {}, fmt.Errorf("OLE automation is only supported on Windows")
}
