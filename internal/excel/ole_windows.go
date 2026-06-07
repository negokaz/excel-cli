//go:build windows

package excel

import (
	"fmt"
	"math"
	"strconv"
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
)

// --- SafeArray 2D operations ---
// These proc variables are used for 2D SafeArray get/put operations, which require
// passing an array of indices (one per dimension) to the Windows API. The go-ole
// package only exposes 1D wrappers, so we declare our own procs here.

var (
	modOleAut32   = syscall.NewLazyDLL("oleaut32.dll")
	procSAGetElem = modOleAut32.NewProc("SafeArrayGetElement")
	procSAPutElem = modOleAut32.NewProc("SafeArrayPutElement")
	procSACreate  = modOleAut32.NewProc("SafeArrayCreate")
	procSADestroy = modOleAut32.NewProc("SafeArrayDestroy")
	procSAGetLB   = modOleAut32.NewProc("SafeArrayGetLBound")
	procSAGetUB   = modOleAut32.NewProc("SafeArrayGetUBound")
)

// safeArray2DGetVariant gets a VARIANT element from a 2D SafeArray.
// dim1 is the first dimension index (rows, 1-based per Excel COM convention),
// dim2 is the second dimension index (cols, 1-based).
func safeArray2DGetVariant(sa *ole.SafeArray, dim1, dim2 int32) (v ole.VARIANT, err error) {
	indices := [2]int32{dim1, dim2}
	hr, _, _ := procSAGetElem.Call(
		uintptr(unsafe.Pointer(sa)),
		uintptr(unsafe.Pointer(&indices[0])),
		uintptr(unsafe.Pointer(&v)),
	)
	if hr != 0 {
		err = ole.NewError(hr)
	}
	return
}

// safeArray2DPutVariant puts a VARIANT element into a 2D SafeArray.
// dim1 is the first dimension (rows, 1-based), dim2 is the second (cols, 1-based).
func safeArray2DPutVariant(sa *ole.SafeArray, dim1, dim2 int32, v *ole.VARIANT) error {
	indices := [2]int32{dim1, dim2}
	hr, _, _ := procSAPutElem.Call(
		uintptr(unsafe.Pointer(sa)),
		uintptr(unsafe.Pointer(&indices[0])),
		uintptr(unsafe.Pointer(v)),
	)
	if hr != 0 {
		return ole.NewError(hr)
	}
	return nil
}

// safeArrayCreate2DVariant creates a 2D VT_VARIANT SafeArray with 1-based lower bounds.
// dim1Size is rows (first dimension), dim2Size is cols (second dimension).
func safeArrayCreate2DVariant(dim1Size, dim2Size int) (*ole.SafeArray, error) {
	bounds := [2]ole.SafeArrayBound{
		{Elements: uint32(dim1Size), LowerBound: 1},
		{Elements: uint32(dim2Size), LowerBound: 1},
	}
	sa, _, _ := procSACreate.Call(
		uintptr(ole.VT_VARIANT),
		uintptr(2),
		uintptr(unsafe.Pointer(&bounds[0])),
	)
	if sa == 0 {
		return nil, fmt.Errorf("SafeArrayCreate returned nil")
	}
	return (*ole.SafeArray)(unsafe.Pointer(sa)), nil
}

// safeArrayDestroy2D destroys a 2D SafeArray created by safeArrayCreate2DVariant.
func safeArrayDestroy2D(sa *ole.SafeArray) error {
	hr, _, _ := procSADestroy.Call(uintptr(unsafe.Pointer(sa)))
	if hr != 0 {
		return ole.NewError(hr)
	}
	return nil
}

// safeArrayGetLB returns the lower bound for a given dimension (1-based dim index).
func safeArrayGetLB(sa *ole.SafeArray, dim uint32) (int32, error) {
	var lb int32
	hr, _, _ := procSAGetLB.Call(
		uintptr(unsafe.Pointer(sa)),
		uintptr(dim),
		uintptr(unsafe.Pointer(&lb)),
	)
	if hr != 0 {
		return 0, ole.NewError(hr)
	}
	return lb, nil
}

// safeArrayGetUB returns the upper bound for a given dimension (1-based dim index).
func safeArrayGetUB(sa *ole.SafeArray, dim uint32) (int32, error) {
	var ub int32
	hr, _, _ := procSAGetUB.Call(
		uintptr(unsafe.Pointer(sa)),
		uintptr(dim),
		uintptr(unsafe.Pointer(&ub)),
	)
	if hr != 0 {
		return 0, ole.NewError(hr)
	}
	return ub, nil
}

// variantToFormulaString converts a VARIANT element from Range.Formula to a string.
// For formula cells the result is the formula string (e.g. "=SUM(A1:A10)").
// For non-formula cells, numeric values are rendered as decimals.
func variantToFormulaString(v *ole.VARIANT) string {
	return stringifyCellValue(variantToFormulaAny(v))
}

func variantToFormulaAny(v *ole.VARIANT) any {
	switch v.VT {
	case ole.VT_BSTR:
		return v.ToString()
	case ole.VT_EMPTY, ole.VT_NULL:
		return ""
	default:
		// Numeric non-formula cells (rare but possible in bulk Formula array)
		goVal := v.Value()
		switch val := goVal.(type) {
		case float64:
			if val == math.Trunc(val) {
				return int64(val)
			}
			return val
		case float32:
			f64 := float64(val)
			if f64 == math.Trunc(f64) {
				return int64(f64)
			}
			return f64
		case int8:
			return int64(val)
		case int16:
			return int64(val)
		case int32:
			return int64(val)
		case int64:
			return val
		case nil:
			return ""
		default:
			return goVal
		}
	}
}

// variantToValueString converts a VARIANT element from Range.Value/Value2 to a string.
// Strings and booleans are converted to their natural representation.
// Numeric values are rendered as shortest exact decimal. Errors use Excel error notation.
func variantToValueString(v *ole.VARIANT) string {
	return stringifyCellValue(variantToValueAny(v))
}

func variantRequiresDisplayedText(v *ole.VARIANT) bool {
	vt := v.VT &^ (ole.VT_ARRAY | ole.VT_BYREF)
	return vt == ole.VT_DATE || vt == ole.VT_CY
}

func variantToValueAny(v *ole.VARIANT) any {
	switch v.VT {
	case ole.VT_BSTR:
		return v.ToString()
	case ole.VT_BOOL:
		if (v.Val & 0xffff) != 0 {
			return true
		}
		return false
	case ole.VT_EMPTY, ole.VT_NULL:
		return ""
	case ole.VT_ERROR:
		return xlCVErrorToString(int32(v.Val))
	default:
		goVal := v.Value()
		switch val := goVal.(type) {
		case float64:
			if val == math.Trunc(val) {
				return int64(val)
			}
			return val
		case float32:
			f64 := float64(val)
			if f64 == math.Trunc(f64) {
				return int64(f64)
			}
			return f64
		case int8:
			return int64(val)
		case int16:
			return int64(val)
		case int32:
			return int64(val)
		case int64:
			return val
		case uint8:
			return uint64(val)
		case uint16:
			return uint64(val)
		case uint32:
			return uint64(val)
		case uint64:
			return val
		case nil:
			return ""
		default:
			return goVal
		}
	}
}

// float64ToString converts a float64 to its shortest string representation.
// Whole numbers within the safe integer range are rendered without a decimal point.
func float64ToString(f float64) string {
	if f == math.Trunc(f) && math.Abs(f) < 1e15 {
		return strconv.FormatInt(int64(f), 10)
	}
	return strconv.FormatFloat(f, 'f', -1, 64)
}

// xlCVErrorToString converts an Excel XlCVError value to its error notation string.
func xlCVErrorToString(code int32) string {
	switch code {
	case 2000:
		return "#NULL!"
	case 2007:
		return "#DIV/0!"
	case 2015:
		return "#VALUE!"
	case 2023:
		return "#REF!"
	case 2029:
		return "#NAME?"
	case 2036:
		return "#NUM!"
	case 2042:
		return "#N/A"
	default:
		return "#ERROR!"
	}
}

// valueToVariant converts a Go value to an OLE VARIANT for use in a VT_VARIANT SafeArray.
// The caller is responsible for calling v.Clear() after the VARIANT is no longer needed.
func valueToVariant(val any) ole.VARIANT {
	if val == nil {
		return ole.NewVariant(ole.VT_EMPTY, 0)
	}
	switch v := val.(type) {
	case string:
		bstr := ole.SysAllocString(v)
		return ole.NewVariant(ole.VT_BSTR, int64(uintptr(unsafe.Pointer(bstr))))
	case bool:
		if v {
			return ole.NewVariant(ole.VT_BOOL, -1) // VARIANT_TRUE = -1 (0xFFFF as int16)
		}
		return ole.NewVariant(ole.VT_BOOL, 0)
	case int:
		return ole.NewVariant(ole.VT_I8, int64(v))
	case int8:
		return ole.NewVariant(ole.VT_I1, int64(v))
	case int16:
		return ole.NewVariant(ole.VT_I2, int64(v))
	case int32:
		return ole.NewVariant(ole.VT_I4, int64(v))
	case int64:
		return ole.NewVariant(ole.VT_I8, v)
	case uint:
		return ole.NewVariant(ole.VT_UI8, int64(v))
	case uint8:
		return ole.NewVariant(ole.VT_UI1, int64(v))
	case uint16:
		return ole.NewVariant(ole.VT_UI2, int64(v))
	case uint32:
		return ole.NewVariant(ole.VT_UI4, int64(v))
	case uint64:
		return ole.NewVariant(ole.VT_UI8, int64(v))
	case float32:
		f64 := float64(v)
		return ole.NewVariant(ole.VT_R8, *(*int64)(unsafe.Pointer(&f64)))
	case float64:
		return ole.NewVariant(ole.VT_R8, *(*int64)(unsafe.Pointer(&v)))
	default:
		// Fallback: convert to string
		bstr := ole.SysAllocString(fmt.Sprintf("%v", v))
		return ole.NewVariant(ole.VT_BSTR, int64(uintptr(unsafe.Pointer(bstr))))
	}
}
