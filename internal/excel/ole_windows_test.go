//go:build windows

package excel

import (
	"testing"
	"unsafe"

	"github.com/go-ole/go-ole"
)

func rawSafeArrayPutVariant(sa *ole.SafeArray, dim1, dim2 int32, v *ole.VARIANT) error {
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

func rawSafeArrayGetVariant(sa *ole.SafeArray, dim1, dim2 int32) (ole.VARIANT, error) {
	indices := [2]int32{dim1, dim2}
	var v ole.VARIANT
	hr, _, _ := procSAGetElem.Call(
		uintptr(unsafe.Pointer(sa)),
		uintptr(unsafe.Pointer(&indices[0])),
		uintptr(unsafe.Pointer(&v)),
	)
	if hr != 0 {
		return v, ole.NewError(hr)
	}
	return v, nil
}

func TestSafeArray2DPutVariantMatchesRawSafeArrayDimensionOrder(t *testing.T) {
	t.Parallel()

	sa, err := safeArrayCreate2DVariant(2, 3)
	if err != nil {
		t.Fatalf("failed to create SafeArray: %v", err)
	}
	defer func() {
		if destroyErr := safeArrayDestroy2D(sa); destroyErr != nil {
			t.Fatalf("failed to destroy SafeArray: %v", destroyErr)
		}
	}()

	elem := valueToVariant("r1c3")
	err = safeArray2DPutVariant(sa, 1, 3, &elem)
	elem.Clear()
	if err != nil {
		t.Fatalf("safeArray2DPutVariant(1,3) failed: %v", err)
	}

	got, err := rawSafeArrayGetVariant(sa, 1, 3)
	if err != nil {
		t.Fatalf("raw SafeArrayGetElement failed: %v", err)
	}
	defer got.Clear()

	if got.ToString() != "r1c3" {
		t.Fatalf("expected r1c3 at row 1 col 3, got %q", got.ToString())
	}
}

func TestSafeArray2DGetVariantMatchesRawSafeArrayDimensionOrder(t *testing.T) {
	t.Parallel()

	sa, err := safeArrayCreate2DVariant(2, 3)
	if err != nil {
		t.Fatalf("failed to create SafeArray: %v", err)
	}
	defer func() {
		if destroyErr := safeArrayDestroy2D(sa); destroyErr != nil {
			t.Fatalf("failed to destroy SafeArray: %v", destroyErr)
		}
	}()

	elem := valueToVariant("r2c3")
	err = rawSafeArrayPutVariant(sa, 2, 3, &elem)
	elem.Clear()
	if err != nil {
		t.Fatalf("raw SafeArrayPutElement failed: %v", err)
	}

	got, err := safeArray2DGetVariant(sa, 2, 3)
	if err != nil {
		t.Fatalf("safeArray2DGetVariant(2,3) failed: %v", err)
	}
	defer got.Clear()

	if got.ToString() != "r2c3" {
		t.Fatalf("expected r2c3 at row 2 col 3, got %q", got.ToString())
	}
}
