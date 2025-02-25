package msi

import (
	"syscall"
	"unsafe"
)

var (
	msi = syscall.NewLazyDLL("msi.dll")

	msiOpenDatabase     = msi.NewProc("MsiOpenDatabaseW")
	msiCloseHandle      = msi.NewProc("MsiCloseHandle")
	msiDatabaseOpenView = msi.NewProc("MsiDatabaseOpenViewW")
	msiViewExecute      = msi.NewProc("MsiViewExecute")
	msiViewFetch        = msi.NewProc("MsiViewFetch")
	msiRecordGetString  = msi.NewProc("MsiRecordGetStringW")
)

func OpenDatabase(path *uint16, persist *uint16, handle *syscall.Handle) error {
	r, _, _ := msiOpenDatabase.Call(
		uintptr(unsafe.Pointer(path)),
		uintptr(unsafe.Pointer(persist)),
		uintptr(unsafe.Pointer(handle)))
	if r != 0 {
		return syscall.Errno(r)
	}
	return nil
}

func CloseHandle(handle syscall.Handle) error {
	r, _, _ := msiCloseHandle.Call(uintptr(handle))
	if r != 0 {
		return syscall.Errno(r)
	}
	return nil
}

func DatabaseOpenView(handle syscall.Handle, query *uint16, view *syscall.Handle) error {
	r, _, _ := msiDatabaseOpenView.Call(
		uintptr(handle),
		uintptr(unsafe.Pointer(query)),
		uintptr(unsafe.Pointer(view)))
	if r != 0 {
		return syscall.Errno(r)
	}
	return nil
}

func ViewExecute(view syscall.Handle, record uintptr) error {
	r, _, _ := msiViewExecute.Call(uintptr(view), record)
	if r != 0 {
		return syscall.Errno(r)
	}
	return nil
}

func ViewFetch(view syscall.Handle, record *syscall.Handle) error {
	r, _, _ := msiViewFetch.Call(
		uintptr(view),
		uintptr(unsafe.Pointer(record)))
	if r != 0 {
		return syscall.Errno(r)
	}
	return nil
}

func RecordGetString(record syscall.Handle, field uint32, buffer *uint16, bufLen *uint32) error {
	r, _, _ := msiRecordGetString.Call(
		uintptr(record),
		uintptr(field),
		uintptr(unsafe.Pointer(buffer)),
		uintptr(unsafe.Pointer(bufLen)))
	if r != 0 {
		return syscall.Errno(r)
	}
	return nil
}
