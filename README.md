wmi
===

wmi.go implements a wrapper to Windows Management Instrumentation API. Currently it implements the
bits to be able to run WMI queries.

The wrapper is not based on IDispatch which makes it somewhat more performant than automation-based access.
The wrapper itself is based on the awesome [com-and-go](https://code.google.com/p/com-and-go) package.

## Example
This is a barebones application to query Windows services:
```
package main

import (
	"github.com/tianlin/com-and-go/v2"
	"github.com/a-palchikov/wmi"
	"fmt"
	"log"
	"unsafe"
)

func main() {
	com.CoInitializeEx(nil, 0)
	if err := wmi.CoInitializeSecurity(nil, -1, nil, nil,
		wmi.RPC_C_AUTHN_LEVEL_DEFAULT, wmi.RPC_C_IMP_LEVEL_IMPERSONATE, nil, wmi.EOAC_NONE, nil); err != nil {
		log.Fatalln("CoInitializeSecurity=", err)
	}

	var locator *wmi.IWbemLocator
	err := com.CoCreateInstance(wmi.CLSID_WbemLocator,
		nil, wmi.CLSCTX_INPROC_SERVER, wmi.IID_IWbemLocator, unsafe.Pointer(&locator))
	if err != nil {
		log.Fatalln("CoCreateInstance=", err)
	}
	defer locator.Release()

	svc := locator.ConnectServer("root\\CIMV2")
	defer svc.Release()

	err = wmi.CoSetProxyBlanket(&svc.IUnknown,
		wmi.RPC_C_AUTHN_WINNT, wmi.RPC_C_AUTHZ_NONE,
		com.BStr{},
		wmi.RPC_C_AUTHN_LEVEL_CALL, wmi.RPC_C_IMP_LEVEL_IMPERSONATE,
		nil,
		wmi.EOAC_NONE)
	if err != nil {
		log.Fatalln("CoSetProxyBlanket=", err)
	}

	enum := svc.ExecQuery("WQL", "select * from win32_service",
		wmi.WBEM_FLAG_FORWARD_ONLY|wmi.WBEM_FLAG_RETURN_IMMEDIATELY)
	defer enum.Release()

	for enum.Next(wmi.WBEM_INFINITE, 1) {
		fmt.Println("Service.Path=", enum.Get("PathName").(string))
	}
	if enum.Err() != nil {
		panic(enum.Err())
	}
}
```
