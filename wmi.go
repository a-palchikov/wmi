package wmi

import (
	"github.com/tianlin/com-and-go/v2"
	"unsafe"
)

var NilStr = com.BStr{}
var (
	ole32                = com.LoadDLL("ole32.dll")
	coInitializeSecurity = ole32.Func("CoInitializeSecurity")
	coSetProxyBlanket    = ole32.Func("CoSetProxyBlanket")
)

var (
	CLSID_WbemLocator        = com.NewGUID("{4590f811-1d3a-11d0-891f-00aa004b2e24}")
	IID_IWbemLocator         = com.NewGUID("{dc12a687-737f-11cf-884d-00aa004b2e24}")
	IID_IEnumWbemClassObject = com.NewGUID("{027947e1-d731-11ce-a357-000000000001}")
	IID_IWbemClassObject     = com.NewGUID("{dc12a681-737f-11cf-884d-00aa004b2e24}")
	IID_IWbemServices        = com.NewGUID("{9556dc99-828c-11cf-a37e-00aa003240c7}")
)

type (
	IWbemLocator struct {
		com.IUnknown
	}
	IWbemServices struct {
		com.IUnknown
	}
	IEnumWbemClassObject struct {
		com.IUnknown
	}
	IWbemClassObject struct {
		com.IUnknown
	}
	EnumWbemClassObject struct {
		enum   *IEnumWbemClassObject
		clsobj *IWbemClassObject
		err    error
	}
)

func (d *IWbemLocator) ConnectServerErr(resource string) (svc *IWbemServices, err error) {
	res := com.ToBStr(resource)
	defer res.Free()
	args := &struct {
		d                *IWbemLocator
		resource         com.BStr
		uid, pwd, locale com.BStr
		flags            int
		authority        com.BStr
		ctx              *uint32 // *IWbemContext, assumed nil here
		svc              **IWbemServices
	}{d, res, NilStr, NilStr, NilStr, 0, NilStr, nil, &svc}
	err = d.VTable[3].CallHR(unsafe.Pointer(args), 9)
	return
}

func (d *IWbemLocator) ConnectServer(resource string) (svc *IWbemServices) {
	var err error
	svc, err = d.ConnectServerErr(resource)
	if err != nil {
		panic(err)
	}
	return
}

func (d *IWbemServices) ExecQueryErr(queryL, query string, flags int) (enum EnumWbemClassObject, err error) {
	qL := com.ToBStr(queryL)
	defer qL.Free()
	q := com.ToBStr(query)
	defer q.Free()
	var enm *IEnumWbemClassObject
	args := &struct {
		d                    *IWbemServices
		queryLanguage, query com.BStr
		flags                int
		ctx                  *uint32 // IWbemContext, assumed nil
		enum                 **IEnumWbemClassObject
	}{d, qL, q, flags, nil, &enm}
	err = d.VTable[20].CallHR(unsafe.Pointer(args), 6)
	enum.enum = enm
	return
}

func (d *IWbemServices) ExecQuery(queryL, query string, flags int) (enum EnumWbemClassObject) {
	var err error
	enum, err = d.ExecQueryErr(queryL, query, flags)
	if err != nil {
		panic(err)
	}
	return
}

func (d *IEnumWbemClassObject) NextErr(timeout int, count uint32) (objs *IWbemClassObject, n uint32, err error) {
	args := &struct {
		d       *IEnumWbemClassObject
		timeout int
		count   uint32
		objs    **IWbemClassObject
		ret     *uint32
	}{d, timeout, count, &objs, &n}
	err = d.VTable[4].CallHR(unsafe.Pointer(args), 5)
	return
}

func (d *EnumWbemClassObject) Release() {
	if d.clsobj != nil {
		d.clsobj.Release()
		d.clsobj = nil
	}
	d.enum.Release()
}

func (d *EnumWbemClassObject) Next(timeout int, count uint32) (ok bool) {
	if d.clsobj != nil {
		d.clsobj.Release()
		d.clsobj = nil
	}
	var n uint32
	d.clsobj, n, d.err = d.enum.NextErr(timeout, count)
	if n == 0 { // No more items
		d.err = nil
		if d.clsobj != nil {
			d.clsobj.Release()
			d.clsobj = nil
		}
		return false
	}
	ok = n != 0
	return
}

func (d *EnumWbemClassObject) Get(name string) (ret interface{}) {
	return d.clsobj.Get(name, 0)
}

func (d *EnumWbemClassObject) Err() error {
	return d.err
}

func (d *IWbemClassObject) GetErr(name string, flags int) (ret interface{}, err error) {
	var res com.Variant
	prop := com.WideString(name)
	args := &struct {
		d        *IWbemClassObject
		name     *uint16 // LPCWSTR
		flags    int
		val      *com.Variant
		pType    *uint32
		plFlavor *uint32
	}{d, prop, flags, &res, nil, nil}
	err = d.VTable[4].CallHR(unsafe.Pointer(args), 6)
	return res.ToInterface(), err
}

func (d *IWbemClassObject) Get(name string, flags int) (ret interface{}) {
	var err error
	ret, err = d.GetErr(name, flags)
	if err != nil {
		panic(err)
	}
	return
}

func CoInitializeSecurity(secDesc *uint32, numAuthSvc int, authSvc *uint32,
	reserved1 *uint32, authLevel, impLevel uint32,
	authList *uint32, caps uint32, reserved3 *uint32) error {
	args := &struct {
		secDesc             *uint32
		numAuthSvc          int
		authSvc             *uint32
		reserved1           *uint32
		authLevel, impLevel uint32
		authList            *uint32
		caps                uint32
		reserved3           *uint32
	}{secDesc, numAuthSvc, authSvc, reserved1, authLevel,
		impLevel, authList, caps, reserved3}
	return coInitializeSecurity.CallHR(unsafe.Pointer(args), 9)
}

func CoSetProxyBlanket(proxy *com.IUnknown, authnSvc, authzSvc uint32, serverPrincName com.BStr,
	authnLevel, impLevel uint32, authInfo *uint32, caps uint32) error {
	args := &struct {
		proxy               *com.IUnknown
		authnSvc, authzSvc  uint32
		serverPrincName     com.BStr
		authLevel, impLevel uint32
		authInfo            *uint32
		caps                uint32
	}{proxy, authnSvc, authzSvc, serverPrincName, authnLevel, impLevel, authInfo, caps}
	return coSetProxyBlanket.CallHR(unsafe.Pointer(args), 8)
}

const (
	CLSCTX_INPROC_SERVER = 0x1
)

const (
	RPC_C_AUTHN_LEVEL_DEFAULT = iota
	RPC_C_AUTHN_LEVEL_NONE
	RPC_C_AUTHN_LEVEL_CONNECT
	RPC_C_AUTHN_LEVEL_CALL
	RPC_C_AUTHN_LEVEL_PKT
	RPC_C_AUTHN_LEVEL_PKT_INTEGRITY
	RPC_C_AUTHN_LEVEL_PKT_PRIVACY
)

const (
	RPC_C_AUTHN_NONE          = 0
	RPC_C_AUTHN_DCE_PRIVATE   = 1
	RPC_C_AUTHN_DCE_PUBLIC    = 2
	RPC_C_AUTHN_DEC_PUBLIC    = 4
	RPC_C_AUTHN_GSS_NEGOTIATE = 9
	RPC_C_AUTHN_WINNT         = 10
	RPC_C_AUTHN_GSS_SCHANNEL  = 14
	RPC_C_AUTHN_GSS_KERBEROS  = 16
	RPC_C_AUTHN_DPA           = 17
	RPC_C_AUTHN_MSN           = 18
	RPC_C_AUTHN_DIGEST        = 21
	RPC_C_AUTHN_KERNEL        = 20
	RPC_C_AUTHN_NEGO_EXTENDER = 30
	RPC_C_AUTHN_PKU2U         = 31
	RPC_C_AUTHN_MQ            = 100
	RPC_C_AUTHN_DEFAULT       = 0xffffffff
)

const (
	RPC_C_IMP_LEVEL_DEFAULT     = 0
	RPC_C_IMP_LEVEL_ANONYMOUS   = 1
	RPC_C_IMP_LEVEL_IDENTIFY    = 2
	RPC_C_IMP_LEVEL_IMPERSONATE = 3
	RPC_C_IMP_LEVEL_DELEGATE    = 4
)

const (
	RPC_C_AUTHZ_NONE    = 0
	RPC_C_AUTHZ_NAME    = 1
	RPC_C_AUTHZ_DCE     = 2
	RPC_C_AUTHZ_DEFAULT = 0xffffffff
)

const (
	EOAC_NONE              = 0
	EOAC_MUTUAL_AUTH       = 0x1
	EOAC_STATIC_CLOAKING   = 0x20
	EOAC_DYNAMIC_CLOAKING  = 0x40
	EOAC_ANY_AUTHORITY     = 0x80
	EOAC_MAKE_FULLSIC      = 0x100
	EOAC_DEFAULT           = 0x800
	EOAC_SECURE_REFS       = 0x2
	EOAC_ACCESS_CONTROL    = 0x4
	EOAC_APPID             = 0x8
	EOAC_DYNAMIC           = 0x10
	EOAC_REQUIRE_FULLSIC   = 0x200
	EOAC_AUTO_IMPERSONATE  = 0x400
	EOAC_NO_CUSTOM_MARSHAL = 0x2000
	EOAC_DISABLE_AAA       = 0x1000
)

const (
	WBEM_FLAG_RETURN_IMMEDIATELY     = 0x10
	WBEM_FLAG_RETURN_WBEM_COMPLETE   = 0
	WBEM_FLAG_BIDIRECTIONAL          = 0
	WBEM_FLAG_FORWARD_ONLY           = 0x20
	WBEM_FLAG_NO_ERROR_OBJECT        = 0x40
	WBEM_FLAG_RETURN_ERROR_OBJECT    = 0
	WBEM_FLAG_SEND_STATUS            = 0x80
	WBEM_FLAG_DONT_SEND_STATUS       = 0
	WBEM_FLAG_ENSURE_LOCATABLE       = 0x100
	WBEM_FLAG_DIRECT_READ            = 0x200
	WBEM_FLAG_SEND_ONLY_SELECTED     = 0
	WBEM_RETURN_WHEN_COMPLETE        = 0
	WBEM_RETURN_IMMEDIATELY          = 0x10
	WBEM_MASK_RESERVED_FLAGS         = 0x1f000
	WBEM_FLAG_USE_AMENDED_QUALIFIERS = 0x20000
	WBEM_FLAG_STRONG_VALIDATION      = 0x100000
)

const (
	WBEM_NO_WAIT  = 0
	WBEM_INFINITE = -1
)
