// WinHostScriptEngine.cpp: implementation of the CScriptEngine class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"
#include "WinScriptHost.h"
#include "ScriptEngine.h"
#include <atlsafe.h>

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

using namespace PaintsNow;

template <class Impl, class Base>
struct CComPtrInternalCreator {
	CComPtrInternalCreator(CComModule& m) : module(m) {}
	inline Base* operator () () {
		Base* ptrType;
		IClassFactory* classFactory;
		if (SUCCEEDED(module.GetClassObject(__uuidof(Impl), IID_IClassFactory, (void**)&classFactory))) {
			if (SUCCEEDED(classFactory->CreateInstance(nullptr, __uuidof(Base), (void**)&ptrType))) {
				classFactory->Release();
				return ptrType;
			}

			classFactory->Release();
		}

		return nullptr;
	}

	CComModule& module;
};

template <class T, class B> 
CComPtr<B> CoCreateInstanceInternal() {
	static CComPtrInternalCreator<T, B> creator(_Module);
	CComPtr<B> ptr;
	ptr.Attach(creator());
	return ptr;
}

CScriptEngine::CScriptEngine(IThread& threadApi, const String& lang) : IScript(threadApi), language(lang), defaultRequest(*this) {
	Reset();
}

IScriptControlPtr& CScriptEngine::GetScriptControl() {
	return engine;
}

CScriptEngine::~CScriptEngine() {
	// to grantee destruction order, release them manually
	engine.Release();
	globalObject.Release();
}

void CScriptEngine::Reset() {
	IScript::Reset();

	if (engine) {
		engine.Release();
	}

	if (globalObject) {
		globalObject.Release();
	}

	proxyPool.clear();

	HRESULT engineResult = engine.CreateInstance(__uuidof(ScriptControl));
	CComBSTR str = (LPCWSTR)Utf8ToSystem(language).data();
	engine->PutLanguage((BSTR)str);

	globalObject = CoCreateInstanceInternal<ObjectProxy, IObjectProxy>();
	globalObject->SetHost((LONGLONG)this);
	engine->AddObject(L"window", globalObject, VARIANT_TRUE);

	if (language == "JScript") {
		LPCWSTR code = L"function _CallFunction(func) { return func.apply(null, arguments); }";
		engine->AddCode(code);
	} else if (language == "VBScript") {
		// no extra requirements.
	} else {
		fprintf(stderr, "+ WinScriptHost could not load %s.\n", language.c_str());
		return;
	}

	printf("+ WinScriptHost installed with %s.\n", language.c_str());
}

IScript::Request* CScriptEngine::NewRequest(const String& entry) {
	return new CScriptEngine::Request(*this);
}

IScript::Request& CScriptEngine::GetDefaultRequest() {
	return defaultRequest;
}

void CScriptEngine::RemoveDelegate(IScript::Object* proxy) {
	proxyPool.erase(proxy);
}

// Request Apis

CScriptEngine::Request::Request(CScriptEngine& h) : host(h), idx(0), initCount(0), tableLevel(0) {}
CScriptEngine::Request::~Request() {}

IScript* CScriptEngine::Request::GetScript() {
	return &host;
}

int CScriptEngine::Request::GetCount() {
	return buffer.size();
}

std::vector<IScript::Request::Key> CScriptEngine::Request::Enumerate() {
	std::vector<Key> keys;
	assert(host.GetLockCount() != 0);
	assert(tableLevel != 0);
	// get table
	CComVariant& var = buffer[buffer.size() - 2];
	assert(var.vt == VT_DISPATCH);
	IDispatch* dispatch = var.pdispVal;

	ITypeInfo* info;
	if (SUCCEEDED(dispatch->GetTypeInfo(0, 0, &info))) {
		// create success
		TYPEATTR* attr;
		if (SUCCEEDED(info->GetTypeAttr(&attr))) {
			for (DISPID j = 0; j < attr->cVars; j++) {
				VARDESC* var;
				if (SUCCEEDED(info->GetVarDesc(j, &var)) && var->memid < 0x10000000) {
					CComBSTR bstrName;
					CComBSTR bstrDoc;
					info->GetDocumentation(var->memid, &bstrName, &bstrDoc, NULL, NULL);
					keys.push_back(Key(SystemToUtf8(String((const char*)(BSTR)bstrName, ::SysStringByteLen(bstrName))), PaintsNow::IScript::Request::STRING));
				}
			}

			// Load funcs
			for (DISPID i = 0; i < attr->cFuncs; i++) {
				FUNCDESC* func = NULL;
				// Standard methods of IUnknown/IDispatch will hold memids which exceed 0x60000000, skip them
				if (SUCCEEDED(info->GetFuncDesc(i, &func)) && func->memid < 0x10000000) {
					CComBSTR bstrName;
					CComBSTR bstrDoc;
					info->GetDocumentation(func->memid, &bstrName, &bstrDoc, NULL, NULL);
					keys.push_back(Key(SystemToUtf8(String((const char*)(BSTR)bstrName, ::SysStringByteLen(bstrName))), PaintsNow::IScript::Request::FUNCTION));
				}
			}
		}
	}

	return keys;
}


static int IncreaseTableIndex(std::vector<CComVariant>& buffer, int count = 1) {
	CComVariant& var = buffer[buffer.size() - 1];
	assert(var.vt == VT_I4 || var.vt == VT_INT);
	if (count == 0) {
		var.intVal = 0;
	} else {
		var.intVal += count;
	}

	return var.intVal - 1;
}

template <class T>
IDispatch* CreateDispatch(T) {
	assert(false);
	return NULL;
}

template <>
IDispatch* CreateDispatch(IDispatch* dispatch) {
	return dispatch;
}

template <class C>
inline void Write(CScriptEngine::Request& request, C& value) {
	std::vector<CComVariant>& buffer = request.buffer;
	int& tableLevel = request.tableLevel;
	int& idx = request.idx;
	String& key = request.key;
	IScriptControlPtr& control = request.host.GetScriptControl();

	if (tableLevel != 0) {
		CComVariant& arr = buffer[buffer.size() - 2];
		assert(arr.vt & VT_DISPATCH);
		IDispatch* dispatch = arr.pdispVal;
		IObjectProxy* proxy;
		if (SUCCEEDED(dispatch->QueryInterface(__uuidof(IObjectProxy), (void**)&proxy))) {
			if (key.empty()) {
				int index = IncreaseTableIndex(buffer);
				WCHAR str[64];
				wsprintfW(str, L"%d", index);
				CComVariant var(value);
				proxy->SetAttrib(str, &var);
			} else {
				CComVariant var(value);
				CComBSTR bstr = (LPCWSTR)Utf8ToSystem(String(key)).data();
				proxy->SetAttrib(bstr, &var);
			}

			proxy->Release();
		}

		key = "";
	} else {
		buffer.push_back(CComVariant(value));
	}
}

IScript::Request& CScriptEngine::Request::operator << (const ArrayStart&) {
	return *this << TableStart();
}

IScript::Request& CScriptEngine::Request::operator << (const TableStart&) {
	assert(host.GetLockCount() != 0);
	CComPtr<IObjectProxy> proxy = CoCreateInstanceInternal<ObjectProxy, IObjectProxy>();
	proxy->SetHost(reinterpret_cast<LONGLONG>(&host));
	CComVariant var(proxy);

	if (tableLevel != 0) {
		CComVariant& arr = buffer[buffer.size() - 2];
		assert(arr.vt & VT_DISPATCH);
		IDispatch* dispatch = arr.pdispVal;
		IObjectProxy* proxy;
		if (SUCCEEDED(dispatch->QueryInterface(__uuidof(IObjectProxy), (void**)&proxy))) {
			if (key.empty()) {
				int index = IncreaseTableIndex(buffer);
				WCHAR str[64];
				wsprintfW(str, L"%d", index);
				proxy->SetAttrib(str, &var);
			} else {
				CComBSTR bstr = (LPCWSTR)Utf8ToSystem(String(key)).data();
				proxy->SetAttrib(bstr, &var);
			}

			proxy->Release();
		}
	} else {
		buffer.push_back(var);
	}

	key = "";
	// load table
	buffer.push_back(var);
	buffer.push_back(CComVariant((int)0));
	tableLevel++;
	return *this;
}

template <class C>
void ReadEx(C& value, const CComVariant& var) {
	value = *(C*)&var.parray;
}

template <>
void ReadEx(CComVariant& value, const CComVariant& var) {
	value = var;
}

template <class C>
inline void Read(CScriptEngine::Request& request, C& value) {
	std::vector<CComVariant>& buffer = request.buffer;
	int& tableLevel = request.tableLevel;
	int& idx = request.idx;
	String& key = request.key;

	if (tableLevel != 0) {
		CComVariant& arr = buffer[buffer.size() - 2];
		CComVariant result;
		assert(arr.vt & VT_DISPATCH);
		IDispatch* dispatch = arr.pdispVal;
		IObjectProxy* proxy;
		if (SUCCEEDED(dispatch->QueryInterface(__uuidof(IObjectProxy), (void**)&proxy))) {
			if (key.empty()) {
				int index = IncreaseTableIndex(buffer);
				WCHAR str[64];
				wsprintfW(str, L"%d", index);
				proxy->GetAttrib(str, &result);
			} else {
				CComBSTR bstr = (LPCWSTR)Utf8ToSystem(String(key)).data();
				proxy->GetAttrib(bstr, &result);
			}

			ReadEx(value, result);

			proxy->Release();
		}

		key = "";
	} else {
		if (idx == (int)buffer.size()) {
			buffer.push_back(CComVariant());
		}

		CComVariant& result = buffer[idx++];
		ReadEx(value, result);
	}
}


IScript::Request& CScriptEngine::Request::operator >> (ArrayStart& ts) {
	TableStart t;
	*this >> t;
	ts.count = t.count;
	return *this;
}

IScript::Request& CScriptEngine::Request::operator >> (TableStart& ts) {
	assert(host.GetLockCount() != 0);
	ts.count = 0; // not supported

	if (tableLevel == 0) {
		if (idx == (int)buffer.size()) {
			// create empty table
			CComPtr<IObjectProxy> proxy = CoCreateInstanceInternal<ObjectProxy, IObjectProxy>();
			proxy->SetHost(reinterpret_cast<LONGLONG>(&host));
			CComVariant var(proxy);
			buffer.push_back(var);
		}

		buffer.push_back(buffer[idx++]);
	} else {
		CComVariant var;
		CComVariant& arr = buffer[buffer.size() - 2];
		assert(arr.vt & VT_DISPATCH);
		IDispatch* dispatch = arr.pdispVal;
		IObjectProxy* proxy;
		if (SUCCEEDED(dispatch->QueryInterface(__uuidof(IObjectProxy), (void**)&proxy))) {
			if (key.empty()) {
				int index = IncreaseTableIndex(buffer);
				WCHAR str[64];
				wsprintfW(str, L"%d", index);
				proxy->GetAttrib(str, &var);
			} else {
				CComBSTR bstr = (LPCWSTR)Utf8ToSystem(String(key)).data();
				proxy->GetAttrib(bstr, &var);
			}

			proxy->Release();
		}

		if (var.vt != VT_DISPATCH) {
			CComPtr<IObjectProxy> proxy = CoCreateInstanceInternal<ObjectProxy, IObjectProxy>();
			proxy->SetHost(reinterpret_cast<LONGLONG>(&host));
			CComVariant var(proxy);
			buffer.push_back(var);
		} else {
			buffer.push_back(var);
		}
	}

	key = "";
	buffer.push_back(CComVariant((int)0));
	tableLevel++;
	return *this;
}

IScript::Request& CScriptEngine::Request::Push() {
	assert(host.GetLockCount() != 0);
	assert(key.empty());
	buffer.push_back(CComVariant(idx));
	buffer.push_back(CComVariant(tableLevel));
	buffer.push_back(CComVariant(initCount));
	initCount = buffer.size();
	idx = initCount;
	tableLevel = 0;

	return *this;
} 

IScript::Request& CScriptEngine::Request::Pop() {
	assert(host.GetLockCount() != 0);
	assert(key.empty());
	assert(tableLevel == 0);

	size_t org = initCount;

	initCount = buffer[org - 1].intVal;
	tableLevel = buffer[org - 2].intVal;
	idx = buffer[org - 3].intVal;

	buffer.resize(org - 3);
	return *this;
}



IScript::Request::Ref CScriptEngine::Request::Load(const String& script, const String& pa) {
	assert(host.GetLockCount() != 0);
	try {
		host.engine->AddCode((LPCWSTR)Utf8ToSystem(script).data());
		CComSafeArray<VARIANT> arr;
		arr.Create();
		LPSAFEARRAY p = arr;
		CComVariant var = host.engine->Run(_T("Main"), &p);
		if (var.vt == VT_DISPATCH) {
			return IScript::Request::Ref(reinterpret_cast<size_t>(var.pdispVal));
		}
	} catch (_com_error& e) {
		printf("Exception(%X): %S\n", e.Error(), (BSTR)e.Description());
	}

	// JScript doesn't support chunk function
	// So just return 0
	return IScript::Request::Ref(0);
}

IScript::Request& CScriptEngine::Request::operator << (const Nil&) {
	assert(host.GetLockCount() != 0);
	CComVariant var;
	var.ChangeType(VT_EMPTY);
	Write(*this, var);
	return *this;
}

IScript::Request& CScriptEngine::Request::operator << (const Global&) {
	assert(host.GetLockCount() != 0);
	buffer.push_back(CComVariant((IDispatch*)host.globalObject));
	buffer.push_back(CComVariant(0));
	tableLevel++;
	return *this;
}

IScript::Request& CScriptEngine::Request::operator << (const Local&) {
	assert(host.GetLockCount() != 0);
	assert(false); // Not implemented.
	return *this;
}

IScript::Request& CScriptEngine::Request::operator << (const Key& k) {
	assert(host.GetLockCount() != 0);
	assert(key.empty());
	key = k.GetKey();
	return *this;
}

IScript::Request& CScriptEngine::Request::operator >> (const Key& k) {
	assert(host.GetLockCount() != 0);
	key = k.GetKey();
	return *this;
}

IScript::Request& CScriptEngine::Request::operator << (double value) {
	assert(host.GetLockCount() != 0);
	Write(*this, value);
	return *this;
}

IScript::Request& CScriptEngine::Request::operator >> (double& value) {
	assert(host.GetLockCount() != 0);
	Read(*this, value);

	return *this;
}

IScript::Request& CScriptEngine::Request::operator << (const String& str) {
	assert(host.GetLockCount() != 0);
	String s = Utf8ToSystem(str);
	CComBSTR w = (LPCWSTR)s.data();
	BSTR r = (BSTR)w;
	Write(*this, r);

	return *this;
}

IScript::Request& CScriptEngine::Request::operator >> (String& str) {
	assert(host.GetLockCount() != 0);
	BSTR ptr = NULL;
	Read(*this, ptr);

	if (ptr != NULL) {
		str = SystemToUtf8(String((const char*)ptr, ::SysStringByteLen(ptr)));
	} else {
		str = "";
	}

	return *this;
}

IScript::Request& CScriptEngine::Request::operator << (const char* str) {
	assert(host.GetLockCount() != 0);
	assert(str != NULL);
	return *this << String(str);
}

IScript::Request& CScriptEngine::Request::operator >> (const char*& str) {
	assert(host.GetLockCount() != 0);
	assert(false); // Not allowed
	return *this;
}

IScript::Request& CScriptEngine::Request::operator << (bool b) {
	assert(host.GetLockCount() != 0);
	Write(*this, b);
	return *this;
}

IScript::Request& CScriptEngine::Request::operator >> (bool& b) {
	assert(host.GetLockCount() != 0);
	Read(*this, b);
	return *this;
}

IScript::Request& CScriptEngine::Request::operator << (const BaseDelegate& value) {
	assert(host.GetLockCount() != 0);
	IScript::Object* ptr = value.GetRaw();
	// lookup
	const std::map<IScript::Object*, IDispatch*>::const_iterator p = host.proxyPool.find(ptr);
	if (p != host.proxyPool.end()) {
		IDispatch* dispatch = (*p).second;
		Write(*this, dispatch);
	} else {
		CComPtr<IDelegateProxy> proxy = CoCreateInstanceInternal<DelegateProxy, IDelegateProxy>();
		proxy->SetProxy(reinterpret_cast<LONGLONG>(ptr));
		proxy->SetHost(reinterpret_cast<LONGLONG>(&host));
		IDispatch* dispatch = static_cast<IDispatch*>(proxy);
		Write(*this, dispatch);
		host.proxyPool[ptr] = dispatch;
	}

	return *this;
}

IScript::Request& CScriptEngine::Request::operator >> (BaseDelegate& value) {
	assert(host.GetLockCount() != 0);
	IDispatch* dispatch = NULL;
	Read(*this, dispatch);
	if (dispatch != NULL) {
		IDelegateProxy* proxy;
		if (SUCCEEDED(dispatch->QueryInterface(__uuidof(IDelegateProxy), (void**)&proxy))) {
			LONGLONG base;
			proxy->GetProxy(&base);
			value = IScript::BaseDelegate(reinterpret_cast<IScript::Object*>(base));
			proxy->Release();
		}
	}
	return *this;
}

IScript::Request& CScriptEngine::Request::operator << (const AutoWrapperBase& wrapper) {
	assert(host.GetLockCount() != 0);
	// register function
	assert(tableLevel != 0);
	CComVariant var;
	var.vt = VT_UI1 | VT_BYREF;
	var.pbVal = (BYTE*)&wrapper;
	Write(*this, var);
	var.punkVal = NULL;

	return *this;
}

IScript::Request& CScriptEngine::Request::operator << (int64_t u) {
	assert(host.GetLockCount() != 0);
	// Only 32-bit integers were supported
	LONG ul = (LONG)u;
	Write(*this, ul);
	return *this;
}

IScript::Request& CScriptEngine::Request::operator >> (int64_t& u) {
	assert(host.GetLockCount() != 0);
	LONG ul;
	Read(*this, ul);
	u = ul;

	return *this;
}

IScript::Request& CScriptEngine::Request::MoveVariables(IScript::Request& target, size_t count) {
	assert(host.GetLockCount() != 0);
	std::vector<CComVariant>& req = static_cast<CScriptEngine::Request&>(target).buffer;

	for (size_t i = buffer.size() - count - 1; i < buffer.size(); i++) {
		req.push_back(buffer[i]);
	}

	// shrink
	buffer.resize(buffer.size() - count);
	return *this;
}

IScript::Request& CScriptEngine::Request::operator << (const ArrayEnd&) {
	return *this << TableEnd();
}

IScript::Request& CScriptEngine::Request::operator << (const TableEnd&) {
	assert(host.GetLockCount() != 0);
	buffer.resize(buffer.size() - 2);
	tableLevel--;
	return *this;
}

IScript::Request& CScriptEngine::Request::operator >> (const ArrayEnd&) {
	return *this >> TableEnd();
}

IScript::Request& CScriptEngine::Request::operator >> (const TableEnd&) {
	assert(host.GetLockCount() != 0);
	buffer.resize(buffer.size() - 2);
	tableLevel--;
	return *this;
}

// TODO: 

IScript::Request& CScriptEngine::Request::operator >> (Ref& ref) {
	assert(host.GetLockCount() != 0);
	Read(*this, reinterpret_cast<IDispatch*&>(ref.value));
	return *this;
}

IScript::Request& CScriptEngine::Request::operator << (const Ref&) {
	assert(host.GetLockCount() != 0);
	Write(*this, reinterpret_cast<IDispatch*&>(ref.value));
	return *this;
}

void CScriptEngine::Request::FlushDeferredCalls() {}

bool CScriptEngine::Request::Call(const AutoWrapperBase& wrapper, const Request::Ref& g) {
	assert(host.GetLockCount() != 0);
	IDispatch* dispatch = reinterpret_cast<IDispatch*>(g.value);
	if (dispatch != NULL) {
		CComSafeArray<VARIANT> arr;
		arr.Create();
		arr.Add(CComVariant(dispatch));
		for (size_t i = initCount; i < buffer.size(); i++) {
			arr.Add(buffer[i]);
		}

		buffer.resize(initCount);
		LPSAFEARRAY ptr = arr;
		CComVariant var = host.engine->Run(L"_CallFunction", &ptr);
		buffer.push_back(var);

		if (!wrapper.IsSync()) {
			wrapper.Execute(*this);
		}

		return true;
	} else {
		return false;
	}
}

IScript::Request& CScriptEngine::Request::operator >> (const Skip& skip) {
	assert(host.GetLockCount() != 0);
	if (tableLevel != 0) {
		if (key.empty()) {
			IncreaseTableIndex(buffer, skip.count);
		}
	} else {
		idx += skip.count;
	}

	return *this;
}

IScript::Request& PaintsNow::CScriptEngine::Request::CleanupIndex() {
	idx = initCount;
	return *this;
}


IScript::Request& CScriptEngine::Request::operator << (const CComVariant& v) {
	assert(host.GetLockCount() != 0);
	Write(*this, v);
	return *this;
}

IScript::Request& CScriptEngine::Request::operator >> (CComVariant& v) {
	assert(host.GetLockCount() != 0);
	Read(*this, v);
	return *this;
}

IScript::Request::Ref CScriptEngine::Request::Reference(const Ref& d) {
	assert(host.GetLockCount() != 0);
	IDispatch* dispatch = reinterpret_cast<IDispatch*>(d.value);
	if (dispatch != NULL) {
		dispatch->AddRef();
	}

	return IScript::Request::Ref(d.value);
}

IScript::Request::TYPE CScriptEngine::Request::GetCurrentType() {
	assert(false); // not implemented.
	return NIL;
}

IScript::Request::TYPE CScriptEngine::Request::GetReferenceType(const Ref& d) {
	return IScript::Request::OBJECT;
}

void CScriptEngine::Request::Dereference(Ref& ref) {
	assert(host.GetLockCount() != 0);
	IDispatch* dispatch = reinterpret_cast<IDispatch*>(ref.value);
	if (dispatch != NULL) {
		dispatch->Release();
		ref.value = 0;
	}
}

const char* CScriptEngine::GetFileExt() const {
	static const char* extjs = "js";
	static const char* extvb = "vbs";
	return language == "JScript" ? extjs : extvb;
}

bool CScriptEngine::IsTypeCompatible(Request::TYPE target, Request::TYPE source) const {
	// everything is an object ...
	return source == IScript::Request::OBJECT || source == target;
}

IScript* CScriptEngine::NewScript() const {
	return new CScriptEngine(threadApi, language);
}