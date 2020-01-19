// WinHostScriptEngine.h: interface for the CScriptEngine class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_WINHOSTSCRIPTENGINE_H__73EC6330_BEC9_4DF8_8517_E0EDD301842E__INCLUDED_)
#define AFX_WINHOSTSCRIPTENGINE_H__73EC6330_BEC9_4DF8_8517_E0EDD301842E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


#import "Msscript.ocx" no_namespace

namespace PaintsNow {
	class CScriptEngine : public IScript {
	public:
		CScriptEngine(PaintsNow::IThread& threadApi, const String& language);
		virtual ~CScriptEngine();

		virtual void Reset();
		virtual IScript* NewScript() const;
		virtual IScript::Request* NewRequest(const String& entry);
		virtual IScript::Request& GetDefaultRequest();
		virtual const char* GetFileExt() const;
		virtual bool IsTypeCompatible(Request::TYPE target, Request::TYPE source) const;

		class Request : public IScript::Request {
		public:
			Request(CScriptEngine& host);
			virtual ~Request();

			virtual IScript* GetScript();
			virtual int GetCount();
			virtual bool Call(const AutoWrapperBase& base, const Ref& func);
			virtual void FlushDeferredCalls();
			virtual std::vector<Key> Enumerate();
			virtual TYPE GetCurrentType();
			virtual IScript::Request::Ref Load(const String& script, const String& pathname);
			virtual IScript::Request& Push();
			virtual IScript::Request& Pop();
			virtual IScript::Request& operator >> (Ref&);
			virtual IScript::Request& operator << (const Ref&);
			virtual IScript::Request& operator << (const Nil&);
			virtual IScript::Request& operator << (const BaseDelegate&);
			virtual IScript::Request& operator >> (BaseDelegate&);
			virtual IScript::Request& operator << (const Global&);
			virtual IScript::Request& operator << (const Local&);
			virtual IScript::Request& operator << (const TableStart&);
			virtual IScript::Request& operator >> (TableStart&);
			virtual IScript::Request& operator << (const TableEnd&);
			virtual IScript::Request& operator >> (const TableEnd&);
			virtual IScript::Request& operator << (const ArrayStart&);
			virtual IScript::Request& operator >> (ArrayStart&);
			virtual IScript::Request& operator << (const ArrayEnd&);
			virtual IScript::Request& operator >> (const ArrayEnd&);
			virtual IScript::Request& operator << (const Key&);
			virtual IScript::Request& operator >> (const Key&);
			virtual IScript::Request& operator << (double value);
			virtual IScript::Request& operator >> (double& value);
			virtual IScript::Request& operator << (const String& str);
			virtual IScript::Request& operator >> (String& str);
			virtual IScript::Request& operator << (const char* str);
			virtual IScript::Request& operator >> (const char*& str);
			virtual IScript::Request& operator << (bool b);
			virtual IScript::Request& operator >> (bool& b);
			virtual IScript::Request& operator << (const AutoWrapperBase& wrapper);
			virtual IScript::Request& operator << (int64_t u);
			virtual IScript::Request& operator >> (int64_t& u);
			IScript::Request& operator << (const CComVariant& v);
			IScript::Request& operator >> (CComVariant& v);
			virtual IScript::Request& operator >> (const Skip& skip);
			virtual IScript::Request& CleanupIndex();

			virtual Ref Reference(const Ref& d);
			virtual TYPE GetReferenceType(const Ref& d);
			virtual void Dereference(Ref& ref);
			virtual IScript::Request& MoveVariables(IScript::Request& target, size_t count);

			CScriptEngine& host;
			std::vector<CComVariant> buffer;
			String key;
			int idx;
			int initCount;
			int tableLevel;
		};

		friend class Request;

		IScriptControlPtr& GetScriptControl();
		void RemoveDelegate(IScript::Object* proxy);

		CComPtr<IObjectProxy> globalObject;
	private:
		IScriptControlPtr engine;
		Request defaultRequest;
		String language;
		std::map<IScript::Object*, IDispatch*> proxyPool;
	};
}

#endif // !defined(AFX_WINHOSTSCRIPTENGINE_H__73EC6330_BEC9_4DF8_8517_E0EDD301842E__INCLUDED_)
