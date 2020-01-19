#include "StdAfx.h"

#include "WinScriptHost.h"
#include "ScriptEngine.h"
using namespace PaintsNow;

class Factory : public TFactoryBase<IScript> {
public:
	Factory() : TFactoryBase<IScript>(Wrap(this, &Factory::CreateEx)) {}

	IScript* CreateEx(const String& info = "") const {
		static ZThreadPthread pthread(false);
		return new CScriptEngine(pthread, info.empty() ? "JScript" : info);
	}
} theFactory;
