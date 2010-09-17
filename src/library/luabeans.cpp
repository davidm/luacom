#include <stdlib.h>
#include <stdio.h>
#include <string.h>
extern "C"
{
#include <lua.h>
#include <lauxlib.h>
#include "LuaCompat.h"
}
#include "luabeans.h"


#include "tLuaCOMException.h"
#include "LuaAux.h"
#include "tUtil.h"

char* LuaBeans::tag_name          = NULL;
char* LuaBeans::udtag_name        = NULL;
LuaBeans::Events* LuaBeans::pEvents         = NULL;
const char* LuaBeans::module_name = NULL;

void LuaBeans::createBeans(lua_State *L,
                           const char* p_module_name,
                           const char* name)
{
  LUASTACK_SET(L);

  char lua_name[500];

  module_name = tUtil::strdup(p_module_name);

  sprintf(lua_name, "%s" ,name);
  tag_name = tUtil::strdup(lua_name);
  luaCompat_newLuaType(L, module_name, tag_name);

  sprintf(lua_name,"%s_UDTAG",name);
  udtag_name = tUtil::strdup(lua_name);
  luaCompat_newLuaType(L, module_name, udtag_name);
  LUASTACK_CLEAN(L, 0);
}


void LuaBeans::Clean()
{
  free(LuaBeans::tag_name);
  free(LuaBeans::udtag_name);
}

void LuaBeans::registerObjectEvents(lua_State* L, class Events& events)
{
  LUASTACK_SET(L);

  luaCompat_pushTypeByName(L, module_name, tag_name);

  if(events.gettable)
  {
    lua_pushcfunction(L, events.gettable);
    // there is no gettable_event in Lua5 with the semantics of Lua4
  }

  if(events.settable)
  {
    lua_pushstring(L, "__newindex");
    lua_pushcfunction(L, events.settable);
    lua_settable(L, -3);
  }

  if(events.noindex)
  {
    lua_pushcfunction(L, events.noindex);
    lua_setfield(L, -2, "__index");
  }

  if(events.call)
  {
    lua_pushcfunction(L, events.call);
    lua_setfield(L, -2, "__call");
  }

  if(events.gc)
  {
    lua_pushcfunction(L, events.gc);
    lua_setfield(L, -2, "__gc");
  }

  lua_pop(L, 1);

  LUASTACK_CLEAN(L, 0);
}

void LuaBeans::registerPointerEvents(lua_State* L, class Events& events)
{
  LUASTACK_SET(L);

  luaCompat_pushTypeByName(L, module_name, udtag_name);

  if(events.gettable)
  {
    lua_pushcfunction(L, events.gettable);
    // there is no gettable_event in Lua5 with the semantics of Lua4
  }

  if(events.settable)
  {
    lua_pushstring(L, "__newindex");
    lua_pushcfunction(L, events.settable);
    lua_settable(L, -3);
  }

  if(events.noindex)
  {
    lua_pushcfunction(L, events.noindex);
    lua_setfield(L, -2, "__index");
  }

  if(events.gc)
  {
    lua_pushcfunction(L, events.gc);
    lua_setfield(L, -2, "__gc");
  }

  lua_pop(L, 1);

  LUASTACK_CLEAN(L, 0);
}

void LuaBeans::push(lua_State* L, void* userdata )
{
  LUASTACK_SET(L);

  lua_newtable(L);

  lua_pushstring(L, "_USERDATA_REF_");

  luaCompat_pushTypeByName(L, module_name, udtag_name);
  luaCompat_newTypedObject(L, userdata);

  lua_settable(L, -3);

  luaCompat_pushTypeByName(L, module_name, tag_name);
  lua_setmetatable(L, -2);

  LUASTACK_CLEAN(L, 1);
}

void* LuaBeans::check_tag(lua_State* L, int index)
{
  void* userdata = from_lua(L, index);

  luaL_arg_check(L, (userdata!=NULL), index, "Object type is wrong");

  return userdata;
}

void* LuaBeans::from_lua(lua_State* L, int index)
{
  LUASTACK_SET(L);

  void *obj = NULL;

  lua_pushvalue(L, index);
  if (lua_istable(L, -1) && luaCompat_isOfType(L, module_name, tag_name))
  {
    lua_pushstring(L, "_USERDATA_REF_");
    lua_gettable(L, index);
    obj = *(void **)lua_touserdata(L, -1);
    lua_pop(L, 1);
  }
  lua_pop(L, 1);

  LUASTACK_CLEAN(L, 0);
  
  return obj;
}


/*lua_State* LuaBeans::getLuaState()
{
  return L;
}*/

