# Copyright (C) 2007-2009 LuaDist.
# Created by Peter Kapec, David Manura
# Redistribution and use of this file is allowed according to the terms of the MIT license.
# For details see the COPYRIGHT file distributed with LuaDist.
# Please note that the package source code is licensed under its own license.

PROJECT(luacom C CXX)
CMAKE_MINIMUM_REQUIRED(VERSION 2.6)
INCLUDE(dist.cmake)
INCLUDE(CheckCSourceCompiles)
INCLUDE(CheckSymbolExists) # note: requires 'C' in PROJECT.

# Build
INCLUDE_DIRECTORIES(include src/dll src/library ${CMAKE_CURRENT_BINARY_DIR})

# Detect htmlhelp (not normally available in Cygwin 1.7, MinGW, and Wine).
# Use CHECK_C_SOURCE_COMPILES because these two aren't reliable:
#   CHECK_INCLUDE_FILES("windows.h;htmlhelp.h" HAVE_HTMLHELP)
#          # On MinGW and Wine, htmlhelp.h may exist but not htmlhelp.lib.
#   CHECK_LIBRARY_EXISTS(htmlhelp HtmlHelpA "" HAVE_HTMLHELP) # detection fails in MSVC
set(CMAKE_REQUIRED_LIBRARIES htmlhelp)
CHECK_C_SOURCE_COMPILES("
#include <windows.h>
	#include <htmlhelp.h>
	int main() { HtmlHelp(NULL, NULL, HH_HELP_CONTEXT, 0); return 0; }
" HAVE_HTMLHELP)
set(CMAKE_REQUIRED_LIBRARIES)
IF(NOT HAVE_HTMLHELP)
	ADD_DEFINITIONS(-DNO_HTMLHELP)
ENDIF(NOT HAVE_HTMLHELP)

# _stricmp (is there a simpler way of doing this?)
CHECK_SYMBOL_EXISTS(_stricmp string.h HAVE__STRICMP)
IF(HAVE__STRICMP)
	SET(STRICMP _stricmp)
ENDIF()
IF(NOT STRICMP)
	CHECK_SYMBOL_EXISTS(stricmp string.h HAVE_STRICMP)
	IF(HAVE_STRICMP)
		SET(STRICMP stricmp)
	ENDIF()
ENDIF(NOT STRICMP)
IF(NOT STRICMP)
	MESSAGE(FATAL_ERROR "_stricmp or stricmp not found")
ENDIF()
ADD_DEFINITIONS(-D_stricmp=${STRICMP})

SET(SRC_LIB
	src/library/LuaAux.cpp
	src/library/luabeans.cpp
	src/library/luacom.cpp
	src/library/tLuaCOM.cpp
	src/library/tLuaCOMException.cpp
	src/library/tLuaCOMTypeHandler.cpp
	src/library/tLuaDispatch.cpp
	src/library/tLuaObjList.cpp
	src/library/tLuaVector.cpp
	src/library/tStringBuffer.cpp
	src/library/tUtil.cpp
	src/library/tCOMUtil.cpp
	src/library/tLuaCOMClassFactory.cpp
	src/library/tLuaCOMConnPoints.cpp
	src/library/LuaCompat.cpp
	src/library/tLuaCOMEnumerator.cpp
	src/library/tLuaObject.cpp
	src/library/tLuaControl.cpp
	src/library/tLuaTLB.cpp
)

SET(LIBS gdi32 shell32 advapi32 ole32 winspool uuid oleaut32 shlwapi) # kernel32 user32
IF(HAVE_HTMLHELP)
	SET(LIBS ${LIBS} htmlhelp)
ENDIF(HAVE_HTMLHELP)

ADD_LUA_BIN2C(${CMAKE_CURRENT_BINARY_DIR}/luacom5.loh src/library/luacom5.lua
	${CMAKE_CURRENT_SOURCE_DIR}/mak/bin2c.lua
	${CMAKE_CURRENT_SOURCE_DIR}/mak/luac.lua)
SET_SOURCE_FILES_PROPERTIES(src/library/luacom.cpp PROPERTIES OBJECT_DEPENDS ${CMAKE_CURRENT_BINARY_DIR}/luacom5.loh)

SET(SRC_DLL src/dll/luacom_dll.cpp)
ADD_DEFINITIONS(-DLUACOM_DLL="luacom.dll")

ADD_LUA_MODULE(luacom ${SRC_DLL} ${SRC_LIB} src/dll/luacom_dll.def)
TARGET_LINK_LIBRARIES(luacom ${LIBS})

# Install all files and documentation
INSTALL (TARGETS luacom DESTINATION ${INSTALL_CMOD})
#~ INSTALL (FILES src/library/luacom5.lua DESTINATION ${INSTALL_LMOD})

INSTALL (FILES INSTALL README COPYRIGHT announce.txt todo.txt DESTINATION ${INSTALL_DATA})
INSTALL (FILES doc/luacom.gif doc/luacom.pdf www/index.html DESTINATION ${INSTALL_DOC})
INSTALL (DIRECTORY demo DESTINATION ${INSTALL_EXAMPLE})

ADD_LUA_TEST(${CMAKE_CURRENT_SOURCE_DIR}/src/test/luacom_tests5.lua)
