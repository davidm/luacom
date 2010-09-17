# Copyright (C) 2007-2009 LuaDist.
# Redistribution and use of this file is allowed according to the terms of the MIT license.
# For details see the COPYRIGHT file distributed with LuaDist (and copied below).
# Please note that the package source code is licensed under its own license.

PROJECT(luacom C CXX)
CMAKE_MINIMUM_REQUIRED(VERSION 2.6)
INCLUDE(CheckIncludeFiles)
INCLUDE(CheckSymbolExists)  # note: doesn't support C++

# Cmake configuration:
SET (CMAKE_ALLOW_LOOSE_LOOP_CONSTRUCTS true)

# Where to install module parts:
SET (INSTALL_CMOD share/lua/cmod CACHE PATH "Directory to install Lua binary modules.")
SET (INSTALL_DATA . CACHE PATH "Directory the package can store documentation, tests or other data in.")
SET (INSTALL_DOC ${INSTALL_DATA}/doc CACHE PATH "Recommended directory to install documentation into.")
SET (INSTALL_TEST ${INSTALL_DATA}/test CACHE PATH "Recommended directory to install tests into.")
SET (INSTALL_EXAMPLE ${INSTALL_DATA}/example CACHE PATH "Recommended directory to install examples into.")

CHECK_INCLUDE_FILES("windows.h;htmlhelp.h" HAVE_HTMLHELP)
IF(NOT HAVE_HTMLHELP)
	ADD_DEFINITIONS(-DNO_HTMLHELP)
ENDIF(NOT HAVE_HTMLHELP)

# _stricmp (is there a simpler way of doing this?)
CHECK_SYMBOL_EXISTS(_stricmp string.h HAVE__STRICMP)
IF(HAVE__STRICMP)
	SET(STRICMP _stricmp)
ENDIF(HAVE__STRICMP)
IF(NOT STRICMP)
	CHECK_SYMBOL_EXISTS(stricmp string.h HAVE_STRICMP)
	IF(HAVE_STRICMP)
		SET(STRICMP stricmp)
	ENDIF(HAVE_STRICMP)
ENDIF(NOT STRICMP)
IF(NOT STRICMP)
	MESSAGE(FATAL_ERROR "_stricmp or stricmp not found")
ENDIF(NOT STRICMP)
ADD_DEFINITIONS(-D_stricmp=${STRICMP})

FIND_PACKAGE(Lua51 REQUIRED)
INCLUDE_DIRECTORIES(${LUA_INCLUDE_DIR})

#~ improve?
FIND_PROGRAM(LUA NAMES lua lua.bat)

# Compile luacom5.lua to luacom5.loh
ADD_CUSTOM_COMMAND(OUTPUT ${CMAKE_CURRENT_BINARY_DIR}/luacom5.loh
    DEPENDS src/library/luacom5.lua
	COMMAND "${LUA}" "${CMAKE_CURRENT_SOURCE_DIR}/mak/luac.lua" "-o" "${CMAKE_CURRENT_BINARY_DIR}/luacom5.lo" "${CMAKE_CURRENT_SOURCE_DIR}/src/library/luacom5.lua"
	COMMAND "${LUA}" "${CMAKE_CURRENT_SOURCE_DIR}/mak/bin2c.lua" "${CMAKE_CURRENT_BINARY_DIR}/luacom5.lo" ">${CMAKE_CURRENT_BINARY_DIR}/luacom5.loh")


# Build
INCLUDE_DIRECTORIES(include ${CMAKE_CURRENT_BINARY_DIR} src/dll src/library)

SET(SRC_LIB
	src/library/LuaAux.cpp
	src/library/luabeans.cpp
	src/library/luacom.cpp
	src/library/LuaCompat.cpp
	src/library/tCOMUtil.cpp
	src/library/tLuaCOM.cpp
	src/library/tLuaCOMClassFactory.cpp
	src/library/tLuaCOMConnPoints.cpp
	src/library/tLuaCOMEnumerator.cpp
	src/library/tLuaCOMException.cpp
	src/library/tLuaCOMTypeHandler.cpp
	src/library/tLuaControl.cpp
	src/library/tLuaDispatch.cpp
	src/library/tLuaObjList.cpp
	src/library/tLuaObject.cpp
	src/library/tLuaVector.cpp
	src/library/tLuaTLB.cpp
	src/library/tStringBuffer.cpp
	src/library/tUtil.cpp
)

SET_SOURCE_FILES_PROPERTIES(src/library/luacom.cpp PROPERTIES OBJECT_DEPENDS ${CMAKE_CURRENT_BINARY_DIR}/luacom5.loh)


SET(SRC_DLL src/dll/luacom_dll.cpp)
ADD_DEFINITIONS(-DLUACOM_DLL="luacom.dll")

IF(MSVC)
        ADD_DEFINITIONS(-D_CRT_SECURE_NO_DEPRECATE)
ENDIF(MSVC)

SET(LIBS kernel32 user32 gdi32 shell32 advapi32 ole32 winspool uuid oleaut32 shlwapi)
IF(HAVE_HTMLHELP)
	SET(LIBS ${LIBS} htmlhelp)
ENDIF(HAVE_HTMLHELP)


ADD_LIBRARY(luacom MODULE ${SRC_DLL} ${SRC_LIB})
TARGET_LINK_LIBRARIES(luacom ${LUA_LIBRARY} ${LIBS})
SET_TARGET_PROPERTIES(luacom PROPERTIES PREFIX "")

IF(MSVC)
	SET_TARGET_PROPERTIES(luacom PROPERTIES LINK_FLAGS "/DEF:${CMAKE_CURRENT_SOURCE_DIR}/src/dll/luacom_dll.def")
ENDIF(MSVC)

# Install all files and documentation
INSTALL (TARGETS luacom DESTINATION ${INSTALL_CMOD})
INSTALL (FILES INSTALL README announce.txt todo.txt DESTINATION ${INSTALL_DATA})
INSTALL (FILES doc/luacom.gif doc/luacom.pdf www/index.html DESTINATION ${INSTALL_DOC})
INSTALL (DIRECTORY demo DESTINATION ${INSTALL_EXAMPLE})

#debugging only: INSTALL (FILES src/library/luacom5.lua DESTINATION ${INSTALL_LMOD})

# ===============================================================================
#
# Copyright (C) 2007-2008 LuaDist.
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# ===============================================================================
