set LUA_VERSION=5
set LUA_INC=t:\lib\lua5\include
set LUA_LIB=t:\lib\lua5\lib\vc6
set LUA_LIBD=t:\lib\lua5\lib\dll
set DEBUG=YES
set BIN2C=bin2c5
set LUAC=luac5


nmake %1 %2 %3 %4 %5 %6
