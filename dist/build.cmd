call "C:\BuildTools\VC\Auxiliary\Build\vcvars64.bat"
cd /d "F:\aaaaaaaaaaa\audio optimizer"
rc /nologo /fo "F:\aaaaaaaaaaa\audio optimizer\dist\resource.res" resource.rc
cl /nologo /std:c++17 /EHsc /W4 /DUNICODE /D_UNICODE /MT /O2 AudioOptimizer.cpp "F:\aaaaaaaaaaa\audio optimizer\dist\resource.res" /link /SUBSYSTEM:WINDOWS /OUT:"F:\aaaaaaaaaaa\audio optimizer\dist\AudioOptimizer.exe" user32.lib gdi32.lib comctl32.lib ole32.lib propsys.lib shell32.lib advapi32.lib powrprof.lib
