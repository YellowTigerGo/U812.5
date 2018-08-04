@ECHO ON
cls
echo  将NET程序转化成VB程序可以直接引用的类
set FrameworkPath=c:\WINDOWS\Microsoft.NET\Framework\v2.0.50727
pushd %FrameworkPath%

rem 设置文件路径
set U8setup=c:\U8soft\Portal

 regasm /codebase /tlb:%U8setup%\USNPASink.dll.tlb %U8setup%\USNPASink.dll
 

pause