@ECHO ON
cls
echo  ��NET����ת����VB�������ֱ�����õ���
set FrameworkPath=c:\WINDOWS\Microsoft.NET\Framework\v2.0.50727
pushd %FrameworkPath%

rem �����ļ�·��
set U8setup=c:\U8soft\Portal

 regasm /codebase /tlb:%U8setup%\USNPASink.dll.tlb %U8setup%\USNPASink.dll
 

pause