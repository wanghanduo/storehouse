@echo off
mode con cols=85 lines=35
ver | find "�汾" > NUL && title ��ˮ��KMS�ű� V22.09.28 || title Cangshui's KMS script V22.09.28
setlocal EnableDelayedExpansion&color 70 & cd /d "%~dp0"
%1 %2
ver | find "5."> NUL && goto :start

setlocal
set uac=~uac_permission_tmp_%random%
md "%SystemRoot%\system32\%uac%" 2>nul
if %errorlevel%==0 ( rd "%SystemRoot%\system32\%uac%" >nul 2>nul ) else (
    echo set uac = CreateObject^("Shell.Application"^)>"%temp%\%uac%.vbs"
    echo uac.ShellExecute "%~s0","","","runas",1 >>"%temp%\%uac%.vbs"
    echo WScript.Quit >>"%temp%\%uac%.vbs"
    "%temp%\%uac%.vbs" /f
    del /f /q "%temp%\%uac%.vbs" & exit )
endlocal

:start
chcp 936 > NUL
for /f "tokens=3 delims= " %%i in ('reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "CurrentBuild"') do set CurrentBuild=%%i
if  %CurrentBuild% LEQ 17762 (
  set systabs=0
) else (
  set systabs=1
)
set KMS_Sev=kms-shanghai01.cangshui.net
ver | find "6.0." > NUL &&  set winv=vista
ver | find "6.1." > NUL &&  set winv=7
ver | find "6.2." > NUL &&  set winv=8
ver | find "6.3." > NUL &&  set winv=8.1
ver | find "10.0." > NUL &&  set winv=10
ver | find "�汾" >NUL && set syslang=cn
ver | find "�汾" >nul && echo ���ʽ���������http://kms.cangshui.net || echo Feedback: http://kms.cangshui.net
ver | find "�汾" >nul && echo �������������http://shop.cangshui.net || echo Tip: http://shop.cangshui.net
if  "%syslang%"=="cn" (
  if  "%systabs%"=="1" ( echo �X�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T����ѡ��T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�[ )
  echo �U��A��KMS����Windows                                                               �U
  echo �U��B��KMS����Office                                                                �U
  echo �U��C�����Windows KMS                                                              �U
  echo �U��D�����Office KMS                                                               �U
  echo �U��E���鿴֧�ֵ�windows�汾                                                        �U
  ) else (
  if  "%systabs%"=="1" ( echo �X�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�TActivation option�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�[ )
  echo �U[A] KMS activate windows                                                          �U
  echo �U[B] KMS activate Office                                                           �U
  echo �U[C] Clear Windows KMS                                                             �U
  echo �U[D] Clear Office KMS                                                              �U
  echo �U[E] Supported windows version                                                     �U
)

if  "%systabs%"=="1" ( echo �d�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�������ߨT�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�g )
echo �U��1��ȥ��Office��ʾ�����֤�������桱                                             �U 
echo �U��2��ȥ����ݷ�ʽС��ͷ                                                           �U
echo �U��3���ָ���ݷ�ʽС��ͷ                                                           �U
echo �U��4��Win11�л��ɰ������Ҽ��˵�                                                    �U
echo �U��5��Win11�ָ��°������Ҽ��˵�                                                    �U 
echo �U��6��ȥ����ݷ�ʽС����                                                           �U
echo �U��7���ָ���ݷ�ʽС����                                                           �U 
echo �U��8��ȥ��������ݷ�ʽʱ�ĺ�׺��-��ݷ�ʽ��                                        �U
echo �U��9��ȥ�����п�ִ���ļ�ʱ�ľ��浯��                                               �U
echo �U��10����������ӡ��˵��ԡ�ͼ��                                                    �U
if  "%systabs%"=="1" ( echo �d�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T����ѡ��T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�T�g )

ver | find "�汾" >nul && set /p xuanze=�U ���������ѡ��: || set /p xuanze=�U Please enter your choice:
if /i "%xuanze%"=="a" cls&goto start1
if /i "%xuanze%"=="b" cls&goto start2
if /i "%xuanze%"=="c" cls&goto start3
if /i "%xuanze%"=="d" cls&goto start4
if /i "%xuanze%"=="e" cls&goto start5
if /i "%xuanze%"=="g" cls&goto Feedback
if /i "%xuanze%"=="1" cls&goto removewarn
if /i "%xuanze%"=="2" cls&goto removearrow
if /i "%xuanze%"=="3" cls&goto recoveryarrow
if /i "%xuanze%"=="4" cls&goto classicmenu
if /i "%xuanze%"=="5" cls&goto modernmenu
if /i "%xuanze%"=="6" cls&goto removeshield
if /i "%xuanze%"=="7" cls&goto recoveshield
if /i "%xuanze%"=="8" cls&goto shortcut
if /i "%xuanze%"=="9" cls&goto removerunwarn
if /i "%xuanze%"=="10" cls&goto addmypcico

:start2
cls
echo.
ver | find "�汾" >nul && echo ���ʽ���������http://kms.cangshui.net || echo Feedback: http://kms.cangshui.net
ver | find "�汾" >nul && echo �������������http://shop.cangshui.net || echo Tip: http://shop.cangshui.net
echo.
if  "%KMS_Sev%"=="kms-shanghai01.cangshui.net" (
    echo ���ڼ���ܷ����ӵ�KMS��������...
    ) else (
    echo ���ӵ�KMS��������ʧ�ܣ����л������÷�����...
)
dir /a "tcping.exe" | find "258,560"  > NUL && set tcpingstatus=successful
if  "%tcpingstatus%"=="successful" (
    echo tcping�������...���ȴ�ʱ�䳬��60��ɳ����������нű� && tcping.exe %KMS_Sev% 1688 | find "0 successful" > NUL && goto failb
    ) else (
       if  "%winv%"=="10" (
          echo ======================================��ʾ��Ϣ=======================================
          echo ��ϵͳ�Դ���ping�����޷�׼ȷ�жϷ������Ƿ���ã���˽��Զ�����TCPing����
          echo TCPingΪ��ȫ�Ŀ�Դ���ߣ���Դ��ַΪhttps://github.com/jtilander/tcping
          echo ��������TCPing�������...
          echo ======================================��ʾ��Ϣ=======================================          
          curl --ssl-no-revoke --connect-timeout 3 -m 10 -s -O https://cangshui.net/-otherweb/kms/tcping.exe    
        ) else (
          echo. 
        )
) 


dir /a "tcping.exe" | find "258,560"  > NUL && set tcpingstatus2=successful
if  "%tcpingstatus2%"=="successful" (
    if "%tcpingstatus%"=="successful" ( echo. ) else ( echo tcping�������...���ȴ�ʱ�䳬��60��ɳ����������нű� && tcping.exe %KMS_Sev% 1688 | find "0 successful" > NUL && goto failb)
) else (
        if  "%winv%"=="10" (
          echo TCPing������ʧ�ܻ�����ԭ���²����ã�����ping�����������Ƿ���ã����Ĳ��Խ������һ��׼ȷ   
        ) else (
          echo ======================================��ʾ��Ϣ=======================================
          echo ���ϵͳ��windows10�����ϰ汾 �޷��Զ�����TCPing����
          echo ���ֻ����ping�����������Ƿ���ã����Ĳ��Խ������һ��׼ȷ
          echo ������������ش� https://cangshui.net/-otherweb/kms/tcping.exe ������
          echo ��������ڱ��ű�ͬĿ¼�£��������нű�����
          echo TCPing���߽�Ϊ���������Ƿ���ã�ȱʧҲ������������ϵͳ
          echo TCPingΪ��ȫ�Ŀ�Դ���ߣ���Դ��ַΪhttps://github.com/jtilander/tcping
          echo ======================================��ʾ��Ϣ=======================================
        )
    echo.
    echo ��ʼPing����...���ȴ�ʱ�䳬��60��ɳ����������нű�
    ping %KMS_Sev% | find "100% ��ʧ"  > NUL &&  goto failb
    ping %KMS_Sev% | find "100% loss"  > NUL &&  goto failb
    ping %KMS_Sev% | find "�Ҳ�������"  > NUL &&  goto failb
    ping %KMS_Sev% | find "not find host"  > NUL &&  goto failb
    ping %KMS_Sev% | find "ʧ��"  > NUL &&  goto failb
    ping %KMS_Sev% | find "fail"  > NUL &&  goto failb    
)


if  "%KMS_Sev%"=="kms-shanghai01.cangshui.net" (
    echo �����ܹ���������KMS��������...
    ) else (
    echo �����ܹ���������KMS���÷�����...
)
goto office

:office
echo ��鰲װ��office����
call :GetOfficePath 14 Office2010
call :ActOffice 14 Office2010
call :GetOfficePath 15 Office2013
call :ActOffice 15 Office2013
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" set _Office16Path=%ProgramFiles%\Microsoft Office\Office16
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" set _Office16Path=%ProgramFiles(x86)%\Microsoft Office\Office16
if DEFINED _Office16Path (echo.&echo �ѷ��� Office2016ϵ�����[����2016/2019/365/2021]
    ping 127.0.0.1 -n 2 > nul
    call :ActOffice 16 Office2016
  ) else (echo.&echo δ���� Office2016ϵ�����[����2016/2019/365/2021])


echo.&pause
exit

:ActOffice
if DEFINED _Office%1Path (
    cd /d "!_Office%1Path!"
    if %1 EQU 16 call :Licens16
    echo.&echo ���Լ�������Office ...&echo.
cscript //nologo ospp.vbs /sethst:%KMS_Sev% > NUL
cscript //nologo ospp.vbs /act | find /i "successful" && (
        echo.&echo ***** ����ɹ� *****   & echo.) || (echo.&echo ***** ����ʧ�� ***** & echo.)
)    
cd /d "%~dp0"
goto :EOF

:GetOfficePath
echo.&echo ���ڼ�� %2 ϵ�в�Ʒ�İ�װ·��...
set _Office%1Path=
set _Reg32=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\%1.0\Common\InstallRoot
set _Reg64=HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\%1.0\Common\InstallRoot
reg query "%_Reg32%" /v "Path" > nul 2>&1 && FOR /F "tokens=2*" %%a IN ('reg query "%_Reg32%" /v "Path"') do SET "_OfficePath1=%%b"
reg query "%_Reg64%" /v "Path" > nul 2>&1 && FOR /F "tokens=2*" %%a IN ('reg query "%_Reg64%" /v "Path"') do SET "_OfficePath2=%%b"
if DEFINED _OfficePath1 (if exist "%_OfficePath1%ospp.vbs" set _Office%1Path=!_OfficePath1!)
if DEFINED _OfficePath2 (if exist "%_OfficePath2%ospp.vbs" set _Office%1Path=!_OfficePath2!)
set _OfficePath1=
set _OfficePath2=
if DEFINED _Office%1Path (echo.&echo �ѷ��� %2) else (echo.&echo δ���� %2)
goto :EOF

:Licens16
cls
echo ��A������ΪOffice2021�汾(��2021�����ϰ汾��ѡ)
echo ��B������ΪOffice2019�汾(��2019�����ϰ汾��ѡ)
echo ��C������ΪOffice2016�汾(ȫ�汾ͨ��)
echo PS��Office365�汾��û�����������ģ��������365�汾ѡC����
set /p xuanze=��ѡ��...
if /i "%xuanze%"=="a" cls&goto installOffice21
if /i "%xuanze%"=="b" cls&goto installOffice19
if /i "%xuanze%"=="c" cls&goto installOffice16


:installOffice21
echo ��װ2021֤��
for /f %%x in ('dir /b ..\root\Licenses16\proplus2021vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\proplus2021vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\proplus2021previewvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\proplus2021previewvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\pkeyconfig-office-client15.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\pkeyconfig-office.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\visioPro2021VL_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\visioPro2021VL_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\visiopro2021previewvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\visiopro2021previewvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\projectpro2021vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\projectpro2021vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\projectpro2021previewvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\projectpro2021previewvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
cscript "%_Office16Path%\ospp.vbs" /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH > NUL
cscript "%_Office16Path%\ospp.vbs" /inpkey:KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4 > NUL
cscript "%_Office16Path%\ospp.vbs" /inpkey:FTNWT-C6WBT-8HMGF-K9PRX-QV9H8 > NUL
goto :EOF
exit

:installOffice19
echo ��װ2019֤��
for /f %%x in ('dir /b ..\root\Licenses16\proplus2019xc2rvl*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\proplus2019vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\proplus2019vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\pkeyconfig-office-client15.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\pkeyconfig-office.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\visioPro2019vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\visioPro2019vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\projectpro2019vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\projectpro2019vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\projectpro2019xc2rvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\projectpro2019xc2rvl_makc2r*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
cscript "%_Office16Path%\ospp.vbs" /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP > NUL
cscript "%_Office16Path%\ospp.vbs" /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB > NUL
cscript "%_Office16Path%\ospp.vbs" /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B > NUL
goto :EOF
exit


:installOffice16
echo ��װ2016֤��
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\client-issuance*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\pkeyconfig-office-client15.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\pkeyconfig-office.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\visioProvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\visioProvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\projectprovl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\projectprovl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\projectproxc2rvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
for /f %%x in ('dir /b ..\root\Licenses16\projectproxc2rvl_makc2r*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" > NUL
cscript "%_Office16Path%\ospp.vbs" /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 > NUL
cscript "%_Office16Path%\ospp.vbs" /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK > NUL
cscript "%_Office16Path%\ospp.vbs" /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT > NUL

goto :EOF
exit





:start1
cls
ver | find "�汾" >nul && echo ���ʽ���������http://kms.cangshui.net || echo Feedback: http://kms.cangshui.net
ver | find "�汾" >nul && echo �������������http://shop.cangshui.net || echo Tip: http://shop.cangshui.net
echo.
if  "%KMS_Sev%"=="kms-shanghai01.cangshui.net" (
    echo ���ڼ���ܷ����ӵ�KMS��������...
    ) else (
    echo ���ӵ�KMS��������ʧ�ܣ����л������÷�����...
)
dir /a "tcping.exe" | find "258,560"  > NUL && set tcpingstatus=successful
if  "%tcpingstatus%"=="successful" (
    echo tcping�������...���ȴ�ʱ�䳬��60��ɳ����������нű� && tcping.exe %KMS_Sev% 1688 | find "0 successful" > NUL && goto faila
) else (
       if  "%winv%"=="10" (
          echo ======================================��ʾ��Ϣ=======================================
          echo ��ϵͳ�Դ���ping�����޷�׼ȷ�жϷ������Ƿ���ã���˽��Զ�����TCPing����
          echo TCPingΪ��ȫ�Ŀ�Դ���ߣ���Դ��ַΪhttps://github.com/jtilander/tcping
          echo ��������TCPing�������...
          echo ======================================��ʾ��Ϣ=======================================     
          curl --ssl-no-revoke --connect-timeout 3 -m 10 -s -O https://cangshui.net/-otherweb/kms/tcping.exe   
        ) else (
          echo.
        )
        
) 


dir /a "tcping.exe" | find "258,560"  > NUL && set tcpingstatus2=successful
if  "%tcpingstatus2%"=="successful" (
    if "%tcpingstatus%"=="successful" ( echo. ) else ( echo tcping�������...���ȴ�ʱ�䳬��60��ɳ����������нű� && tcping.exe %KMS_Sev% 1688 | find "0 successful" > NUL && goto faila)
) else (
    if  "%winv%"=="10" (
          echo TCPing������ʧ�ܻ�����ԭ���²����ã�����ping�����������Ƿ���ã����Ĳ��Խ������һ��׼ȷ   
        ) else (
          echo ======================================��ʾ��Ϣ=======================================
          echo ���ϵͳΪwindows7 �޷��Զ�����TCPing����
          echo ���ֻ����ping�����������Ƿ���ã����Ĳ��Խ������һ��׼ȷ
          echo ������������ش� https://cangshui.net/-otherweb/kms/tcping.exe ������
          echo ��������ڱ��ű�ͬĿ¼�£��������нű�����
          echo TCPing���߽�Ϊ���������Ƿ���ã�ȱʧҲ������������ϵͳ
          echo TCPingΪ��ȫ�Ŀ�Դ���ߣ���Դ��ַΪhttps://github.com/jtilander/tcping
          echo ======================================��ʾ��Ϣ=======================================
        )
    echo.
    echo ��ʼPing����...���ȴ�ʱ�䳬��60��ɳ����������нű�
    ping %KMS_Sev% | find "100% ��ʧ"  > NUL &&  goto faila
    ping %KMS_Sev% | find "100% loss"  > NUL &&  goto faila
    ping %KMS_Sev% | find "�Ҳ�������"  > NUL &&  goto faila
    ping %KMS_Sev% | find "not find host"  > NUL &&  goto faila
    ping %KMS_Sev% | find "ʧ��"  > NUL &&  goto faila
    ping %KMS_Sev% | find "fail"  > NUL &&  goto faila    
)

if  "%KMS_Sev%"=="kms-shanghai01.cangshui.net" (
    echo �����ܹ���������KMS��������...
    ) else (
    echo �����ܹ���������KMS���÷�����...
    )

echo ======================================������Ϣ=======================================

ver | find "6.0." > NUL &&  goto winvista
ver | find "6.1." > NUL &&  goto win7
ver | find "6.2." > NUL &&  goto win8
ver | find "6.3." > NUL &&  goto win81
ver | find "10.0." > NUL &&  goto win10
echo δ�ҵ����ʵ�NT6ϵͳ��������WinXP��Win2003��
goto office

:winvista
echo ��ǰΪWindows Vista/2008��
set Business=YFKBB-PQJJV-G996G-VWGXY-2V3X8
set BusinessN=HMBQG-8H2RH-C77VX-27R82-VMQBT
set Enterprise=VKK3X-68KWM-X2YGT-QR4M6-4BWMV
set EnterpriseN=VTC42-BM838-43QHV-84HX6-XJXKV
set ServerWeb=WYR28-R7TFJ-3X2YQ-YCY4H-M249D
set ServerStandard=TM24T-X9RMF-VWXK6-X8JC9-BFGM2
set ServerStandardV=W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ
set ServerEnterprise=YQGMW-MPWTJ-34KDK-48M3W-X4Q6V
set ServerEnterpriseV=39BXF-X8Q23-P2WWT-38T2F-G3FPG
set ServerWeb=RCTX3-KWVHP-BR6TB-RB6DM-6X7HP
set ServerDatacenter=7M67G-PC374-GR742-YH8V4-TCBY3
set ServerDatacenterV=22XQ2-VRXRG-P8D42-K34TD-G3QQC
set ServerEnterpriseIA64=4DWFP-JF3DJ-B7DTH-78FJB-PDRHK
goto windowsstart

:win7
echo ��ǰΪWindows 7/2008 R2��
for /f "tokens=*" %%i in ('reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "ProductName"') do set ProductNamea=%%i
echo "%ProductNamea%" | find "Ultimate" >nul && (  
  msg %username% /time:99999999 "Windows7 �콢���޷�ʹ��KMS��������ϵͳ�汾���ȡ������ʽ����ϵͳ��"
  pause
  exit
  ) || (  
  echo .
)
set Professional=FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
set ProfessionalN=MRPKT-YTG23-K7D7T-X2JMM-QY7MG
set ProfessionalE=W82YF-2Q76Y-63HXB-FGJG9-GF7QX
set Enterprise=33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
set EnterpriseN=YDRBP-3D83W-TY26F-D46B2-XCKRJ
set EnterpriseE=C29WB-22CC8-VJ326-GHFJW-H9DH4
set ServerWeb=6TPJF-RBVHG-WBW2R-86QPH-6RTM4
set ServerHPC=TT8MH-CG224-D3D7Q-498W2-9QCTX
set ServerStandard=YC6KT-GKW9T-YTKYR-T4X34-R7VHC
set ServerEnterprise=489J6-VHDMP-X63PK-3K798-CPX3Y
set ServerDatacenter=74YFP-3QFB3-KQT8W-PMXWJ-7M648
set ServerEnterpriseIA64=GT63C-RJFQ3-4GMB6-BRFB9-CB83V
goto windowsstart

:win8
echo ��ǰΪWindows 8/2012��
set Professional=NG4HW-VH26C-733KW-K6F98-J8CK4
set ProfessionalN=XCVCF-2NXM9-723PB-MHCB7-2RYQQ
set Core=BN3D2-R7TKB-3YPBD-8DRP2-27GG4
set Enterprise=32JNW-9KQ84-P47T8-D8GGY-CWCK7
set EnterpriseN=JMNMF-RHW7P-DMY6X-RF3DR-X2BQT
set CoreN=8N2M2-HWPGY-7PGT9-HGDD8-GVGGY
set CoreSingleLanguage=2WN2H-YGCQR-KFX6K-CD6TF-84YXQ
set CoreCountrySpecific=4K36P-JN4VD-GDC6V-KDT89-DYFKP
set ServerMultiPointPremium=XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G
set ServerMultiPointStandard=HM7DN-YVMH3-46JC3-XYTG7-CYQJJ
set ServerStandard=XC9B7-NBPP2-83J2H-RHMBY-92BT4
set ServerDatacenter=48HP8-DN98B-MYWDG-T2DCC-8W83P
goto windowsstart

:win81
echo ��ǰΪWindows 8.1��
set Professional=GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
set ProfessionalN=HMCNV-VVBFX-7HMBH-CTY9B-B4FXY
set Enterprise=MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
set EnterpriseN=TT4HM-HN7YT-62K67-RGRQJ-JFFXW
set ServerSolution=KNC87-3J2TX-XB4WP-VCPJV-M4FWM
set ServerStandard=D2N9P-3P6X9-2R39C-7RTCD-MDVJX
set ServerDatacenter=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
set EmbeddedIndustry=32JNW-9KQ84-P47T8-D8GGY-CWCK7
goto windowsstart

:win10
for /f "tokens=*" %%i in ('reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "ProductName"') do set ProductNameb=%%i
echo "%ProductNameb%" | find "Server" >nul && (  
  goto win10Server
  ) || (  
    echo "%ProductNameb%" | find "Enterprise" >nul && (  
      goto win10Enterprise
      ) || (  
      echo ��ǰΪWindows 10��
      )
)
set Core=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
set CoreCountrySpecific=PVMJN-6DFY6-9CCP6-7BKTT-D3WVR
set CoreN=3KHY7-WNT83-DGQKR-F7HPR-844BM
set CoreSingleLanguage=7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
set Professional=W269N-WFGWX-YVC9B-4J6C9-T83GX
set ProfessionalN=MH37W-N47XK-V7XM9-C7227-GCQG9
set Education=NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
set EducationN=2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
set ProfessionalEducation=6TP4R-GNPTD-KYYHQ-7B7DP-J447Y
set ProfessionalEducationN=YVWGF-BXNMC-HTQYQ-CPQ99-66QFC
set ProfessionalWorkstation=NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
set ProfessionalWorkstations=NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
set ProfessionalWorkstationsN=9FNHH-K3HBT-3W4TD-6383H-6XYWF
goto windowsstart


:win10Enterprise
for /f "tokens=*" %%i in ('reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "ProductName"') do set ProductNameh=%%i
echo "%ProductNameh%" | findstr "2021" >nul && ( 
  echo ��ǰΪWindows Enterprise LTSC 2021�� 
  set EnterpriseS=M7XTQ-FN8P6-TTKYV-9D4CC-J462D
  set EnterpriseSN=92NFX-8DJQP-P6BBQ-THF9C-7CG2H
  ) || (      
     echo "%ProductNameh%" | findstr "2019" >nul && ( 
      echo ��ǰΪWindows Enterprise LTSC 2019�� 
      set EnterpriseS=M7XTQ-FN8P6-TTKYV-9D4CC-J462D
      set EnterpriseSN=92NFX-8DJQP-P6BBQ-THF9C-7CG2H
      ) || ( 
         echo "%ProductNameh%" | findstr "2016" >nul && ( 
           echo ��ǰΪWindows Enterprise LTSB 2016�� 
           set EnterpriseS=DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ
           set EnterpriseSN=QFFDN-GRT3P-VKWWX-X7T3R-8B639
           ) || (
                echo "%ProductNameh%" | findstr "2015" >nul && ( 
                  echo ��ǰΪWindows Enterprise LTSB 2015��
                  set EnterpriseS=WNMTR-4C88C-JK8YV-HQ7T2-76DF9
                  set EnterpriseSN=2F77B-TNFGY-69QQF-B8YKP-D69TJ
                  ) || (
                  echo ������ĳ����ҵ���ư汾...����֤�ܼ���ɹ�...
                  set EnterpriseS=M7XTQ-FN8P6-TTKYV-9D4CC-J462D
                  set EnterpriseSN=92NFX-8DJQP-P6BBQ-THF9C-7CG2H
                  set EnterpriseG=YYVX9-NTFWV-6MDM3-9PT4T-4M68B
                  set EnterpriseGN=44RPN-FTY23-9VTTB-MP9BX-T84FV
                  set Enterprise=NPPR9-FWDCX-D2C8J-H872K-2YT43
                  set EnterpriseN=DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
                )
          )
     )
)       
goto windowsstart




:win10Server
for /f "tokens=*" %%i in ('reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "ProductName"') do set ProductNamec=%%i
echo "%ProductNamec%" | findstr "2022" >nul && ( 
  echo ��ǰΪWindows server 2022�� 
  set ServerDatacenter=WX4NM-KYWYW-QJJR4-XV3QB-6VM33
  set ServerStandard=VDYBN-27WPP-V4HQT-9VMD4-VMK7H
  ) || ( 
       echo "%ProductNamec%" | findstr "2019" >nul && ( 
         echo ��ǰΪWindows server 2019�� 
         set ServerDatacenter=WMDGN-G9PQG-XVVXX-R3X43-63DFG
         set ServerStandard=N69G4-B89J2-4G8F4-WWYCC-J464C
         set ServerEssentials=WVDHN-86M7X-466P6-VHXV7-YY726
         set ServerRdsh=CPWHC-NT2C7-VYW78-DHDB2-PG3GK
         ) || ( 
              echo "%ProductNamec%" | findstr "2016" >nul && ( 
                echo ��ǰΪWindows server 2016�� 
                set ServerDatacenter=CB7KF-BWN84-R7R2Y-793K2-8XDDG
                set ServerStandard=WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
                set ServerEssentials=JCKRF-N37P4-C2D82-9YXRT-4M63B
                ) || ( 
                echo �޷�ʶ��ϵͳ�汾����
                goto Feedback
              )
          
        )
)
goto windowsstart



:windowsstart
echo ����Windows Update ����Ϊ�Զ�������...
sc config wuauserv start=auto > NUL
set winupdate=0
net start | find "Windows Update" > NUL && set winupdate=1
if "%winupdate%"==0 ( echo. > NUL ) else ( net start wuauserv > NUL )
for /f "tokens=3 delims= " %%i in ('reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "EditionID"') do set EditionID=%%i
if defined %EditionID% (
  echo �汾IDΪ%EditionID%
  goto windowsstart2
) else (
	echo �Ҳ������кš���
  goto Feedback
)
echo.&pause
exit

:windowsstart2
for /f "delims=" %%i in ('cscript //Nologo %windir%\system32\slmgr.vbs /ipk !%EditionID%!') do set kmsresulta=%%i
echo %kmsresulta%
echo %kmsresulta% | find "�Ǻ��İ汾�ļ����" > NUL && goto keyerror
echo %kmsresulta% | find "Windows non-core edition" > NUL && goto keyerror
cscript //Nologo %windir%\system32\slmgr.vbs /skms %KMS_Sev%
for /f "delims=" %%i in ('cscript //Nologo %windir%\system32\slmgr.vbs /ato') do set kmsresultc=%%i
echo %kmsresultc%
echo %kmsresultc% | find "�޷���ϵ�κ���Կ�������" > NUL && set retrya=1 && goto networkerror
echo %kmsresultc% | find "could be contacted" > NUL && set retrya=1 && goto networkerror 
echo ======================================������Ϣ======================================
echo.&pause
exit

:start4
set nextunnum=0
echo �Ƿ����Ҫ���Office��KMS����(��֧��Office2016�����ϰ汾)��
set /p xuanze=��Y������   ��N���ر�
if /i "%xuanze%"=="y" goto nextun
if /i "%xuanze%"=="n" exit

:nextun
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs"  (
goto nextun64
) else (
goto nextun32
)

:nextun64
cscript "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /dstatus | find  /I "No installed product keys detected"  > NUL && goto nextunsuccess
for /f "tokens=*" %%i in (' cscript "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /dstatus  ^| find /I "Last 5 characters of installed product key:" ') do set office5key=%%i
set "office5key=%office5key:~-5,5%"
cscript  "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /unpkey:%office5key% > NUL
cscript  "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /remhst > NUL
set /a nextunnum+=1
cls
echo �������%nextunnum%/10
if "%nextunnum%"=="10" ( goto nextunsuccess )
goto nextun64
pause
exit

:nextun32
cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /dstatus | find  /I "No installed product keys detected"  > NUL && goto nextunsuccess
for /f "tokens=*" %%i in (' cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /dstatus  ^| find /I "Last 5 characters of installed product key:" ') do set office5key=%%i
set "office5key=%office5key:~-5,5%"
cscript  "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /unpkey:%office5key% > NUL
cscript  "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /remhst > NUL
set /a nextunnum+=1
cls
echo �������%nextunnum%/10
if "%nextunnum%"=="10" ( goto nextunsuccess )
goto nextun32
pause
exit

:nextunsuccess
cls
echo ������
pause
exit

:start3
set /p xuanze=�Ƿ����Ҫ���Windows��KMS����Y������   ��N���ر�
if /i "%xuanze%"=="y" goto nextunw
if /i "%xuanze%"=="n" exit
:nextunw
slmgr /upk
slmgr /ckms
slmgr /rearm
cls
echo �����ɣ�����������
ping 127.0.0.1 -n 10 > nul



:start5
cls
echo.
echo windows 11��
echo Windows 11 ������                    Windows 11 רҵ������
echo Windows 11 ��ҵ��                    Windows 11 רҵ����վ��
echo Windows 11 רҵ��    
echo. 
echo Windows 10��
echo Windows 10 ������                    Windows 10 רҵ������
echo Windows 10 ��ҵ��                    Windows 10 רҵ����վ�� 
echo Windows 10 רҵ��                 
echo. 
echo Windows Server��
echo Windows Server version 1709-1909 �������İ�  Windows Server version 1709-1909 ��׼��
echo Windows Server 2012 �������İ�                         Windows Server 2012 ��׼��
echo Windows Server 2016 �������İ�                         Windows Server 2016 ��׼��
echo Windows Server 2019 �������İ�                         Windows Server 2019 ��׼��
echo Windows Server 2022 �������İ�                         Windows Server 2022 ��׼��
echo.
echo Windows Enterprise��
echo Windows LTSC 2019                   Windows LTSB 2016
echo Windows LTSB 2015
echo. 
echo Windows 8.1��
echo Windows 8.1 רҵ��                    Windows 8.1 ��ҵ��
echo. 
echo Windows 7��
echo Windows 7 רҵ��                       Windows 7 ��ҵ��
pause
cls
goto start




:faila
cls
if  "%KMS_Sev%"=="kms-shanghai01.cangshui.net" (
    set KMS_Sev=kms.cangshui.net && goto start1
    ) else (
    ver | find "�汾" >nul && echo ���ӵ�KMS��/����������ʧ�ܣ����������нű�����������������... || echo Unable to connect to KMS server
    )
pause



:failb
cls
if  "%KMS_Sev%"=="kms-shanghai01.cangshui.net" (
    set KMS_Sev=kms.cangshui.net && goto start2
    ) else (
    ver | find "�汾" >nul && echo ���ӵ�KMS��/����������ʧ�ܣ����������нű�����������������... || echo Unable to connect to KMS server
    )
pause

:networkerror
set /a retrya=1+%retrya%
if "%retrya%" LEQ "5" (
  goto networkerror2
) else (
  echo ======================================������Ϣ=======================================
  echo ��������KMS���������ʧ��...������������...
  goto Feedback
)
pause
exit

:networkerror2
echo �򱾻�����KMS������ʧ�ܣ����ڽ��е�%retrya%������..
for /f "delims=" %%i in ('cscript //Nologo %windir%\system32\slmgr.vbs /ato') do set kmsresultd=%%i
echo %kmsresultd% | find "�޷���ϵ�κ���Կ�������" > NUL  && goto networkerror
echo %kmsresultd% | find "could be contacted" > NUL &&  goto networkerror 
pause
exit

:keyerror
echo ======================================������Ϣ=======================================
echo �����ܳ״��󣬿����ǽű���֧�����ϵͳ�汾...
goto Feedback
pause
exit


:Feedback
echo.
if "!%EditionID%!"=="" ( echo. > NUL ) else ( echo windows����ʱʹ�õ��ܳ�Ϊ!%EditionID%!  )
if "%KMS_Sev%"=="" ( echo. > NUL ) else ( echo ����ʹ�õķ�����Ϊ%KMS_Sev%  )
for /f "tokens=*" %%d in ('reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "ProductName"') do set ProductNamed=%%d
echo �汾Ϊ%ProductNamed%
for /f "tokens=*" %%f in ('reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "EditionID"') do set EditionID=%%f
echo IDΪ%EditionID%
where curl > NUL
if "%errorlevel%"=="0" ( Echo ϵͳ�Ѱ�װcurl���� ) else ( echo ϵͳδ��װcurl���� )
where tcping > NUL
if "%errorlevel%"=="0" ( Echo ϵͳ�Ѱ�װTcping���� ) else ( echo ϵͳδ��װTcping���� )
whoami /groups | find "S-1-16-12288" >NUL && Echo �ű�ӵ�й���ԱȨ��
echo ======================================������Ϣ=======================================
pause
cls
goto start

:removearrow
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons" /v 29 /d "%systemroot%\system32\imageres.dll,197" /t reg_sz /f > nul
taskkill /f /im explorer.exe > nul
start explorer > nul
echo ȥ����ݷ�ʽ��ͷ�������...
pause
cls
goto start

:recoveryarrow
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons" /v 29 /f > nul
taskkill /f /im explorer.exe > nul
start explorer > nul
echo �ָ���ݷ�ʽ��ͷ�������...
pause
cls
goto start



:modernmenu
reg delete "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" /f  > nul
taskkill /f /im explorer.exe > nul
start explorer > nul
echo �л�Ϊ�ִ������Ҽ��˵��������...
pause
cls
goto start

:classicmenu
reg add "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" /f  > nul
taskkill /f /im explorer.exe > nul
start explorer > nul
echo �л�Ϊ���������Ҽ��˵��������...
pause
cls
goto start


:removewarn
reg add HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v AudienceId /t REG_SZ /d "55336B82-A18D-4DD6-B5F6-9E5095C314A6" /f > nul
reg add HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v CDNBaseUrl /t REG_SZ /d "http://officecdn.microsoft.com/pr/55336B82-A18D-4DD6-B5F6-9E5095C314A6" /f > nul
reg add HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v UpdateChannel /t REG_SZ /d "http://officecdn.microsoft.com/pr/55336B82-A18D-4DD6-B5F6-9E5095C314A6" /f > nul
reg delete HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v UpdateUrl /f  > nul
reg delete HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v UpdateToVersion /f  > nul
reg delete HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Updates /v UpdateToVersion /f > nul
reg delete HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Common\OfficeUpdate\ /f > nul
"%CommonProgramFiles%\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user > nul
echo ��ȴ�����������Office���´��ڡ��������...
echo ����ʾ��Ҫ�ر�office������������Ȼ���ٴδ�Office�鿴Ч��...
pause
cls
goto start

:shortcut
reg add HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer /v Link /t REG_BINARY /d "00000000" /f > nul
echo ȥ��������ݷ�ʽʱ�ĺ�׺��-��ݷ�ʽ�������ɹ�...
echo ������Ҫ���������������Ч...
pause
cls
goto start


:removeshield
reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons" /v 77 /d "%systemroot%\system32\imageres.dll,197" /t reg_sz /f > nul
taskkill /f /im explorer.exe > nul
start explorer > nul
echo ȥ����ݷ�ʽ���Ʋ������...
pause
cls
goto start


:recoveshield
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons" /v 77 /f > nul
taskkill /f /im explorer.exe > nul
start explorer > nul
echo �ָ���ݷ�ʽ���Ʋ������...
pause
cls
goto start


:removerunwarn
reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Associations /v ModRiskFileTypes /t REG_SZ /d .exe;.bat;.vbs;.py;.cmd;.msi;.ps1;.js /f
gpupdate /force
echo ȥ����ִ���ļ��İ�ȫ���浯���������...
pause
cls
goto start

:addmypcico
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel" /v "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" /t REG_DWORD /d "0" /f
taskkill /f /im explorer.exe > nul
start explorer > nul
echo ��������ӡ��˵��ԡ�ͼ��������...
pause
cls
goto start