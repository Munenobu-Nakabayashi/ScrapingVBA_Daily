@echo off

rem chcp 65001

setlocal enabledelayedexpansion

pushd %~dp0
set folderpath=%~dp0

if not exist .\error.log (
        type nul > .\error.log
)

echo ---Start--- >> .\error.log
echo %date% %time% >> .\error.log
echo. >> .\error.log

echo "scraping_naikakukambou.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_naikakukambou.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_naikakukambou.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_US_CERT.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_US_CERT.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_US_CERT.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_Ubuntu.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_Ubuntu.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_Ubuntu.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_TrendMicro.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_TrendMicro.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_TrendMicro.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_tenable.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_tenable.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_tenable.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_SOMPO_Glossary.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_SOMPO_Glossary.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_SOMPO_Glossary.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_SOMPO.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_SOMPO.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_SOMPO.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_SIOS_valunerability.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_SIOS_valunerability.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_SIOS_valunerability.xlsm is cloed." >> .\error.log
echo. >> .\error.log

echo "scraping_SecurityNext2nd.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_SecurityNext2nd.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_SecurityNext2nd.xlsm is cloed." >> .\error.log
echo. >> .\error.log

echo "scraping_ScanNetSecurity.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_ScanNetSecurity.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_ScanNetSecurity.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_proofpoint.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_proofpoint.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_proofpoint.xlsm is cloed." >> .\error.log
echo. >> .\error.log

echo "scraping_OracleUnbreakableLinuxNetwork.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_OracleUnbreakableLinuxNetwork.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_OracleUnbreakableLinuxNetwork.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_NISC_Doc.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_NISC_Doc.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_NISC_Doc.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_NISC.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_NISC.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_NISC.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_Nikkei.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_Nikkei.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_Nikkei.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_Microsoft_JapanSecurityTeam.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_Microsoft_JapanSecurityTeam.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_Microsoft_JapanSecurityTeam.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_madonomori.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_madonomori.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_madonomori.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_JVNDB.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_JVNDB.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_JVNDB.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_JVN_iPedia.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_JVN_iPedia.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_JVN_iPedia.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_JVN.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_JVN.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_JVN.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_JPCERT.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_JPCERT.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_JPCERT.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_JPA.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_JPA.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_JPA.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_JIPDEC.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_JIPDEC.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_JIPDEC.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_jc3.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_jc3.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_jc3.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_IPA.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_IPA.xlsm" "sheet1.main" >> .\error.log  2>&1
echo "scraping_IPA.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_digitalforensic.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_digitalforensic.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_digitalforensic.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_debian.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_debian.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_debian.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_Cisco.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_Cisco.xlsm" "sheet1.main" >> .\error.log  2>&1
echo "scraping_Cisco.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_amazonlinux.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_amazonlinux.xlsm" "sheet1.main" >> .\error.log  2>&1
echo "scraping_amazonlinux.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_act1_cybersecurity_news.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_act1_cybersecurity_news.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_act1_cybersecurity_news.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo "scraping_Cisco_thousandeyes.xlsm will be opened." >> .\error.log
cscript .\runVBS.vbs %folderpath%"\scraping_Cisco_thousandeyes.xlsm" "sheet1.main" >> .\error.log 2>&1
echo "scraping_Cisco_thousandeyes.xlsm is closed." >> .\error.log
echo. >> .\error.log

echo. >> .\error.log
echo %date% %time% >> .\error.log
echo ----End---- >> .\error.log
rem //空行を追記
echo. >> .\error.log

popd

exit 0