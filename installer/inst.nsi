Name "VRPG2"

OutFile "../bin/VRPG214ModSetup.exe"

RequestExecutionLevel admin

Page license

LicenseData "license.txt"

Page components

ComponentText "Depending on the target operating system, you may not require to install all the files" "Pick up your OS :" "Select more or less components :"
InstType "windows 98"
InstType "windows 2000 or XP"
InstType "windows Vista or 7"

Page directory
Page instfiles

Section "Game files"
  SectionIn 1 2 3
  SetOutPath $INSTDIR
  File "Data.pak"
  File "..\bin\VRPG2_2014-01-13.exe"
  File "..\bin\MPQExtractor.exe"
SectionEnd

Section "VB runtime DLL files"
  SectionIn 3
  SetOutPath $SYSDIR
  SetOverwrite off
  File "msvbvm60.dll"
  RegDLL "msvbvm60.dll"
  File "VB6STKIT.dll"
  #RegDll "VB6STKIT.dll"
SectionEnd

Section "Game OCX files (win98/XP)"
  SectionIn 1 2
  SetOutPath $INSTDIR
  File "MpqCtl.ocx"
SectionEnd

Section "Game OCX files (Vista/7)"
  SectionIn 3
  SetOutPath $SYSDIR
  File "MpqCtl.ocx"
  RegDll "MpqCtl.ocx"
SectionEnd

Section "Vista and 7 missing DLLs"
  SectionIn 3
  SetOutPath $SYSDIR
  SetOverwrite off
  File "COMDLG32.OCX"
  RegDll "COMDLG32.OCX"
  File "d3drm.dll"
  File "dx7vb.dll" 
  RegDLL "dx7vb.dll"
  RegDLL "d3drm.dll"
 SectionEnd
 
Section "Cheat codes list"
  SetOutPath $INSTDIR
  SetOverwrite ifdiff
  File "_cheats.txt"
SectionEnd

Section "Mods (Throku's change)"
  SetOutPath $INSTDIR
  SetOverwrite ifdiff
  File "ExVRPG-Throku-mod.bat"
  CreateDirectory $INSTDIR\Throkuchanges2
  SetOutPath $INSTDIR\Throkuchanges2
  File /r "..\src\Throkuchanges2\*.*"
SectionEnd

UninstPage uninstConfirm
UninstPage instfiles
