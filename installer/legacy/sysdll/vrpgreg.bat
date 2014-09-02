@echo unregistering DLL stuff
regsvr32 /u msvbvm50.dll
regsvr32 /u msvbcm60.dll
regsvr32 /u mpqctl.ocx
regsvr32 /u DMC2.ocx
regsvr32 /u dx7vb.dll

@echo registering DLL stuff
regsvr32 msvbvm50.dll
regsvr32 msvbcm60.dll
regsvr32 mpqctl.ocx
regsvr32 DMC2.ocx
regsvr32 dx7vb.dll
