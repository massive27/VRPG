VRPG2 - winvista/win7 compatibility tricks.


Those are the mandatory DLL's files to launch the game.

dx7vb.dll
d3drm.dll
msvbvm50.dll
msvbvm60.dll *
DMC2.ocx
MpqCtl.ocx
asycfilt.dll *

* : those files usually are already included with Windows. Don't overwrite them.



== Where to put them :

If you're using windows vista or 7, 32-bits, any edition, copy them in your windows\system32 directory (don't overwrite if they already exists)

If you're using windows vista or 7, 64 bits, any edition, copy them in windows\syswow64 directory (don't overwrite if they already exists). You WILL need administrator privilege to do so.



== Troubleshoot :

If the game doesn't launch (says it miss one or more files, before the main screen even appear) put all those file in the game directory.


If the game still doesn't launch afterward, try executing the little vrpgreg.bat included.
Execute it with administrative rights, or in safemode. Ignore (yet) if there is error messages. That file must be in the same directory as the various dlls files !


== Side notes :
> The DMC2.ocx error means that msvbvm50.dll and DMC2.ocx arent in the same directory.
> The 430 error and/or CreateFromFile error means d3drm.dll or dx7vb.dll is missing in the windows\system32 directory.
> The msvbvm60.dll missing error means msvbvm60.dll is... missing. Copy it either in game directory or windows\system32. Be sure to have administrator privilege.
> If the game freeze at main screen, the little bass.dll in the game directory isn't compatible, try finding another one.


Check Eka's portal forum, Duamutef's dimension, for further help !