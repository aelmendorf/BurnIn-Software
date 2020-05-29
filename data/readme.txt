For release 1.61 (timeout feature):

NatUSB.dll is the same as for release 1.6
NatUSB_32_ezusb.dll is the same as NatUSB.dll
NatUSB_32_winusb.dll <-- %svn%/uniusb/trunk/UniUSB/src/NatUSB/win64/NatUSB/release
NatUSB_64.dll        <-- %svn%/uniusb/trunk/UniUSB/src/NatUSB/win64/NatUSB/x64/release


--------------------------------------


For release 1.6:

The following DLL files were obtained from the following location:

svn export %svn%/spectrasuite/trunk/Jars
NatDL_32.dll
NatDL_64.dll
NatHRTiming_32.dll
NatHRTiming_64.dll
NatUSB_64.dll

---

NatUSB_32.dll comes from the svn project %svn%/uniusb/trunk/UniUSB/src/NatUSB/winusb_32bit
NatUSB.dll is simply a copy of NatUSB_32.dll

---

The script to create the directory tree 
(that is ready to be used as input
to Advanced Installer)
must choose either the 32-bit or
64-bit variants of each of these files
and rename them by removing the
trailing "_32" or "_64",
depending on whether you are
targeting a 32-bit operating system
or a 64-bit operating system.

