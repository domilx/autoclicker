taskKill = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & taskKill & "\root\cimv2")

Set processlistFull= objWMIService.ExecQuery _
("Select * from Win32_Process Where name = 'wscript.exe'")

For Each killPro In processlistFull
killPro.Terminate()
Next

taskKill = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & taskKill & "\root\cimv2")

Set processlistFull= objWMIService.ExecQuery _
("Select * from Win32_Process Where name = 'wscript.exe'")

For Each killPro In processlistFull
killPro.Terminate()
Next

  