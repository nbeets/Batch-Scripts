Option Explicit 

dim args, strComputer, ObjRegistry,strPath, Return, oReg, strKeyPath, strUser
'Get arguments
args = WScript.Arguments.Count

If args < 1 then
  WScript.Echo "usage: createhidden.vbs %username%"
  WScript.Quit
end If


'On Error Resume Next
'Hide User Account
dim HKEY_LOCAL_MACHINE

HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."

Set ObjRegistry = _
    GetObject("winmgmts:{impersonationLevel = impersonate}!\\" _
    & strComputer & "\root\default:StdRegProv")

strPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList"

Return = objRegistry.CreateKey(HKEY_LOCAL_MACHINE, strPath)

Set oReg=GetObject( _
    "winmgmts:{impersonationLevel=impersonate}!\\" &_ 
    strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList"
oReg.SetDWORDValue _ 
    HKEY_LOCAL_MACHINE,strKeyPath,WScript.Arguments.Item(0),0


WScript.Echo "Username " & WScript.Arguments.Item(0) & " has been hidden." 
WScript.Quit