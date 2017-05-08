' SetProxy.vbs
' VBScript to set correct proxy server 
' Author Luiz.Lourenco@outlook.com.br or Luiz.Lourenco@dhl.com   
' Version 1.1 - 08-May-2017
' ------------------------------------------------' 

' Global Variable Declare
Option Explicit 
Dim strComputer, objWMIService, colOperatingSystems, objOperatingSystem, WindowsVersionName, oShell, TextPrompt, TextTitle


' Set text of variables to can show
TextTitle  = "Internet Proxy"  
TextPrompt = "Caro usu"&ChrW(0225)&"rio," & vbNewLine & _
             "Voc"&ChrW(0234)&" deseja acessar a Internet na rede DHL ou usando a Cisco VPN?" & vbNewLine & _
             "Se afirmativo, clique em Yes" & vbNewLine & vbNewLine & _
             "-----------------------" & vbNewLine & vbNewLine & _
             "Dear User," & vbNewLine & _
             "Would you like to access the internet within the DHL network or using the Cisco VPN?" & vbNewLine & _
             "If so, please click the Yes button"

' Set variables as object to Wscript.Shell
set oShell = Wscript.CreateObject("Wscript.Shell")

'----------------------------------------------------------------------------
' Main Program
'----------------------------------------------------------------------------

GetOSName() 'Get of Operating System which who version that is running.

'Ask to user if like user internet intro company network'
If MsgBox(TextPrompt, vbQuestion or vbYesNo,TextTitle) = vbYes then
    If (WindowsVersionName = "Windows 7") Then
        Wscript.echo "Voce esta usando Microsoft Windows 7 Enterprise"
        SetScriptProxy("http://localhost/proxy.pac")
    End If
    If (WindowsVersionName = "Windows 10") Then
        Wscript.echo "Voce esta usando Microsoft Windows 10 Enterprise"
        SetScriptProxy("http://localhost/proxy.pac")
    End If
Else
    DropScriptProxy()
End If



'----------------------------------------------------------------------------
'SubRoutine to get of Operating System which who version that is running.
'----------------------------------------------------------------------------
Sub GetOSName() 
    strComputer = "." 
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
    Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem") 
    For Each objOperatingSystem in colOperatingSystems 
    Select Case Trim(objOperatingSystem.Caption)
        Case "Microsoft Windows 7 Enterprise" ' My version
            'wsh.echo "Voce esta usando Microsoft Windows 7 Enterprise"
            WindowsVersionName = "Windows 7"

        Case "Microsoft Windows 10 Enterprise" ' My version
            'wsh.echo "Voce esta usando Microsoft Windows 10 Enterprise" 
            WindowsVersionName = "Windows 10"
    End Select 
    Next
End Sub


'----------------------------------------------------------------------------
'SubRoutine to set the proxy address
'----------------------------------------------------------------------------
Sub SetScriptProxy(ProxyAddress) 
    oShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoConfigURL", ProxyAddress, "REG_EXPAND_SZ"
    SetScriptProxyOn()
End Sub

'----------------------------------------------------------------------------
'SubRoutine to set the proxy address as enable
'----------------------------------------------------------------------------
Sub SetScriptProxyOn() 
    oShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"
End Sub

'----------------------------------------------------------------------------
'SubRoutine to Drop the proxy setting to internet is done directly
'----------------------------------------------------------------------------
Sub DropScriptProxy() 
    SetScriptProxy("http://localhost/proxy.pac") ' I'm set any information to proxy address, to don't show a error message if has twice execution and selected button is NO
    oShell.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoConfigURL"
End Sub


' Emptying all variable 
Set oShell = Nothing
Set strComputer = Nothing
Set objWMIService = Nothing
Set colOperatingSystems = Nothing
Set objOperatingSystem = Nothing
Set WindowsVersionName = Nothing
Set TextPrompt = Nothing
Set TextTitle = Nothing