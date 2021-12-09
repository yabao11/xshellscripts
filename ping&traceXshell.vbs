
'
' @Description: 
' @Author: YB11
' @GitHub: https://github.com/yabao11
' @Date: 2021-12-09 13:15:03
' @LastEditors: YB11
' @LastEditTime: 2021-12-09 17:01:21
' @FilePath: \菱信资料\xshell_pings.vbs
'xsh.Session.Sleep 100


Sub Main
	Dim fso,fsw, fConfig, file, path
    	path = xsh.Dialog.Prompt ("Your Xshell path", "Prompt Dialog", "", 0)
	Set fso = CreateObject("Scripting.FileSystemObject") 
	Set fConfig = fso.OpenTextFile(path & "ip.txt", 1)
    	xsh.Screen.Synchronous = True
   	xsh.Screen.send vbcr
	xsh.Screen.waitforstring("#")
Do While fConfig.AtEndOfStream <> True
	lo = fConfig.ReadLine
	xsh.Dialog.MsgBox(lo)
   	xsh.Screen.send "ping " & " " & lo & vbcr
  	xsh.Screen.waitforstring("#")
    	xsh.Screen.send vbcr
    	xsh.Screen.send "trace" & " " & lo & " ttl 3 10"& vbcr
    	xsh.Screen.waitforstring("#")
    	xsh.Screen.send vbcr
	xsh.Session.Sleep(2000)
	loop
    xsh.Dialog.MsgBox("执行完毕！")
End Sub
