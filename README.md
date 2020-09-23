<div align="center">

## RunDosCommand Function


</div>

### Description

Run MS DOS Command and get results back in VB.NET
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Igor Krupitsky](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/igor-krupitsky.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB\.NET
**Category**       |[System Services/ Functions](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/system-services-functions__10-23.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/igor-krupitsky-rundoscommand-function__10-5679/archive/master.zip)





### Source Code

```
'Example Of Use
'sRet = RunDosCommand("CScript.exe C:\MyScript.vbs", 10)
Function RunDosCommand(ByVal sCommandText As String, Optional ByVal iTimeOutSec As Integer = 1) As String
	Dim iPos As Integer = sCommandText.IndexOf(" ")
	Dim sFileName As String
	Dim sArguments As String = ""
	If iPos = -1 Then
		sFileName = sCommandText
	Else
		sFileName = sCommandText.Substring(0, iPos)
		sArguments = sCommandText.Substring(iPos + 1)
	End If
	Dim sRet As String
	Dim oProcess As Process = New Process
	oProcess.StartInfo.UseShellExecute = False
	oProcess.StartInfo.RedirectStandardOutput = True
	oProcess.StartInfo.FileName = sFileName
	oProcess.StartInfo.Arguments = sArguments
	oProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
	oProcess.StartInfo.CreateNoWindow = True
	oProcess.Start()
	oProcess.WaitForExit(1000 * iTimeOutSec)
	If Not oProcess.HasExited Then
		oProcess.Kill()
		Return "Timeout"
	End If
	sRet = oProcess.StandardOutput.ReadToEnd()
	If oProcess.ExitCode <> 0 And sRet = "" Then
		sRet = "ExitCode: " & oProcess.ExitCode
	End If
	oProcess.Close()
	Return sRet
End Function
```

