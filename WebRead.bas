Attribute VB_Name = "WebRead"
Option Explicit

Public Function getHtmlCode$(strURL$) '获取源码
On Error GoTo reStart
reStart:
DoEvents
Dim stime, ntime
Dim XmlHttp
Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
'MsgBox strURL
XmlHttp.Open "GET", strURL, True
XmlHttp.SetRequestHeader "If-Modified-Since", "0"
XmlHttp.Send
stime = Now '获取当前时间
While XmlHttp.ReadyState <> 4
DoEvents
ntime = Now '获取循环时间
If DateDiff("s", stime, ntime) > 13 Then getHtmlCode = "OutTime": Exit Function '判断超出3秒即超时退出过程
Wend

getHtmlCode = StrConv(XmlHttp.ResponseBody, vbUnicode)
'MsgBox getHtmlCode
If getHtmlCode = "" Then getHtmlCode = "OutTime"
'XmlHttp.Close
Set XmlHttp = Nothing
DoEvents
Exit Function
errline:
  MsgBox "发生未知错误，请重试！", vbCritical
End Function

Public Function StringCheck(Str_Get As String) 'Web字符串检查
    Dim injdata, inj, i
    injdata = "' * % [ ] & # ? ^ / \ !"
    inj = Split(injdata, " ")
    For i = 0 To UBound(inj)
        If InStr(Str_Get, inj(i)) > 0 Then
            StringCheck = True
        End If
    Next
End Function

Public Function PassCodeC(UC)
  '  PassCodeC = Int((UC + 1748) * 0.3 * 1.4)
End Function

