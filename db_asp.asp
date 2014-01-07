<!--#include file="db_asp.inc"-->
<%

Select Case Prm("cmd")
Case "tsv"
  Call OutputTsv
Case "json"
  Call OutputJson
Case Else
  Call UsagePage
End Select

Sub OutputTsv
  sSQL = "SELECT * FROM CODE_M"
  Output(vbCrLf)
  Output(dbExecute(sSQL))
  Output(vbCrLf)
End Sub

Sub OutputJson
  sSQL = "SELECT * FROM CODE_M"
  Output("{" & vbCrLf)
  Output(dbExecuteJSON("QRY1",sSQL))
  Output("}" & vbCrLf)
End Sub

Sub UsagePage
  Output(vbCrLf)
  Output("  db_asp.asp?cmd=<tsv, json>")
  Output(vbCrLf)
End Sub

%>
