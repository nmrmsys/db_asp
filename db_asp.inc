<%@ CodePage=65001 %>
<%

Response.Charset     = "utf-8"
Response.ContentType = "text/plain"

Set Prm = Request.QueryString
Set Frm = Request.Form
Set Env = Request.ServerVariables

Set FSO = Server.CreateObject("Scripting.FileSystemObject")
Set DBC = Server.CreateObject("ADODB.Connection")

'DBC.Open "DSN=SV11;UID=SALASEE;PWD=SALASEE"
DBC.Open "Provider=Microsoft.Jet.OLEDB.4.0" & _
         ";Data Source=" & Server.MapPath("db_asp.mdb")

Sub Output(argSTR)
  sSTR = argSTR
'  sSTR = Replace(sSTR, "&", "&amp;")
'  sSTR = Replace(sSTR, """", "&quot;")
'  sSTR = Replace(sSTR, "<", "&lt;")
'  sSTR = Replace(sSTR, ">", "&gt;")
  Response.Write sSTR
End Sub

Function Input(argSTR)
  sSTR = argSTR
  sSTR = Replace(sSTR, "'", "''")
  Input = sSTR
End Function

Function dbExecute(argSQL)
  sRt = ""
  If UCase(Mid(Trim(argSQL),1,6)) = "SELECT" Then
    Set Rs = DBC.Execute(argSQL)
    nFs = Rs.Fields.Count - 1
    For n = 0 To nFs
      If n <> 0 Then
        sRt = sRt & vbTab
      End If
      sRt = sRt & Rs.Fields(n).Name
    Next
    Do Until Rs.EOF
      sRt = sRt & vbCrLf
      For n = 0 To nFs
        If n <> 0 Then
          sRt = sRt & vbTab
        End If
        sRt = sRt & Nvl(Rs(n).value,"")
      Next
      Rs.MoveNext
    Loop
    dbExecute = sRt
  Else
    Call DBC.Execute(argSQL,nCNT)
    sRt = sRt & nCNT
    dbExecute = sRt
  End If
End Function

Function dbExecuteJSON(argName, argSQL)
'  Response.ContentType = "application/json"
  sRt = ""
'  sRt = sRt & "'" & argName & "_SQL':'" & JsString(argSQL)  & "'," & vbCrLf
  If UCase(Mid(Trim(argSQL),1,6)) = "SELECT" Then
    Set Rs = DBC.Execute(argSQL)
    nFs = Rs.Fields.Count - 1
    sRt = sRt & "  '" & argName & "':{" & vbCrLf
    sRt = sRt & "    'rows':[" & vbCrLf
    Do Until Rs.EOF
      sRt = sRt & "      {"
      For n = 0 To nFs
        If n <> 0 Then
          sRt = sRt & ","
        End If
        sRt = sRt & "'" & Rs.Fields(n).Name & "':'" & JsString(Nvl(Rs(n).value, ""))  & "'"
      Next
      sRt = sRt & "}," & vbCrLf
      Rs.MoveNext
    Loop
    sRt = sRt & "    ]" & vbCrLf
    sRt = sRt & "  }," & vbCrLf
    dbExecuteJSON = sRt
  Else
    Call DBC.Execute(argSQL,nCNT)
    sRt = sRt & "  '" & argName & "':" & nCNT & "," & vbCrLf
    dbExecuteJSON = sRt
  End If
End Function

Function Nvl(argValue,argNullValue)
  If IsNull(argValue) Then
    Nvl = argNullValue
  Else
    Nvl = argValue
  End If
End Function

Function JsString(argValue)
  sTmp = argValue
  sTmp = Replace(sTmp, "'", "\'")
  sTmp = Replace(sTmp, vbCrLf, "\n")
  sTmp = Replace(sTmp, vbCr, "\n")
  sTmp = Replace(sTmp, vbLf, "\n")
  JsString = sTmp
End Function

%>
