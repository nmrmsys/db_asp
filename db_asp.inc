<%@ CodePage=65001 %>
<%
'  db_asp - db stub for classic asp
'
'  サーバ設定など http://technet.microsoft.com/ja-jp/library/ee155445.aspx
'  64bit環境用 http://niyodiary.cocolog-nifty.com/blog/2009/05/net-framework64.html#.A20090502_1_3

Response.Charset     = "utf-8"
Response.ContentType = "text/plain"

Set Prm = Request.QueryString
Set Frm = Request.Form
Set Env = Request.ServerVariables

Set FSO = Server.CreateObject("Scripting.FileSystemObject")
Set DBC = Server.CreateObject("ADODB.Connection")

'MDB接続文字列
DBC.Open "Provider=Microsoft.Jet.OLEDB.4.0" & _
         ";Data Source=" & Server.MapPath("db_asp.mdb")

'SQL Server接続文字列
'DBC.Open "Provider=SQLOLEDB" & _
''         ";Data Source=sv ;Initial Catalog=db ;User ID=usr ;Password=pwd"

'クエリ文字列からタブ改行区切りの結果テキスト出力 or SQL実行&処理件数出力
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

'クエリ文字列からJSONの結果テキスト出力 or SQL実行&処理件数出力
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

'特殊記号をエスケープして出力
Sub Output(argSTR)
  sSTR = argSTR
'  sSTR = Replace(sSTR, "&", "&amp;")
'  sSTR = Replace(sSTR, """", "&quot;")
'  sSTR = Replace(sSTR, "<", "&lt;")
'  sSTR = Replace(sSTR, ">", "&gt;")
  Response.Write sSTR
End Sub

'JS文字列でエスケープ必要な文字を変換
Function JsString(argValue)
  sTmp = argValue
  sTmp = Replace(sTmp, "'", "\'")
  sTmp = Replace(sTmp, vbCrLf, "\n")
  sTmp = Replace(sTmp, vbCr, "\n")
  sTmp = Replace(sTmp, vbLf, "\n")
  JsString = sTmp
End Function

'NVLヘルパー
Function Nvl(argValue,argNullValue)
  If IsNull(argValue) Then
    Nvl = argNullValue
  Else
    Nvl = argValue
  End If
End Function

'入力文字列をクエリ用にエスケープ
Function Input(argSTR)
  sSTR = argSTR
  sSTR = Replace(sSTR, "'", "''")
  Input = sSTR
End Function

%>
