<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
'
' g_pPendCerrarProductoPendiente.asp - 15mar21 - 26mar21
'
Session.lcid      = 1034
Response.CodePage = 65001
Response.CharSet  = "utf-8"
'
Dim updSql, rsSql, barcode, sIP
Dim dd, mm, yy, hh, nn, ss
Dim datevalue, timevalue, dtsnow
'
' Capturar las variables
'
barcode = Request.QueryString("barcode")
sIP     = Request.ServerVariables("REMOTE_ADDR")
'
' Buscar Id de la Semana segun la fecha de validacion del Producto 
'    
dtsnow = Now()
dd = Right("00" & Day(dtsnow), 2)
mm = Right("00" & Month(dtsnow), 2)
yy = Year(dtsnow)
hh = Right("00" & Hour(dtsnow), 2)
nn = Right("00" & Minute(dtsnow), 2)
ss = Right("00" & Second(dtsnow), 2)
datevalue = yy & "-" & mm & "-" & dd
timevalue = hh & ":" & nn & ":" & ss
sUpdate = datevalue & " " & timevalue    
'	
QrySql = vbnullstring
QrySql = QrySql & " SELECT idsemana FROM ss_semana  WHERE '" & datevalue & "' BETWEEN fec_inicio AND fec_fin"
'    
Set rsSql = Server.CreateObject("ADODB.recordset")
rsSql.Open QrySql, conexion
'
if not (rsSql.EOF and rsSql.BOF) then
    idSemana=rsSql(0)
else
    idSemana=0
end if
'
rsSql.close : set rsSql= Nothing	    
'
' Validar todos los codigo de barras ....
'
updSql = vbnullstring
updSql = updSql & " UPDATE PH_Consumo_Detalle_Productos"
updSql = updSql & " SET"
updSql = updSql & " Validado  = 1 ,"
updSql = updSql & " Pendiente = 0 ,"
updSql = updSql & " Resuelto  = 0 ,"    
updSql = updSql & " idSemanaValidacion=" & idSemana & ","
'updSql = updSql & " Fec_Inactivo='" & sUpdate & "',"
updSql = updSql & " USR ='" & Session("Usuario") & "',"
updSql = updSql & " IP='" & sIP & "'"    
updSql = updSql & " WHERE"
updSql = updSql & " Status_registro='G'"
updSql = updSql & " AND"
updSql = updSql & " Numero_codigo_barras='" & barcode & "'"
'
' Response.Write updSql
' Response.end
'
Set objExec = conexion.Execute(updSql)
'
If Err.Number = 0 Then
    '
    updSql = vbnullstring
    updSql = updSql & " UPDATE PH_CB_Producto"
    updSql = updSql & " SET"
    updSql = updSql & " Ind_pendiente  = 0,"    
    updSql = updSql & " Fec_Inactivo='" & sUpdate & "',"
    updSql = updSql & " USR ='" & Session("Usuario") & "',"
    updSql = updSql & " IP='" & sIP & "'"    
    updSql = updSql & " WHERE"
    updSql = updSql & " CodigoBarra='" & barcode & "'"
    '
    ' Response.Write updSql
    ' Response.end
    '
    Set objExec = conexion.Execute(updSql)    
    '
    Response.Write True
else
    Response.Write False
end if
'
Set objExec = Nothing
conexion.Close : Set conexion = Nothing
'
%>