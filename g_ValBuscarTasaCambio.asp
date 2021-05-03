<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
' g_ValBuscarTasaCambio.asp //13ene21 -
'
Session.lcid = 1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'
Dim QrySql 
'
' Buscar la tasa de cambio de la Semana segun la fecha de consumo
'
idMoneda = Request.Form("id_moneda")
idSemana   = Request.Form("id_semana")		
'
QrySql = vbnullstring
QrySql = " SELECT clave_busqueda FROM ph_moneda WHERE id_moneda = " & idMoneda
Set rsSql = Server.CreateObject("ADODB.recordset")
'
'Response.Write QrySql
'Response.end
'
rsSql.Open QrySql, conexion
'
if not (rsSql.EOF and rsSql.BOF) then
	TipoMoneda = rsSql(0)
end if
'
rsSql.close
set rsSql= Nothing	
'//
QrySql = vbnullstring
QrySql = " SELECT " & TipoMoneda  & " FROM ss_semana WHERE idsemana = " & idSemana
Set rsSql = Server.CreateObject("ADODB.recordset")
'
'Response.Write QrySql
'Response.end
'
rsSql.Open QrySql, conexion
'
if not (rsSql.EOF and rsSql.BOF) then
	TasadeCambio = rsSql(0)
	TasadeCambio = replace(TasadeCambio,",",".")
else
	TasadeCambio=1
end if
'
rsSql.close
set rsSql= Nothing	

If Err.Number = 0 Then
	Response.write TasadeCambio
Else
	Response.write  Err.Description
End If		    
'
conexion.Close    
Set conexion = Nothing
'
%>