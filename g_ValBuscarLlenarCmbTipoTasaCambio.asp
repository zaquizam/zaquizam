<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
' g_ValBuscarLlenarCmbTipoTasaCambio.asp // 12ene21 - 07abr21
'
Session.lcid =1 034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'
Dim QrySql, rsSql
'
' Buscar la tasa de cambio de la Semana segun la fecha de consumo
'
idMoneda = Request.Form("idmoneda")
idSemana = Request.Form("id_semana")		
'
' Buscar el campo moneda busqueda para saber cual es la tasa tomar segun la moneda
'
QrySql = vbnullstring
QrySql = " SELECT clave_busqueda FROM PH_Moneda WHERE Id_Moneda = " & idMoneda
Set rsSql = Server.CreateObject("ADODB.recordset")
'
rsSql.Open QrySql, conexion
'
if not (rsSql.EOF and rsSql.BOF) then
   tipoMoneda = rsSql(0)
end if
'
'Response.Write QrySql & " " & tipoMoneda
'Response.end
'
rsSql.close
set rsSql= Nothing	
'
QrySql = vbnullstring
QrySql = " SELECT " & tipoMoneda  & " FROM ss_semana WHERE idsemana = " & idSemana
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