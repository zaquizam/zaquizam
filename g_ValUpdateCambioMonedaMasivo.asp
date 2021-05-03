<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	'
	' g_ValUpdateCambioMonedaMasivo.asp - 23feb21 - 01mar21
	'
	Dim updSql	
	'
	Session.lcid		= 2057
	Response.CodePage	= 65001
	Response.CharSet	= "utf-8"
    '
    ' Capturar las variables
    '	
	idConsumo	= Request.QueryString("idConsumo")
	idMoneda	= Request.QueryString("idMoneda")	
	Moneda		= Request.QueryString("Moneda")	
	idSemana	= Request.QueryString("idSemana")
	'
	Dim QrySql, TasadeCambio 
	'
	' Buscar la tasa de cambio de la Semana segun la fecha de consumo
	'		
	QrySql = vbnullstring
	QrySql = " SELECT clave_busqueda FROM ph_moneda WHERE id_moneda = " & idMoneda
	Set rsSql = Server.CreateObject("ADODB.recordset")
	'
	rsSql.Open QrySql, conexion
	if not (rsSql.EOF and rsSql.BOF) then
		TipoMoneda = rsSql(0)
	end if
	'
	rsSql.close
	set rsSql= Nothing	
	'
	if TipoMoneda = "bolivar" then 
		TasadeCambio=1		
	else
		QrySql = vbnullstring
		QrySql = " SELECT " & TipoMoneda  & " FROM ss_semana WHERE idsemana = " & idSemana
		Set rsSql = Server.CreateObject("ADODB.recordset")
		rsSql.Open QrySql, conexion
		if not (rsSql.EOF and rsSql.BOF) then
			TasadeCambio = rsSql(0)
			TasadeCambio = replace(TasadeCambio,",",".")
		else
			TasadeCambio=1
		end if			
		rsSql.close
		set rsSql= Nothing	
		'
	end if
	'
    ' Actualizar Tasa de Cambio 1ero....
    '
    updSql = vbnullstring
	updSql = updSql & " UPDATE PH_Consumo_Detalle_Productos"
    updSql = updSql & " SET"	
	updSql = updSql & " tasa_de_cambio = " & CDBl(TasadeCambio) 
    updSql = updSql & " WHERE"
    updSql = updSql & " Id_Consumo=" & idConsumo
    '		
    Set objExec = conexion.Execute(updSql)
    Set objExec = Nothing
    '
    ' Actualizar Datos Validando....
    '
    updSql = vbnullstring
	updSql = updSql & " UPDATE PH_Consumo_Detalle_Productos"
    updSql = updSql & " SET"
	updSql = updSql & " id_moneda = "  & idMoneda & ","
	updSql = updSql & " moneda    ='"  & Moneda   & "',"
	'
	' Activado 01mar21
	'	
	updSql = updSql & " total_compra = (Precio_producto * tasa_de_cambio ) * cantidad"
    updSql = updSql & " WHERE"
    updSql = updSql & " Id_Consumo=" & idConsumo
    '	
	' response.write updSql
	' response.end
	'
    Set objExec = conexion.Execute(updSql)
    Set objExec = Nothing
    '		
    If Err.Number = 0 Then
		Response.Write True
	Else
		Response.write "error" 
	End If		
	'
	conexion.Close
    Set objExec = Nothing	
    Set conexion = Nothing
	'
%>