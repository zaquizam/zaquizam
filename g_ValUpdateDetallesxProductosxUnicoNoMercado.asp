<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValUpdateDetallesxProductosxUnicoNoMercado.asp // 03ene21 - 01feb21
	'
	Dim updSql	
	'
	Session.lcid = 2057
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'
	' Dim dd, mm, yy, hh, nn, ss, datevalue, dtsnow
    '
    ' Capturar las variables
    '	
	idConsumo		= Request.QueryString("idConsumo")
	idMoneda		= Request.QueryString("idMoneda")
	TotalCompra     = Request.QueryString("total")
	idtipocomida	= Request.QueryString("idtipocomida")
	nombreLocal		= Request.QueryString("nombreLocal")	
    '
    ' Actualizar Datos Validando....
    '
    updSql = vbnullstring
	updSql = updSql & " UPDATE PH_Consumo"
    updSql = updSql & " SET"
    updSql = updSql & " Nombre_local ='" & nombreLocal & "',"
	updSql = updSql & " id_moneda=" 	 & idMoneda & ","
	updSql = updSql & " id_tipoComida="  & idtipocomida & ","
	updSql = updSql & " Total_Compra=" 	 & TotalCompra & ","
    updSql = updSql & " Validado='1',"
	updSql = updSql & " Resuelto='0'"
	'
    updSql = updSql & " WHERE"
    updSql = updSql & " Id_Consumo=" & idConsumo
    '
    'Response.Write updSql
	'Response.end
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