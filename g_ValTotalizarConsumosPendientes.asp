<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_ValTotalizarConsumosPendientes.asp // 14ene21 - 
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim rsPendientes, rsArray, idSemana, strSQL, intRow
	'	
	idSemana	=	Request.Querystring("id_Semana")
	'
	' Buscar total de cosnumos pendientes por investigar
	'	
	Set rsPendientes = CreateObject("ADODB.Recordset")	
	'	
	strSQL = vbnullstring	
	strSQL = " SELECT COUNT(PH_Consumo_Detalle_Productos.Pendiente) AS TotalPendientes"
	strSQL = strSQL & " FROM"
	strSQL = strSQL & " PH_Consumo_Detalle_Productos"
	strSQL = strSQL & " INNER JOIN PH_Consumo ON"
	strSQL = strSQL & " PH_Consumo_Detalle_Productos.Id_Consumo = PH_Consumo.Id_Consumo"
	strSQL = strSQL & " WHERE"
	strSQL = strSQL & " PH_Consumo_Detalle_Productos.Pendiente = 1"
	strSQL = strSQL & " AND"
	strSQL = strSQL & " PH_Consumo_Detalle_Productos.Resuelto = 0"
	strSQL = strSQL & " AND"
	strSQL = strSQL & " PH_Consumo.Id_Semana = " & idSemana
	''
	rsPendientes.open strSQL, conexion	
	'
	If not rsPendientes.EOF  Then
		' rsArray = rsPendientes.GetRows() 
        ' intRow = UBound(rsArray, 2) + 1 
        response.write rsPendientes(0)
	Else
		Response.write 0
	End If
	'
	rsPendientes.close : set rsPendientes = nothing 
	conexion.Close : Set conexion = Nothing
	'
%>