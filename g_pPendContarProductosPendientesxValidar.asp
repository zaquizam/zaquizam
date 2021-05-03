<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%	
	'
	' g_pPendContarProductosPendientesxValidar.asp - 03mar21
	'
	Session.lcid		= 1034
	Response.CodePage 	= 65001
	Response.CharSet 	= "utf-8"
	'
	idCodigoBarras      = Request.Form("id")	
	idStatus      		= Request.Form("status")	
	'
	Dim rsSql, total
	'	
	' Buscar Los Productos Pendientes por info completa del codigo de Barras
	'	
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT COUNT"
	QrySql = QrySql & " (PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) as total"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo_Detalle_Productos"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo_Detalle_Productos.status_registro='G'" 
	QrySql = QrySql & " AND PH_Consumo_Detalle_Productos.id_hogar>1" 
	'
	QrySql = QrySql & " AND PH_Consumo_Detalle_Productos.Precio_producto is not null"
	QrySql = QrySql & " AND PH_Consumo_Detalle_Productos.tasa_de_cambio is not null"
	QrySql = QrySql & " AND PH_Consumo_Detalle_Productos.Cantidad is not null"
	QrySql = QrySql & " AND PH_Consumo_Detalle_Productos.Id_Moneda is not null"
	QrySql = QrySql & " AND PH_Consumo_Detalle_Productos.Moneda is not null"	
	'
	QrySql = QrySql & " AND PH_Consumo_Detalle_Productos.Numero_codigo_barras = '" & idCodigoBarras & "'"
	QrySql = QrySql & " AND PH_Consumo_Detalle_Productos.Pendiente = " & idStatus 
	'
	'Response.Write QrySql
	'Response.End
	'	
	Set rsSql = Server.CreateObject("ADODB.recordset")
	rsSql.Open QrySql, conexion
	'
	if not (rsSql.EOF and rsSql.BOF) then
	   total  = rsSql(0)
	else
		total  = 0	
	end if
	'
	' Cerrar conexiones
	'
	Response.Write total
	rsSql.close
	set rsSql= Nothing	
	'
	conexion.close
	set conexion = nothing
	'
%>