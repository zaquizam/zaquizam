<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValEliminarTodoelConsumo.asp
	' 30dic20 // 03ene21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idConsumo
	'	
	idConsumo	=	Request.Querystring("idConsumo")	
	'
	' Eliminar todo el Consumo con sus detalles
	'	
	QrySql = vbnullstring
	QrySql = QrySql & " DELETE"	
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo.Id_Consumo = " & idConsumo
	'
	' Response.Write QrySql '& "<BR><BR>"
	' Response.end
	'
	Set objExec = conexion.Execute(QrySql)
    Set objExec = Nothing	
    '
	' Eliminar el Detalle Factura
	'
	QrySql = vbnullstring
	QrySql = QrySql & " DELETE"	
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo_Detalle_Factura"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo_Detalle_Factura.Id_Consumo = " & idConsumo
	'
	Set objExec = conexion.Execute(QrySql)
    Set objExec = Nothing	
	'
	'
	' Eliminar el Detalle Productos
	'
	QrySql = vbnullstring
	QrySql = QrySql & " DELETE"	
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo_Detalle_Productos"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo_Detalle_Productos.Id_Consumo = " & idConsumo
	'
	Set objExec = conexion.Execute(QrySql)
    Set objExec = Nothing	
	'	
    If Err.Number = 0 Then
        Response.write True
    Else
        Response.write (Err.Description)
    End If
    '
    conexion.Close
    Set objExec = Nothing
    Set conexion = Nothing
	'
%>