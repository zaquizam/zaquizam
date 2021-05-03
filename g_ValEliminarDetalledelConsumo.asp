<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValEliminarDetalledelConsumo.asp
	'
	' 05ene21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idConsumo
	'	
	idConsumo	=	Request.Querystring("idConsumo")		
	'
	'
	' Eliminar el Detalle Productos
	'
	QrySql = vbnullstring
	QrySql = QrySql & " DELETE"	
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo_Detalle_Productos"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo_Detalle_Productos.Id_Consumo_detalle_productos = " & idConsumo
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