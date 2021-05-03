<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_pPendUpdatePrecioMasivoProductosPendientes - 25mar21 - 28abr21		
	'
	Dim updSql, idConsumoDetalle, Promedio, tasacambio, idMoneda, precio
	'
	Session.lcid = 2057
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
    ' Capturar las variables
    '	
	precio 			= Request.Form("precio")		
	codigobarras	= Request.Form("barcode")	
	nCheckboxes		= split(Request.Form("checkboxes"),",")
    '
    ' Actualizar el precio ....
    '
	'	
	For i=LBound(nCheckboxes) to UBound(nCheckboxes)
		'
		'Response.Write nCheckboxes(i) + "<br>"
		'							
		updSql = vbnullstring
		updSql = updSql & " UPDATE PH_Consumo_Detalle_Productos"
		updSql = updSql & " SET"
		updSql = updSql & " precio_producto = ("  & precio 	& " / tasa_de_cambio) ,"
		updSql = updSql & " Total_compra = (" & precio & " * cantidad) ,"			
		updSql = updSql & " Validado = '0',"
		updSql = updSql & " Resuelto = '0',"
		updSql = updSql & " Pendiente ='1'"
		updSql = updSql & " WHERE"				
		updSql = updSql & " Id_Consumo_detalle_productos=" & CLng(nCheckboxes(i))
		'
		'Response.Write updSql
		'Response.End
		'
		Set objExec = conexion.Execute(updSql)
		'	
	Next
    '
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