<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_pPendUpdateProductosPendientesxTodos.asp - 09mar21 - 22mar21
	'
	Dim updSql, idConsumoDetalle, Promedio, tasacambio, idMoneda, precio
	'
	Session.lcid = 2057
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
    ' Capturar las variables
    '	
	idConsumoDetalle	= Request.QueryString("idConDetalle")	
	Promedio		  	= Request.QueryString("Promedio")	
	idMoneda   		  	= Request.QueryString("idMoneda")	
	tasacambio			= Request.QueryString("tasaCambio")			
	cantidad 			= Request.QueryString("cantidad")
	'
	' response.write idConsumoDetalle & "<br>"
	' response.write promedio & "<br>"
	' response.write idmoneda & "<br>"
	' response.write tasacambio & "<br>"
	' response.end
	'	
	precio = 0
	totalcompra = 0
	'
	if CInt(idMoneda) <> 2 then
		precio = Promedio / tasacambio
		'precio = round(precio,2)
		'TotalCompra = ( Precio * tasaCambio ) * Cantidad
		TotalCompra = (Precio * Cantidad)
	else
		'bolivar
		precio = Promedio
		TotalCompra =  (Precio * tasaCambio ) * Cantidad
	end if
    '
    ' Actualizar Datos Validando....
    '	
	updSql = vbnullstring
	updSql = updSql & " UPDATE PH_Consumo_Detalle_Productos"
    updSql = updSql & " SET"    
	updSql = updSql & " Precio_producto= "  & precio & ","
	updSql = updSql & " total_compra   = "  & totalcompra
	updSql = updSql & " WHERE"
    updSql = updSql & " Id_Consumo_Detalle_productos = " & idConsumoDetalle
    '	
    ' Response.Write updSql
	' Response.end
    '
    Set objExec = conexion.Execute(updSql)
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
	
	'
%>