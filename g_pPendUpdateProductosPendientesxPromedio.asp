<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_pPendUpdateProductosPendientesxPromedio.asp // 09mar21 - 
	'
	Dim updSql, idConsumoDetalle, Promedio, tasacambio, idMoneda, precio
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
    ' Capturar las variables
    '	
	idConsumoDetalle	= Request.QueryString("idConDetalle")	
	Promedio		  	= Request.QueryString("Promedio")	
	idMoneda   		  	= Request.QueryString("idMoneda")	
	tasacambio			= Request.QueryString("tasaCambio")			
	'
	response.write idConsumoDetalle
	response.write promedio
	response.write idmoneda
	response.write tasacambio
	response.end
	
	
	precio=0
	if CInt(idMoneda) <> 2 then
		precio = cdbl(Promedio) / cdbl(tasacambio)
	else
		precio = cdbl(Promedio)
	end if
    '
    ' Actualizar Datos Validando....
    '	
	updSql = vbnullstring
	updSql = updSql & " UPDATE PH_Consumo_Detalle_Productos"
    updSql = updSql & " SET"    
	updSql = updSql & " Precio_producto= "  & precio
    updSql = updSql & " WHERE"
    updSql = updSql & " Id_Consumo_Detalle_productos = " & idConsumoDetalle
    '
    Response.Write updSql
	Response.end
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