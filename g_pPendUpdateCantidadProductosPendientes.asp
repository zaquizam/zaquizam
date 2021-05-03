<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_pPendUpdateCantidadProductosPendientes.asp - 17mar21 - 22mar21		
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
	cantidad 			= Request.QueryString("cantidad")	
    '
    ' Actualizar Cantidad Validando....
    '	
	updSql = vbnullstring
	updSql = updSql & " UPDATE PH_Consumo_Detalle_Productos"
    updSql = updSql & " SET"    
	updSql = updSql & " cantidad = "  & cantidad & ","
	updSql = updSql & " Total_compra = (precio_producto * tasa_de_cambio) * "  & cantidad	
	updSql = updSql & " WHERE"
    updSql = updSql & " Id_Consumo_Detalle_productos = " & idConsumoDetalle
    '
    'Response.Write updSql
	'Response.end
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