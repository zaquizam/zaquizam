<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_pPendUpdatePendientesxPromedio.asp // 09mar21 - 
	'
	Dim updSql	
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
    ' Capturar las variables
    '	
	idConsumoDetalle	= Request.formQueryString("idConDetalle")	
	Promedio		  	= Request.formQueryString("Promedio")	
	idMoneda   		  	= Request.formQueryString("idMoneda")	
    '
    ' Actualizar Datos Validando....
    '
	' Set rsscroll = CreateObject("ADODB.Recordset")
    ' Dim strSQL, rsscroll, intRow
    ' strSQL = "SELECT COUNT(Cantidad) AS Total FROM PH_Consumo_Detalle_Productos WHERE Validado='0' AND Pendiente=0 AND Id_Consumo =" & idConsumo
    ' rsscroll.open strSQL, conexion
    ' intRow = rsscroll("Total").value
    ' rsscroll.close: set rsscroll = nothing 
	' '
	' 'Response.Write strSQL
	' 'Response.end
	' '
	' if CInt(intRow) = 0 then			
		' 'Actualizar Maestro de Consumo
		' updSql = vbnullstring
		' updSql = updSql & " UPDATE PH_Consumo"
		' updSql = updSql & " SET"
		' updSql = updSql & " Validado='1',"		'
		' updSql = updSql & " Resuelto='0'"
		' updSql = updSql & " WHERE"
		' updSql = updSql & " Id_Consumo =" & idConsumo
		' '		
		' 'Response.Write updSql
		' 'Response.end
		' '		 
		' Set objExec = conexion.Execute(updSql)
		' Set objExec = Nothing
		' '	
	' end if	
	' '
	' Response.Write intRow
	'
	updSql = vbnullstring
	updSql = updSql & " UPDATE PH_Consumo_Detalle_Productos"
    updSql = updSql & " SET"
    'updSql = updSql & " Total_Compra=" & tmonto & ","
	'updSql = updSql & " Total_Items="  & tproducto & ","
    'updSql = updSql & " Id_Canal="     & canal & ","
    'updSql = updSql & " Id_Cadena= "   & cadena & ","    	
	updSql = updSql & " Precio_producto= "   & Promedio 
    'updSql = updSql & " Validado='0'"
	'
    updSql = updSql & " WHERE"
    updSql = updSql & " Id_Consumod_Detalle_productos =" & dConsumoDetalle
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