<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValUpdatePendientesMasivo.asp // 23feb21 - 
	'
	Dim updSql	
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
    ' Capturar las variables
    '	
	' idConsumoDetalle  	= Request.formQueryString("idConsumoDetalle")	
	' idConsumo  			= Request.QueryString("idConsumo")	
    '
	nCheckboxes	= split(Request.form("checkboxes"),",")
	idConsumo  	= Request.form("idConsumo")	
	'
	For i=LBound(nCheckboxes) to UBound(nCheckboxes)
		'
		'Response.Write nCheckboxes(i) + "<br>"
		'							
		updSql = vbnullstring
		updSql = updSql & " UPDATE PH_Consumo_Detalle_Productos"
		updSql = updSql & " SET"
		updSql = updSql & " Validado='0',"
		updSql = updSql & " Resuelto='0',"
		updSql = updSql & " Pendiente ='1'"
		updSql = updSql & " WHERE"		
		updSql = updSql & " Id_Consumo_detalle_productos=" & CLng(nCheckboxes(i))
		'
		Set objExec = conexion.Execute(updSql)
		'	
	Next
	'
	Set objExec = Nothing	
    '
    ' Actualizar Datos Validando....
    '
	Set rsscroll = CreateObject("ADODB.Recordset")
    Dim strSQL, rsscroll, intRow
    strSQL = "SELECT COUNT(Cantidad) AS Total FROM PH_Consumo_Detalle_Productos WHERE Validado='0' AND Pendiente=0 AND Id_Consumo =" & idConsumo
    rsscroll.open strSQL, conexion
    intRow = rsscroll("Total").value
    rsscroll.close: set rsscroll = nothing 
	'
	'Response.Write strSQL
	'Response.end
	'
	if CInt(intRow) = 0 then			
		'Actualizar Maestro de Consumo
		updSql = vbnullstring
		updSql = updSql & " UPDATE PH_Consumo"
		updSql = updSql & " SET"
		updSql = updSql & " Validado='1',"		'
		updSql = updSql & " Resuelto='0'"
		updSql = updSql & " WHERE"
		updSql = updSql & " Id_Consumo =" & idConsumo
		'		
		'Response.Write updSql
		'Response.end
		'		 
		Set objExec = conexion.Execute(updSql)
		Set objExec = Nothing
		'	
	end if	
	'
	Response.Write intRow
	'
	conexion.Close
    Set objExec = Nothing	
    Set conexion = Nothing
	'
%>