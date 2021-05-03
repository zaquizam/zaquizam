<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValUpdateDetallesxProductosMasivoDeshacer.asp // 20ene21 - 
	'
	Dim updSql	
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
    ' Capturar las variables
    '	
	idConsumo = Request.QueryString("idConsumo")
    '
    ' Actualizar Datos Validando....
    '
    updSql = vbnullstring
	updSql = updSql & " UPDATE PH_Consumo_Detalle_Productos "
    updSql = updSql & " SET"
    updSql = updSql & " Validado='0',"
	updSql = updSql & " Resuelto='0'"
	'updSql = updSql & " Pendiente='0',"
    updSql = updSql & " WHERE"
	updSql = updSql & " Pendiente='0'"
	updSql = updSql & " AND"	
    updSql = updSql & " Id_Consumo =" & idConsumo
    '
    ' Response.Write updSql
	' Response.end
    '
    Set objExec = conexion.Execute(updSql)
    Set objExec = Nothing
    '
    ' Actualizar Datos Validando....
    '
	' Set rsscroll = CreateObject("ADODB.Recordset")
    ' Dim strSQL, rsscroll, intRow
    ' strSQL = "SELECT COUNT(Cantidad) AS Total FROM PH_Consumo_Detalle_Productos WHERE Validado='0' AND Id_Consumo =" & idConsumo
    ' rsscroll.open strSQL, conexion
    ' intRow = rsscroll("Total").value
    ' rsscroll.close: set rsscroll = nothing 
	'	
	' if CInt(intRow) = 0 then			
		' 'Actualizar Maestro de Consumo
		' updSql = vbnullstring
		' updSql = updSql & " UPDATE PH_Consumo"
		' updSql = updSql & " SET"
		' updSql = updSql & " Validado='1'"		'
		' updSql = updSql & " WHERE"
		' updSql = updSql & " Id_Consumo =" & idConsumo
		' '		
		' Set objExec = conexion.Execute(updSql)
		' Set objExec = Nothing
		' '
	' else
		'Actualizar Maestro de Consumo
		
	If Err.Number = 0 Then
		updSql = vbnullstring
		updSql = updSql & " UPDATE PH_Consumo"
		updSql = updSql & " SET"
		updSql = updSql & " Validado='0'"		'
		updSql = updSql & " WHERE"
		updSql = updSql & " Id_Consumo =" & idConsumo
		'		
		Set objExec = conexion.Execute(updSql)
		Set objExec = Nothing
		
		If Err.Number = 0 Then		
			Response.write True
		Else
			Response.write (Err.Description)
		End If  
	Else
			Response.write (Err.Description)
	End If   		
	'
	' end if	
	'
	Response.Write intRow
	'
	conexion.Close
    Set objExec = Nothing	
    Set conexion = Nothing
	'
%>