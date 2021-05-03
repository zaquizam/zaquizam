<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_ValUpdateDetallesxProductosxUnicoDirecto.asp // 04ene21 - 19feb21
	'
	Dim updSql	
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
    ' Capturar las variables
    '
	idConsumoDetalle = Request.QueryString("idConsumoDetalle")
	idConsumo        = Request.QueryString("idConsumo")
    '
    ' Actualizar Consumo Detalle Productos Validar....
    '
    updSql = vbnullstring
	updSql = updSql & " UPDATE PH_Consumo_Detalle_Productos "
    updSql = updSql & " SET"
    updSql = updSql & " Validado='1',"
	updSql = updSql & " Resuelto='0'"	
    updSql = updSql & " WHERE"
    updSql = updSql & " Id_Consumo_detalle_Productos =" & idConsumoDetalle
    '        
    Set objExec = conexion.Execute(updSql)
    Set objExec = Nothing
    '
    If Err.Number = 0 Then
		'Response.Write updSql
		'Response.end
		'
		' Actualizar Datos Validando....
		'
		Set rsscroll = CreateObject("ADODB.Recordset")
		Dim strSQL, rsscroll, intRow
		'strSQL = "SELECT COUNT(Cantidad) AS Total FROM PH_Consumo_Detalle_Productos WHERE Validado='0' AND pend AND Id_Consumo =" & CInt(idConsumo)
		strSQL = "SELECT COUNT(Cantidad) AS Total FROM PH_Consumo_Detalle_Productos WHERE Validado='0' AND Pendiente=0 AND Id_Consumo =" & idConsumo
		rsscroll.open strSQL, conexion
		intRow = rsscroll("Total").value
		rsscroll.close: set rsscroll = nothing 
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
			Set objExec = conexion.Execute(updSql)
			Set objExec = Nothing
			'
			'Response.Write updSql
			'Response.end
			''
		end if	
		'
		If Err.Number = 0 Then
			Response.Write intRow
		Else
			Response.write (Err.Description)
		End If
		'
    Else
         Response.write (Err.Description)
    End If
	''
	conexion.Close
    Set objExec = Nothing	
    Set conexion = Nothing
	'
%>