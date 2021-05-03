<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_rRevInvRespuestaInvestigacionConsumo.asp // 14ene21 - 22ene21
	'	
	Dim dd, mm, yy, hh, nn, ss, updSql
	Dim datevalue, timevalue, dtsnow
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
    ' Capturar las variables
    '
	idConsumo 	= Request.QueryString("id_Consumo")	
	observacion	= Request.QueryString("observacion")
	'
	'
	dtsnow = Now()
	dd = Right("00" & Day(dtsnow), 2)
	mm = Right("00" & Month(dtsnow), 2)
	yy = Year(dtsnow)
	hh = Right("00" & Hour(dtsnow), 2)
	nn = Right("00"  & Minute(dtsnow), 2)
	ss = Right("00" & Second(dtsnow), 2)
	datevalue = yy  & "-" & mm & "-" & dd
	timevalue = hh  & ":" & nn & ":" & ss
	sUpdate = datevalue & " " & timevalue
	'
    ' Actualizar Datos Validando....
    '
    updSql = vbnullstring
	updSql = updSql & " UPDATE PH_Consumo"
    updSql = updSql & " SET"
    updSql = updSql & " Enviado_investigar=0,"
	updSql = updSql & " Resuelto=1,"
	updSql = updSql & " Fec_Inactivo= '" & sUpdate & "'"
    updSql = updSql & " WHERE"
    updSql = updSql & " Id_Consumo =" & idConsumo
    '        
    Set objExec = conexion.Execute(updSql)
	Set objExec = Nothing
	'
	 If Err.Number = 0 Then
		''
		' Response.Write updSql
		' Response.end
        '
		' Actualizar Datos ....
		'				
		updSql = vbnullstring
		updSql = updSql & " UPDATE PH_Consumo_Investigar_Detalle"
		updSql = updSql & " SET"
		updSql = updSql & " PH_Consumo_Investigar_Detalle.pendiente=0,"
		updSql = updSql & " PH_Consumo_Investigar_Detalle.resuelto=1,"
		updSql = updSql & " PH_Consumo_Investigar_Detalle.caso_cerrado=1,"
		updSql = updSql & " PH_Consumo_Investigar_Detalle.Fec_Inactivo= '" & sUpdate & "',"
		updSql = updSql & " PH_Consumo_Investigar_Detalle.Observaciones_recibidas= '" & observacion & "'"
		'
		updSql = updSql & " WHERE"
		updSql = updSql & " PH_Consumo_Investigar_Detalle.Id_Consumo =" & idConsumo
		'
		' Response.Write  updSql
		' Response.End
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
    Else
        Response.write (Err.Description)
    End If    
	'
%>