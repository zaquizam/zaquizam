<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
' g_ValinvestigarConsumoUnico.asp - 10ene21 - 15abr21
'
Dim updSql, instSql
'
Session.lcid = 1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'
' Capturar las variables
'
idConsumo 	= Request.QueryString("id_Consumo")
idItemsInv 	= Request.QueryString("id_ItemsInv")
idHogar		= Request.QueryString("id_Hogar")	
observacion	= Request.QueryString("observacion")	
''
if Len(observacion)=0 or isNull(observacion) then observacion="** Sin Comentarios **"
'
'observacion = RemoverSaltodeLinea(observacion)
'
' Actualizar Datos Validando....
'
updSql = vbnullstring
updSql = updSql & " UPDATE PH_Consumo "
updSql = updSql & " SET"
updSql = updSql & " enviado_investigar=1"
updSql = updSql & " WHERE"
updSql = updSql & " Id_Consumo =" & idConsumo
'
' Response.Write updSql
' Response.end
'
Set objExec = conexion.Execute(updSql)
Set objExec = Nothing
'
 If Err.Number = 0 Then
	'
	' Insertar Datos ....
	'
	sIP = Request.ServerVariables("REMOTE_ADDR")
	idPendiente=1
	idValidado=0
	idResuelto=0
	idCasoCerrado=0
	'
	instSql = vbnullstring
	instSql = instSql & " INSERT INTO PH_Consumo_Investigar_Detalle "
	instSql = instSql & " ( Id_items_investigacion, id_Consumo, id_Hogar, Pendiente, Validado, Resuelto, Caso_Cerrado, Observaciones_enviadas, IP, idsession )"
	'
	instSql = instSql & " VALUES "
	'
	instSql = instSql & "(" & idItemsInv & ","
	instSql = instSql & ""  & idConsumo & ","
	instSql = instSql & ""  & idHogar & ","
	instSql = instSql & ""  & idPendiente & ","		
	instSql = instSql & ""  & idValidado & ","
	instSql = instSql & ""  & idResuelto & ","
	instSql = instSql & ""  & idCasoCerrado & ","		
	instSql = instSql & "'" & observacion & "',"
	instSql = instSql & "'" & sIp & "',"
	instSql = instSql & "'" & Session.SessionID & "')"
	'
	'Response.Write  instSql
	'Response.Write  End
	'
	Set objExec = conexion.Execute(instSql)
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
	
FUNCTION RemoverSaltodeLinea(byval str)

	IF isNull(str) THEN str = "" END IF
	str = REPLACE(str,vbCr," ")			'Chr(13)
	str = REPLACE(str,vbLf," ")			'Chr(10)
	str = REPLACE(str,VbCrlf," ")		'Chr(13)+Chr(10)
	str = REPLACE(str,vbNewLine," ")		'vbNewLine
	str = REPLACE(str,vbFormFeed," ")	'Chr(12)
	str = REPLACE(str,vbTab," ")			'Chr(9)
	str = REPLACE(str,vbTab," ")			'Chr(11)
	str = REPLACE(str,"'","`")			'Comillas simples
	str = REPLACE(str,"""", "`") 		'Comillas dobles		
	str = REPLACE(str,",", " ") 		'Comillas dobles
	'
	RemoverSaltodeLinea = TRIM(str)
	'
END FUNCTION
	
%>
