<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idhogar, idencuesta, rsEncuesta
	
	idhogar   = Request.QueryString("idHogar")
	idencuesta= Request.QueryString("idEncuesta")
	'			
	'Borrar Encuessta
	'
	sql = ""
	sql = sql & " DELETE "
	sql = sql & " FROM"
	sql = sql & " PH_EncuestaHogar"
	sql = sql & " WHERE"
	sql = sql & " PH_EncuestaHogar.Id_Hogar =" & idhogar	
	sql = sql & " AND PH_EncuestaHogar.Id_EncuestaEspecial =" & idencuesta 
	'	
	Set objExec = conexion.Execute(sql)
	Set objExec = Nothing
	'
	'Borrar Resultados
	'
	sql = ""
	sql = sql & " DELETE "
	sql = sql & " FROM"
	sql = sql & " PH_EncuestaEspecialResultados"
	sql = sql & " WHERE"
	sql = sql & " PH_EncuestaEspecialResultados.Id_Hogar =" & idhogar	
	sql = sql & " AND PH_EncuestaEspecialResultados.Id_EncuestaEspecial =" & idencuesta 
	'	
	Set objExec = conexion.Execute(sql)
	Set objExec = Nothing
	'		    
	If Err.Number = 0 Then
        Response.write True 
    Else
        Response.write False
    End If		    
	
	'
%>
	
	
	
