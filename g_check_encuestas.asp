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
	set rsEncuesta = CreateObject("ADODB.Recordset")
	rsEncuesta.CursorType = 1 'adOpenKeyset 
	rsEncuesta.LockType =   3 '2 'adLockOptimistic 	
	'
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_EncuestaHogar.Id_EncuestaEspecial,"
	sql = sql & " PH_EncuestaHogar.Id_Hogar,"
	sql = sql & " PH_EncuestaHogar.Ind_Realizada,"
	sql = sql & " PH_EncuestaHogar.Ind_Rechazada"	
	sql = sql & " FROM"
	sql = sql & " PH_EncuestaHogar"
	sql = sql & " WHERE"
	sql = sql & " PH_EncuestaHogar.Id_Hogar =" & idhogar	
	sql = sql & " AND PH_EncuestaHogar.Id_EncuestaEspecial =" & idencuesta 
	'	
    rsEncuesta.Open sql ,conexion
	'
	if rsEncuesta.eof then
	
		rsEncuesta.close
		set rsEncuesta=nothing
		'
		instSql = vbnullstring
		instSql = instSql & " INSERT INTO PH_EncuestaHogar "
		instSql = instSql & " ("
		instSql = instSql & " Id_EncuestaEspecial, id_Hogar,"	
		instSql = instSql & " IP, idsession"
		instSql = instSql & " )"
		'
		instSql = instSql & " VALUES "
		'
		instSql = instSql & "(" & idencuesta & ","
		instSql = instSql & ""  & idHogar & ","    
		instSql = instSql & "'" & sIp & "',"
		'
		instSql = instSql & "" & Session.SessionID & ")"
		'
		Response.Write True
		'Response.End
		'
		Set objExec = conexion.Execute(instSql)
		'		
	else
		Response.Write false
	end if
	'
%>
	
	
	
