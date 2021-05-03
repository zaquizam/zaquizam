
<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	'response.write "<br>LLEGO25" 
	'response.end
	ynum=Request.QueryString("num") 
	
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
	sql = sql & " delete "
	sql = sql & " from ppi_evapractica_entrenador "
	sql = sql & " where "
	sql = sql & " fk_postulacion  = '" & ynum & "'"
	rsx1.Open sql ,conexion

	sql = ""
	sql = sql & " delete "
	sql = sql & " from ppi_entrevistabr "
	sql = sql & " where "
	sql = sql & " fk_postulacion = '" & ynum & "'"
	rsx1.Open sql ,conexion

	sql = ""
	sql = sql & " delete "
	sql = sql & " from ppi_calificacionbr "
	sql = sql & " where "
	sql = sql & " fk_postulacion = '" & ynum & "'"
	rsx1.Open sql ,conexion

	sql = ""
	sql = sql & " delete "
	sql = sql & " from ppi_evateoricobr "
	sql = sql & " where "
	sql = sql & " fk_postulacion = '" & ynum & "'"
	rsx1.Open sql ,conexion

	sql = ""
	sql = sql & " delete "
	sql = sql & " from ppi_evateoricobr "
	sql = sql & " where "
	sql = sql & " fk_postulacion = '" & ynum & "'"
	rsx1.Open sql ,conexion

	sql = ""
	sql = sql & " delete "
	sql = sql & " from ppi_postulantesbr "
	sql = sql & " where id = '" & ynum & "'"
	rsx1.Open sql ,conexion

	%>
	Eliminado
	<%
	
%>