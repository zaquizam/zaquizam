
<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ysem=Request.QueryString("sem") 
	ycat=Request.QueryString("cat")
	
	'response.write "<br>LLEGO25" 
	'response.end
	
	'EXEC ProcesamientoSemanal @Id_Semana = 26, @Id_Categoria = 1
	
	Ejecutar = "EXEC [cacevedo_atenas].[ProcesamientoSemanal] @Id_Semana =" & ysem & ", @Id_Categoria = " & ycat
	'response.write "<br>Ejecutar ="  & Ejecutar
	'response.end
	'	
	Set objExec = conexion.Execute(Ejecutar)
	Set objExec = Nothing
	
	
%>