
<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ymes=Request.QueryString("mes") 
	ycat=Request.QueryString("cat")
	
	
	'response.write "<br>LLEGO25" 
	'response.write "<br>ymes ="  & ymes
	'response.write "<br>ycat ="  & ycat
	'response.end
	
	'EXEC ProcesamientoSemanal @Id_Semana = 26, @Id_Categoria = 1
	
	ymes = replace(ymes,",","*")
	
	Ejecutar = "EXEC [cacevedo_atenas].[ProcesamientoMensual] @Id_Semana =N'" & ymes & "', @Id_Categoria = " & ycat
	response.write "<br>Ejecutar ="  & Ejecutar
	response.end
	Set objExec = conexion.Execute(Ejecutar)
	Set objExec = Nothing
	
	
%>