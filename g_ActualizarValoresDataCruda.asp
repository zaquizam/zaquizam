
<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ycat=Request.QueryString("cat")
	
	
	'response.write "<br>LLEGO25" 
	'response.write "<br>ymes ="  & ymes
	'response.write "<br>ycat ="  & ycat
	'response.end
	
	'EXEC ActualizarValoresDataCruda @Id_Categoria = 67
	
	Ejecutar = "EXEC [cacevedo_atenas].[ActualizarValoresDataCruda] @Id_Categoria = " & ycat
	'response.write "<br>Ejecutar ="  & Ejecutar
	'response.end
	Set objExec = conexion.Execute(Ejecutar)
	Set objExec = Nothing
	
	
%>