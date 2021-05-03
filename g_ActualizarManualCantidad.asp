
<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ynum=Request.QueryString("num") 
	ynum=ynum * -1
	yval=Request.QueryString("man")
	yval1=replace(yval,",",".")
	
	'response.write "<br>LLEGO25" & yval
	'response.end

	dim rsx3
	set rsx3 = CreateObject("ADODB.Recordset")
	rsx3.CursorType = 0
	rsx3.LockType = 3 

	sql = ""
	sql = sql & " Select * from PH_Consumo_Detalle_Productos "
	sql = sql & " Where Id_Consumo_Detalle_Productos = " & ynum
	'response.write "<br>220 sql:=" & sql
	'response.end 
	rsx3.Open sql ,conexion
	rsx3("Cantidad") =  yval1 
	'response.write "<br>LLEGO25" & yval1
	'response.end
	rsx3.Update
	rsx3.Close 
	set rsx3 = nothing 
	
	%>

	<%
	
%>