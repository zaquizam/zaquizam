
<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ynum=Request.QueryString("num") 
	yval=Request.QueryString("pro")
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
	rsx3("Precio_producto") =  yval1 
	'response.write "<br>LLEGO25" & yval1
	'response.end
	rsx3("Tasa_de_cambio") = 1
	sx = yval * cdbl(rsx3("Cantidad"))
	sx = replace(sx,",",".")
	rsx3("Total_compra") =	sx  
	rsx3("Moneda") = "Bolivar Soberano"
	rsx3("id_Moneda") = 2

	rsx3.Update
	rsx3.Close 
	set rsx3 = nothing 
	
	%>

	<%
	
%>