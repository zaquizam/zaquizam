<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	CONST adOpenDynamic = 2
	CONST adUseClient = 3
	CONST adVarChar = 200
	CONST adDouble = 5
	CONST adDecimal  = 14

	dim RsTabla1
	Set RsTabla1 = Server.CreateObject("ADODB.Recordset")

	RsTabla1.CursorLocation = adUseClient
	RsTabla1.CursorType = adOpenDynamic
	RsTabla1.Fields.Append "Precio", adDecimal
	RsTabla1.Fields.Append "Id", adDouble
	RsTabla1.open
	
	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Precio_producto, "
	sql = sql & " Id_Consumo_Detalle_Productos "
	sql = sql & " FROM "
	sql = sql & " PH_Consumo_Detalle_Productos "
	sql = sql & " WHERE "
	sql = sql & " Numero_codigo_barras = '7591211001557'"
	sql = sql & " AND Moneda = 'BolivarSoberano'"
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br> Linea 223 " &
	'response.end
	if rsx1.eof then
		rsx1.close
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
	end if
	ix = ubound(gDatosSol,2) 
	for iReg = 0 to ubound(gDatosSol,2)
		Response.write "<br>40 Precio:= " & gDatosSol(0,iReg)
		Response.write "===>" & gDatosSol(1,iReg)
		RsTabla1.AddNew
		RsTabla1.Fields("Precio") = gDatosSol(0,iReg)
		RsTabla1.Fields("id") = gDatosSol(1,iReg)
		RsTabla1.update	
	next
	Response.write "<br><br>"
	RsTabla1.Sort = "Precio Desc" 
	RsTabla1.movefirst
	for iReg = 0 to ix
		Response.write "<br>61 Precio:= " & RsTabla1("Precio")
		Response.write "===>" & RsTabla1("id")
		RsTabla1.MoveNext
	next
	Response.write "<br><br>"
	
	RsTabla1.movefirst
	sql = ""
	sql = sql & " SELECT RsTabla1.Precio "
	sql = sql & " Count(RsTabla1.id) AS Total "
	sql = sql & " FROM "
	sql = sql & " RsTabla1.table "
	sql = sql & " GROUP BY "
	sql = sql & " RsTabla1.Precio "
	sql = sql & " ORDER BY 1 Desc "
	response.write "<br>77 sql = " & sql & "<br>"
	'response.end
	rsx1.Open sql ,conexion
	
	if rsx1.eof then
		rsx1.close
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
	end if
	
	Response.write "<br>40 Precio:= " & gDatosSol(0,0)
	

	
	
%>