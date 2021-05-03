<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValTotalizarHogaresResueltos.asp // 20ene21 - 
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim rsResueltos, rsArray, idSemana, strSQL, intRow
	'	
	idSemana	=	Request.Querystring("id_Semana")
	'
	' Buscar total de hogares investigados
	'	
	Set rsResueltos = CreateObject("ADODB.Recordset")	
	'
	strSQL = vbnullstring	
	strSQL = " SELECT"
	strSQL = strSQL & " COUNT ( Resuelto)"
	strSQL = strSQL & " FROM"
	strSQL = strSQL & " PH_Consumo"
	strSQL = strSQL & " WHERE"
	strSQL = strSQL & " Id_Semana = " & idSemana
	strSQL = strSQL & " AND"
	strSQL = strSQL & " Resuelto=1"
	strSQL = strSQL & " AND"
	strSQL = strSQL & " Enviado_investigar=0"
	''
	rsResueltos.open strSQL, conexion	
	'
	If not rsResueltos.EOF  Then
		'rsArray = rsResueltos.GetRows()
        'intRow = UBound(rsArray, 2) + 1 		
        response.write rsResueltos(0)
	Else
		Response.write 0
	End If
	'
	rsResueltos.close : set rsResueltos = nothing 
	conexion.Close : Set conexion = Nothing
	'
%>