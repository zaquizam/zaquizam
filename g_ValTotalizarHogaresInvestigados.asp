<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValTotalizarHogaresInvestigados.asp // 13ene21 - 14ene21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim rsPendientes, rsArray, idSemana, strSQL, intRow
	'	
	idSemana	=	Request.Querystring("id_Semana")
	'
	' Buscar total de hogares investigados
	'	
	Set rsPendientes = CreateObject("ADODB.Recordset")	
	'
	strSQL = vbnullstring	
	strSQL = " SELECT"
	strSQL = strSQL & " COUNT ( Enviado_investigar)"
	strSQL = strSQL & " FROM"
	strSQL = strSQL & " PH_Consumo"
	strSQL = strSQL & " WHERE"
	strSQL = strSQL & " Id_Semana = " & idSemana
	strSQL = strSQL & " AND"
	strSQL = strSQL & " Resuelto=0"
	strSQL = strSQL & " AND"
	strSQL = strSQL & " Enviado_investigar=1"
	''
	rsPendientes.open strSQL, conexion	
	'
	If not rsPendientes.EOF  Then
		'rsArray = rsPendientes.GetRows()
        'intRow = UBound(rsArray, 2) + 1 		
        response.write rsPendientes(0)
	Else
		Response.write 0
	End If
	'
	rsPendientes.close : set rsPendientes = nothing 
	conexion.Close : Set conexion = Nothing
	'
%>