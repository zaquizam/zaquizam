<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValTotalizarHogaresxConsumos.asp // 08ene21 - 19ene21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim rsscroll, rsArray, idSemana, strSQL, intRow
	'	
	idSemana	=	Request.Querystring("id_Semana")
	'
	' Buscar los detalles del Consumo
	'	
	'	
	Set rsscroll = CreateObject("ADODB.Recordset")	
	'
	strSQL = vbnullstring	
	strSQL = " SELECT"	
	strSQL = strSQL & " PH_Consumo.Id_Hogar"
	strSQL = strSQL & " FROM"
	strSQL = strSQL & " PH_Consumo"
	strSQL = strSQL & " INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar"
	strSQL = strSQL & " WHERE"
	strSQL = strSQL & " PH_Consumo.Id_Semana = " & idSemana
	strSQL = strSQL & " AND"
	strSQL = strSQL & " PH_PanelHogar.Ind_Activo = 1"
	strSQL = strSQL & " GROUP BY"
	strSQL = strSQL & " PH_Consumo.Id_Hogar"
	strSQL = strSQL & " HAVING"
	strSQL = strSQL & " PH_Consumo.Id_Hogar > 1"	
	'
	rsscroll.open strSQL, conexion	
	'
	If not rsscroll.EOF  Then
		rsArray = rsscroll.GetRows() 
        intRow = UBound(rsArray, 2) + 1 
        response.write intRow
	Else
		Response.write 0
	End If
	'
	rsscroll.close : set rsscroll = nothing 
	conexion.Close : Set conexion = Nothing
	'
%>