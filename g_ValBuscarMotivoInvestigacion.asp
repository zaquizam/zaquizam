<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_ValBuscarMotivoInvestigacion.asp // 10ene21 - 14ene21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'
	Dim idConsumo
	'
	idConsumo = Request.QueryString("idConsumo")
	'
	Dim strSQL, rsscroll
	Set rsscroll = CreateObject("ADODB.Recordset")	
	'	
	strSQL = vbnullstring	
	strSQL = strSQL & " SELECT"
	strSQL = strSQL & " PH_InvestigacionItems.InvestigacionItems AS motivo"
	strSQL = strSQL & " FROM"
	strSQL = strSQL & " PH_Consumo"
	strSQL = strSQL & " INNER JOIN PH_Consumo_Investigar_Detalle ON PH_Consumo.Id_Consumo = PH_Consumo_Investigar_Detalle.Id_Consumo"
	strSQL = strSQL & " INNER JOIN PH_InvestigacionItems ON PH_Consumo_Investigar_Detalle.Id_items_investigacion = PH_InvestigacionItems.Id_InvestigacionItems"
	strSQL = strSQL & " WHERE"
	strSQL = strSQL & " PH_Consumo.Id_Consumo = " & idConsumo
	''
	rsscroll.open strSQL, conexion
	'
	' Response.write strSQL
	' Response.end
	'
	If not rsscroll.EOF  Then
		Response.write rsscroll("motivo")
	Else
		Response.write false
	End If
	'
	rsscroll.close : set rsscroll = nothing 
	conexion.Close : Set conexion = Nothing
	'
%>