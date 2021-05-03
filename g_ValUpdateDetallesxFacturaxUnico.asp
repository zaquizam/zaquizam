<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	'
	' g_ValUpdateDetallesxFacturaxUnico.asp //	03ene21 - 21ene21
	'
	Dim updSql	
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
    ' Capturar las variables
    '			
	idConsumo 	= Request.QueryString("idConsumo")
	canal		= Request.QueryString("canal")
	cadena		= Request.QueryString("cadena")
	tmonto		= Request.QueryString("totalFact")
	tproducto	= Request.QueryString("totalProd")
	idmoneda	= Request.QueryString("idMoneda")
    '
    ' Actualizar Datos Validando....
    '
    updSql = vbnullstring
	updSql = updSql & " UPDATE PH_Consumo"
    updSql = updSql & " SET"
    updSql = updSql & " Total_Compra=" & tmonto & ","
	updSql = updSql & " Total_Items="  & tproducto & ","
    updSql = updSql & " Id_Canal="     & canal & ","
    updSql = updSql & " Id_Cadena= "   & cadena & ","    	
	updSql = updSql & " Id_Moneda= "   & idMoneda 
    'updSql = updSql & " Validado='0'"
	'
    updSql = updSql & " WHERE"
    updSql = updSql & " Id_Consumo =" & idConsumo
    '
    'Response.Write updSql
	'Response.end
    '
    Set objExec = conexion.Execute(updSql)
    Set objExec = Nothing
    '
    If Err.Number = 0 Then
        Response.write True
    Else
        Response.write (Err.Description)
    End If
    '
    conexion.Close
    Set objExec = Nothing
    Set conexion = Nothing
    '
%>