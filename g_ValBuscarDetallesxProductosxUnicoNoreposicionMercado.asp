<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValBuscarDetallesxProductosxUnicoNoreposicionMercado.asp
	' 31ene21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idConsumo, rsComida, arrComida
	'
	idConsumo	=	Request.QueryString("id_Consumo")		
	'	
	' Buscar Resultados
	'
	set rsComida			=	CreateObject("ADODB.Recordset")
	rsComida.CursorType	=	adOpenKeyset 
	rsComida.LockType		=	2 'adLockOptimistic 
	'		
	sql = vbnullstring	
	sql = sql & " SELECT"
	sql = sql & " PH_Consumo.id_tipocomida,"
	sql = sql & " PH_Consumo.nombre_local,"
	sql = sql & " PH_Consumo.Total_compra,"	
	sql = sql & " PH_Consumo.id_Moneda,"
	sql = sql & " PH_Moneda.MONEDA"
	sql = sql & " FROM"
	sql = sql & " PH_Consumo"
	sql = sql & " INNER JOIN PH_MONEDA ON PH_Consumo.Id_Moneda = PH_Moneda.Id_MONEDA"
	sql = sql & " WHERE"
	sql = sql & " PH_Consumo.Id_Consumo = " & idConsumo
	'	
	'Response.Write sql
	'Response.End
	'
    rsComida.Open sql, conexion
	'
	if not rsComida.eof then
		arrComida = rsComida.GetRows()  ' Convert recordset to 2D Array					
	end if
		'
	rsComida.Close
	Set rsComida = Nothing
	'
	'Response.ContentType = "application/json"		
	'
	sTabla=vbnullstring
	
	if IsArray(arrComida) then
	
		For i = 0 to ubound(arrComida, 2)
			sTabla    =   	chr(123) &  chr(34) & "idtipocomida"	& chr(34) & ":" & chr(34) & arrComida(0,i)       & chr(34) & chr(44)
			sTabla    =    	sTabla   &  chr(34) & "nombrelocal" 	& chr(34) & ":" & chr(34) & arrComida(1,i)       & chr(34) & chr(44)
			sTabla    =    	sTabla   &  chr(34) & "idmoneda"	    & chr(34) & ":" & chr(34) & Cstr(arrComida(3,i)) & chr(34) & chr(44)
			sTabla    =    	sTabla   &  chr(34) & "moneda"	        & chr(34) & ":" & chr(34) & Cstr(arrComida(4,i)) & chr(34) & chr(44)
			total     =		replace(arrComida(2,i),",",".")
			sTabla    =    	sTabla   &  chr(34) & "totalcompra"    & chr(34) & ":" & chr(34) & total & chr(34) & chr(125) & chr(44)
			'
			sTablaJson = sTablaJson & sTabla
			sTabla=vbnullstring
		next
		'
		sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
		JsonData= chr(91) & sTabla & chr(93) '& chr(125)
		'
	else
		'Eof()
		sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	end if
	Response.Write(JsonData)
	conexion.close
	set conexion = nothing	
	
%>