<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValTotalizarConsumosxValidadosyPendientes.asp // 08ene21 - 18Feb21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idhogar, rsTotalConsumo, arrTotalConsumo
	'	
	idSemana	=	Request.Querystring("id_Semana")
	'
	' Buscar los detalles del Consumo
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " validado,"
	QrySql = QrySql & " COUNT(validado) total"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " Id_Semana = " & idSemana
	QrySql = QrySql & " AND"
	QrySql = QrySql & " Id_Hogar > 1"	
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " validado"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " validado DESC;"
	'
	' Response.Write QrySql '& "<BR><BR>"
	' Response.end
	'
	Set rsTotalConsumo = Server.CreateObject("ADODB.recordset")
	rsTotalConsumo.Open QrySql, conexion
	'
	if not rsTotalConsumo.EOF then
    	arrTotalConsumo = rsTotalConsumo.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"		
	'
	sTabla=vbnullstring
	
	if IsArray(arrTotalConsumo) then
	
		For i = 0 to ubound(arrTotalConsumo, 2)
		
			sTabla    =   chr(123) &  chr(34) & "validado"  & chr(34) & ":" & chr(34) & arrTotalConsumo(0,i) & chr(34) & chr(44)
			'
			sTabla    =    sTabla  &  chr(34) & "total"     & chr(34) & ":" & chr(34) & arrTotalConsumo(1,i) & chr(34) & chr(125) & chr(44)
			
			sTablaJson = sTablaJson & sTabla
			sTabla=vbnullstring
			
		next				
		'
	else
		'Eof()
		'sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		'
		sTabla    =   chr(123) &  chr(34) & "validado"  & chr(34) & ":" & chr(34) & "False" & chr(34) & chr(44)
		'
		sTabla    =    sTabla  &  chr(34) & "total"     & chr(34) & ":" & chr(34) & "0" & chr(34) & chr(125) & chr(44)
		
		sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		
		'JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	end if
	'	
	sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData= chr(91) & sTabla & chr(93) '& chr(125)
	Response.Write(JsonData)
	'
	conexion.close
	set conexion = nothing		
%>