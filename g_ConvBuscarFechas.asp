<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	'
	'g_ConvBuscarFechas.asp - 09ABR21 - 
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'
	Dim rsFechas, arrFechas
	'	
	' Buscar Los Hogares asociados al Estado
	'	
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " ss_Periodo.Semanas,"
	QrySql = QrySql & " ss_Periodo.Periodo"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " ss_Periodo"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " ss_Periodo.Semanas IS NOT NULL"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " ss_Periodo.idPeriodo ASC"
	'
	'Response.Write QrySql '& "<BR><BR>"
	'Response.end
	'
	Set rsFechas = Server.CreateObject("ADODB.recordset")
	rsFechas.Open QrySql,conexion
	'
	if not rsFechas.EOF then
    	arrFechas = rsFechas.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	'Response.ContentType = "application/json"
	'
	'Crear Archivo Array Json
	'
	sTabla=""

    if IsArray(arrFechas) then

        For i = 0 to ubound(arrFechas, 2)
            '
			' tiene comillas doble adicionales'
    		'sTabla     =  chr(123) &  chr(34) & "id" 	 & chr(34) & ":" & chr(34) &  cstr(Replace(arrFechas(0,i),",","-")) & chr(34) & chr(44)
			sTabla     =  chr(123) &  chr(34) & "id" 	 & chr(34) & ":" & chr(34) &  arrFechas(0,i) & chr(34) & chr(44)
			sTabla     =  sTabla   &  chr(34) & "name"   & chr(34) & ":" & chr(34) &  arrFechas(1,i) & chr(34) & chr(125) & chr(44)
            sTablaJson =  sTablaJson & sTabla
            sTabla=""
            '
        next

    else
        'Eof()
        sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
        sTabla    =   sTabla &   chr(34) & "name"   & chr(34)& ":" & chr(34) & "No hay Registros" 	& chr(34) & chr(125) & chr(44)
        '
        sTablaJson = sTablaJson & sTabla
        sTabla=""

    end if
	''
	sTabla 		= 	Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData	= 	chr(91) & sTabla & chr(93) '& chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'
	rsFechas.Close
	Set rsFechas = Nothing
	'
	conexion.close
	set conexion = nothing
	'
%>