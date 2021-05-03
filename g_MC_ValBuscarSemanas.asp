<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	'
	'g_MC_ValBuscarSemanas.asp - 26feb21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'
	Dim rsSemanas, arrSemanas
	'	
	' Buscar Los Hogares asociados al Estado
	'	
	QrySql = vbnullstring
	'QrySql = QrySql & " SELECT TOP 6"
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " ss_Semana.Id,"
	QrySql = QrySql & " ss_Semana.Semana"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " ss_Semana"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " ss_Semana.IdSemana > 14"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " ss_Semana.Id DESC"	
	'	
	'
	'Response.Write QrySql '& "<BR><BR>"
	'Response.end
	'
	Set rsSemanas = Server.CreateObject("ADODB.recordset")
	rsSemanas.Open QrySql,conexion
	'
	if not rsSemanas.EOF then
    	arrSemanas = rsSemanas.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	'
	'Crear Archivo Array Json
	'
	sTabla=""

    if IsArray(arrSemanas) then

        For i = 0 to ubound(arrSemanas, 2)
            '
            'sTabla     =  chr(123) &  chr(34) & "Id" 	 & chr(34) & ":" & chr(34) & arrSemanas(0,i)  & chr(34) & chr(44)
			sTabla     =  chr(123) &  chr(34) & "Id" 	 & chr(34) & ":" & arrSemanas(0,i) & chr(44)
            sTabla     =  sTabla   &  chr(34) & "Name"   & chr(34) & ":" & chr(34) & arrSemanas(1,i)  & chr(34) & chr(125) & chr(44)
            sTablaJson =  sTablaJson & sTabla
            sTabla=""
            '
        next

    else
        'Eof()
        sTabla    =   chr(123)&  chr(34) & "Id" 	& chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
        sTabla    =   sTabla &   chr(34) & "Name"   & chr(34)& ":" & chr(34) & "No hay Registros" 	& chr(34) & chr(125) & chr(44)
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
	rsSemanas.Close
	Set rsSemanas = Nothing
	'
	conexion.close
	set conexion = nothing
	'
%>