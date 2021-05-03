<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	'g_ValTipoInvestigacion.asp
	'10ene21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'
	Dim rsTipoInvestigacion, arrTipoInvestigacion
	'	
	' Buscar Los Hogares asociados al Estado
	'	
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_InvestigacionItems.Id_InvestigacionItems As Id,"
	QrySql = QrySql & " PH_InvestigacionItems.InvestigacionItems as Nombre"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_InvestigacionItems"	
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_InvestigacionItems.InvestigacionItems"	
	'	
	'
	'Response.Write QrySql '& "<BR><BR>"
	'Response.end
	'
	Set rsTipoInvestigacion = Server.CreateObject("ADODB.recordset")
	rsTipoInvestigacion.Open QrySql,conexion
	'
	if not rsTipoInvestigacion.EOF then
    	arrTipoInvestigacion = rsTipoInvestigacion.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	'
	'Crear Archivo Array Json
	'
	sTabla=""

    if IsArray(arrTipoInvestigacion) then

        For i = 0 to ubound(arrTipoInvestigacion, 2)
            '
			sTabla     =  chr(123) &  chr(34) & "id" 	  & chr(34) & ":" & arrTipoInvestigacion(0,i) & chr(44)
            sTabla     =  sTabla   &  chr(34) & "nombre"  & chr(34) & ":" & chr(34) & arrTipoInvestigacion(1,i)  & chr(34) & chr(125) & chr(44)
            sTablaJson =  sTablaJson & sTabla
            sTabla=""
            '
        next

    else
        'Eof()
        sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
        sTabla    =   sTabla &   chr(34) & "nombre"   & chr(34)& ":" & chr(34) & "No hay Registros" 	& chr(34) & chr(125) & chr(44)
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
	rsTipoInvestigacion.Close
	Set rsTipoInvestigacion = Nothing
	'
	conexion.close
	set conexion = nothing
	'
%>