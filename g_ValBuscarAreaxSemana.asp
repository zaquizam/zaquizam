<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	'g_ValBuscarAreaxSemana.asp
	'29dic20
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'
	Dim rsEstado, idArea, arrEstado
	'
	idSemana = Request.QueryString("id_Semana")
	'
	' Buscar Los Estados asociados al Area
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_GArea.Id_Area,"
	QrySql = QrySql & " PH_GArea.Area"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo"
	QrySql = QrySql & " INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar"
	QrySql = QrySql & " INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado"
	QrySql = QrySql & " INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo.Id_Semana =" & idSemana
	QrySql = QrySql & " AND PH_Consumo.Status_registro ='G'"
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_GArea.Id_Area,"
	QrySql = QrySql & " PH_GArea.Area"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_GArea.Area ASC"
	'
	'Response.Write QrySql '& "<BR><BR>"
	'Response.end
	'
	Set rsEstado = Server.CreateObject("ADODB.recordset")
	rsEstado.Open QrySql,conexion
	'
	if not rsEstado.EOF then
    	arrEstado = rsEstado.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	'
	'Crear Archivo Array Json
	'
	sTabla=""

    if IsArray(arrEstado) then

        For i = 0 to ubound(arrEstado, 2)
            '
			sTabla     =  chr(123) &  chr(34) & "Id" 	 & chr(34) & ":" & arrEstado(0,i) & chr(44)
            sTabla     =  sTabla   &  chr(34) & "Name"   & chr(34) & ":" & chr(34) & arrEstado(1,i)  & chr(34) & chr(125) & chr(44)
            sTablaJson =  sTablaJson & sTabla
            sTabla=""
            '
        next

    else
        'Eof()
        sTabla    =   chr(123)&  chr(34) & "Id" 	& chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
        sTabla    =   sTabla &   chr(34) & "Name"   & chr(34)& ":" & chr(34) & "No Aplica" 	& chr(34) & chr(125) & chr(44)
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
	rsEstado.Close
	Set rsEstado = Nothing
	'
	conexion.close
	set conexion = nothing
	'
%>