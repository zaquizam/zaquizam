<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	'
	' g_rRevInvLlenarCmbConsumosInvestigados.asp //  12ene21 - 14ene21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'
	Dim rsHogar, arrHogar, QrySql, idHogar
	'
	idHogar = Request.QueryString("id_Hogar")	
	'
	' Buscar todos consumos investigados por hogar 
	'			
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_Consumo.Id_Consumo,"
	QrySql = QrySql & " ROW_NUMBER() OVER(ORDER BY PH_Consumo.Id_Consumo ASC) AS Item,"	
	QrySql = QrySql & " (CASE DATENAME(dw,fecha_creacion) when 'Monday' then 'LUN' when 'Tuesday' then 'MAR' when 'Wednesday' then 'MIE' when 'Thursday' then 'JUE' when 'Friday' then 'VIE' when 'Saturday' then 'SAB' when 'Sunday' then 'DOM' END) AS DIA,"
	QrySql = QrySql & " FORMAT (PH_Consumo.fecha_creacion, 'dd-MM-yyyy ') AS FECHA,"
	QrySql = QrySql & " SUBSTRING( ss_Semana.Semana, 1, 2) AS semana"	
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo"
	QrySql = QrySql & " INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar= PH_PanelHogar.Id_PanelHogar"
	QrySql = QrySql & " INNER JOIN ss_Semana ON PH_Consumo.Id_Semana= ss_Semana.IdSemana"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo.Enviado_investigar = 1"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Resuelto = 0"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Id_HOGAR = " & idHogar
	'	
	' Response.Write QrySql '& "<BR><BR>"
	' Response.end
	'
	Set rsHogar = Server.CreateObject("ADODB.recordset")
	rsHogar.Open QrySql,conexion
	'
	if not rsHogar.EOF then
    	arrHogar = rsHogar.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	'Response.ContentType = "application/json"
	'
	' Crear Archivo Array Json
	'
	sTabla=""

    if IsArray(arrHogar) then

        For i = 0 to UBound(arrHogar, 2)
            '
			sTabla     =  chr(123) &  chr(34) & "Id" 	  & chr(34) & ":" & cstr(arrHogar(0,i))  & chr(44)
			sTabla     =  sTabla   &  chr(34) & "Nombre"  & chr(34) & ":" & chr(34) &  Right("00" & arrHogar(1,i), 2)   & " - " & arrHogar(2,i) & " - " & arrHogar(3,i)  &  chr(34) & chr(125) & chr(44)
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
	rsHogar.Close
	Set rsHogar = Nothing
	'
	conexion.close
	set conexion = nothing
	'
%>