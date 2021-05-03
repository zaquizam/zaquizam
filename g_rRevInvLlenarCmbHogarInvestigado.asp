<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_rRevInvLlenarCmbHogarInvestigado.asp //  12ene21 - 03feb21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'
	Dim rsHogar, arrHogar
	'			
	' QrySql = vbnullstring
	' QrySql = QrySql & " SELECT DISTINCT"
	' QrySql = QrySql & " PH_PanelHogar.Id_PanelHogar,"
	' QrySql = QrySql & " PH_PanelHogar.CodigoHogar,"
	' QrySql = QrySql & " PH_TipoConsumo.tipoconsumo,"
	' QrySql = QrySql & " SUBSTRING (ss_Semana.Semana , 1 , 4 )  AS semana,"
	' QrySql = QrySql & " ss_Semana.IdSemana"
	' QrySql = QrySql & " FROM"
	' QrySql = QrySql & " PH_Consumo"
	' QrySql = QrySql & " INNER JOIN PH_PanelHogar  ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar"
	' QrySql = QrySql & " INNER JOIN PH_TipoConsumo ON PH_Consumo.Id_TipoConsumo = PH_TipoConsumo.Id_TipoConsumo"
	' QrySql = QrySql & " INNER JOIN ss_Semana      ON PH_Consumo.Id_Semana = ss_Semana.idSemana"
	' QrySql = QrySql & " WHERE"
	' QrySql = QrySql & " PH_Consumo.Enviado_investigar = 1"
	' QrySql = QrySql & " AND PH_Consumo.Resuelto = 0"
	' QrySql = QrySql & " AND PH_Consumo.validado = 0"
	' QrySql = QrySql & " ORDER BY"
	' QrySql = QrySql & " ss_Semana.IdSemana  ASC,"
	' QrySql = QrySql & " PH_PanelHogar.Id_PanelHogar ASC"
	'
	'O3FEB21
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT DISTINCT"
	QrySql = QrySql & " PH_PanelHogar.Id_PanelHogar,"
	QrySql = QrySql & " PH_PanelHogar.CodigoHogar,"
	QrySql = QrySql & " PH_TipoConsumo.TipoConsumo,"
	QrySql = QrySql & " SUBSTRING(cacevedo_atenas.ss_Semana.Semana,1,4) AS semana,"
	QrySql = QrySql & " ss_Estado.Estado,"
	QrySql = QrySql & " ss_Semana.IdSemana"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_PanelHogar"
	QrySql = QrySql & " INNER JOIN PH_Consumo ON PH_PanelHogar.Id_PanelHogar =PH_Consumo.Id_Hogar"
	QrySql = QrySql & " INNER JOIN PH_TipoConsumo ON PH_Consumo.id_TipoConsumo =PH_TipoConsumo.Id_TipoConsumo"
	QrySql = QrySql & " INNER JOIN ss_Semana ON PH_Consumo.Id_Semana =ss_Semana.IdSemana"
	QrySql = QrySql & " INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado =ss_Estado.Id_Estado"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo.Enviado_investigar = 1"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Validado = 0"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Resuelto = 0"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " ss_Estado.Estado ASC,"
	QrySql = QrySql & " ss_Semana.IdSemana ASC"
	'	
	'Response.Write QrySql
	'Response.end
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
	sTabla = vbnullstring

    if IsArray(arrHogar) then

        For i = 0 to UBound(arrHogar, 2)
            '
			sTabla     =  chr(123) &  chr(34) & "Id" 	  & chr(34) & ":" & cstr(arrHogar(0,i))  & chr(44)
            sTabla     =  sTabla   &  chr(34) & "Nombre"  & chr(34) & ":" & chr(34) & arrHogar(0,i) & " - " & arrHogar(1,i) & " - " & arrHogar(2,i) & " - Sem: " & arrHogar(3,i) & " - " & arrHogar(4,i) & chr(34) & chr(125) & chr(44)
            sTablaJson =  sTablaJson & sTabla
            sTabla = vbnullstring
            '
        next

    else
        'Eof()
        sTabla     =   chr(123)&  chr(34) & "Id" 	& chr(34)& ":" & chr(34) & "0" 					& chr(34) & chr(44)
        sTabla     =   sTabla &   chr(34) & "Name"   & chr(34)& ":" & chr(34) & "No hay Registros" 	& chr(34) & chr(125) & chr(44)
        sTablaJson = sTablaJson & sTabla
        sTabla = vbnullstring

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