<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_rRevInvLlenarCmbHogarInvestigado.asp //  12ene21 - 14ene21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'
	Dim rsHogar, arrHogar
	'
	' Buscar Los Estados asociados al Area
	'	
	' QrySql = vbnullstring
	' QrySql = QrySql & " SELECT"
	' QrySql = QrySql & " PH_Consumo.Id_Consumo AS id,"
	' QrySql = QrySql & " PH_PanelHogar.Id_PanelHogar AS idhogar,"
	' QrySql = QrySql & " PH_PanelHogar.CodigoHogar AS nombre"
	' QrySql = QrySql & " FROM"
	' QrySql = QrySql & " PH_Consumo"
	' QrySql = QrySql & " INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar"
	' QrySql = QrySql & " WHERE"
	' QrySql = QrySql & " PH_Consumo.Enviado_investigar = 1"
	' QrySql = QrySql & " AND"
	' QrySql = QrySql & " PH_Consumo.Resuelto = 0"
	' QrySql = QrySql & " ORDER BY"
	' QrySql = QrySql & " PH_Consumo.Id_Hogar ASC"
	' '15ene21
	' QrySql = vbnullstring
	' QrySql = QrySql & " SELECT"
	' QrySql = QrySql & " PH_Consumo.Id_Consumo,"
	' QrySql = QrySql & " PH_PanelHogar.Id_PanelHogar,"
	' QrySql = QrySql & " PH_PanelHogar.CodigoHogar,"
	' QrySql = QrySql & " SUBSTRING ( ss_Semana.Semana , 1 , 2 ) AS semana"
	' QrySql = QrySql & " FROM"
	' QrySql = QrySql & " PH_Consumo"
	' QrySql = QrySql & " INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar"
	' QrySql = QrySql & " INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana"
	' QrySql = QrySql & " WHERE"
	' QrySql = QrySql & " PH_Consumo.Enviado_investigar = 1"
	' QrySql = QrySql & " AND PH_Consumo.Resuelto = 0"
	' QrySql = QrySql & " ORDER BY"
	' QrySql = QrySql & " PH_Consumo.Id_Hogar ASC,"
	' QrySql = QrySql & " PH_Consumo.Id_Consumo ASC,"
	' QrySql = QrySql & " PH_Consumo.Id_Semana ASC"
	' '15ene21
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT DISTINCT"
	QrySql = QrySql & " PH_PanelHogar.Id_PanelHogar,"
	QrySql = QrySql & " PH_PanelHogar.CodigoHogar,"
	QrySql = QrySql & " PH_Consumo.Enviado_investigar"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo"
	QrySql = QrySql & " INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo.Enviado_investigar = 1"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Resuelto = 0"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_PanelHogar.Id_PanelHogar ASC"
	'	
	'Response.Write QrySql '& "<BR><BR>"
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
	sTabla=""

    if IsArray(arrHogar) then

        For i = 0 to UBound(arrHogar, 2)
            '
			sTabla     =  chr(123) &  chr(34) & "Id" 	  & chr(34) & ":" & cstr(arrHogar(0,i))  & chr(44)
            sTabla     =  sTabla   &  chr(34) & "Nombre"  & chr(34) & ":" & chr(34) & arrHogar(1,i) & " - " & arrHogar(2,i) & " - Sem: " & arrHogar(3,i)    & chr(34) & chr(125) & chr(44)
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