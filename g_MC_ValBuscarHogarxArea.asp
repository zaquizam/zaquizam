<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	'g_MC_ValBuscarHogarxArea.asp 26feb21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'
	Dim rsHogar, idEstado, arrHogar, idSemana, idArea
	'
	idEstado = Request.QueryString("id_Estado")
	idSemana = Request.QueryString("id_Semana")
	idArea	 = Request.QueryString("id_Area")
	idMostrar= Request.QueryString("id_Mostrar")
	'		
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_PanelHogar.Id_PanelHogar,"
	QrySql = QrySql & " PH_PanelHogar.CodigoHogar"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo"
	QrySql = QrySql & " INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar"
	QrySql = QrySql & " INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo.Status_registro = 'G' AND"
	'Filtrar por Validados o Pendientes
	'
	If CInt(idMostrar)=1 then
		'Pendientes=1
		QrySql = QrySql & " PH_Consumo.Validado = 0 AND"
	ElseIf CInt(idMostrar)=2 then
		'Validados=2
		QrySql = QrySql & " PH_Consumo.Validado = 1 AND"
	End if
	'
	QrySql = QrySql & " PH_Consumo.Id_Semana = " & idSemana
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_GAreaEstado.Id_Area = " & idArea
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_GAreaEstado.Id_Estado = " & idEstado
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_PanelHogar.Id_PanelHogar,"
	QrySql = QrySql & " PH_PanelHogar.CodigoHogar"
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
	Response.ContentType = "application/json"
	'
	'Crear Archivo Array Json
	'
	sTabla=""

    if IsArray(arrHogar) then
		'
        For i = 0 to ubound(arrHogar, 2)
            '
        	sTabla     =  chr(123) &  chr(34) & "Id" 	 & chr(34) & ":" & arrHogar(0,i) & chr(44)
            sTabla     =  sTabla   &  chr(34) & "Name"   & chr(34) & ":" & chr(34) & arrHogar(0,i) & " - " & arrHogar(1,i)  & chr(34) & chr(125) & chr(44)
            sTablaJson =  sTablaJson & sTabla
            sTabla=vbnullstring
            '
        Next
		'
    else
        'Eof()
        sTabla    =   chr(123)&  chr(34) & "Id" 	& chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
        sTabla    =   sTabla &   chr(34) & "Name"   & chr(34)& ":" & chr(34) & "No Aplica" 	& chr(34) & chr(125) & chr(44)
        '
        sTablaJson = sTablaJson & sTabla
        sTabla=""
		'
    end if
	'
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