<!--#include file="conexion.asp"-->
<%
	' ph_mEncuestaEspecialResultadosData.asp // 01sep21 - 02sep21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idEncuesta, idOpcion
	'	
	idEncuesta	=	Request.Querystring("id_Encuesta")
	idOpcion   	=	Request.Querystring("id_Opcion")
	
	if idOpcion = 1 then 	
		'
		' Buscar los detalles de la Encuesta
		'
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT"
		QrySql = QrySql & " PH_GArea.Area,"
		QrySql = QrySql & " ss_Estado.Estado,"
		QrySql = QrySql & " PH_EncuestaHogar.Id_Hogar,"
		QrySql = QrySql & " PH_PanelHogar.CodigoHogar,"
		QrySql = QrySql & " PH_Panelistas.Nombre1,"
		QrySql = QrySql & " PH_Panelistas.Apellido1,"
		QrySql = QrySql & " PH_Panelistas.Celular"
		QrySql = QrySql & " FROM"
		QrySql = QrySql & " PH_EncuestaHogar"
		QrySql = QrySql & " INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar"
		QrySql = QrySql & " INNER JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar"
		QrySql = QrySql & " INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado"
		QrySql = QrySql & " INNER JOIN PH_GArea"
		QrySql = QrySql & " INNER JOIN PH_GAreaEstado ON PH_GArea.Id_Area = PH_GAreaEstado.Id_Area ON ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado"
		QrySql = QrySql & " AND"
		QrySql = QrySql & " PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " PH_EncuestaHogar.Id_EncuestaEspecial =" & idEncuesta
		QrySql = QrySql & " AND"
		QrySql = QrySql & " PH_EncuestaHogar.Ind_Realizada = 1"
		QrySql = QrySql & " AND"
		QrySql = QrySql & " PH_Panelistas.ResponsablePanel = 1"
		QrySql = QrySql & " AND"
		QrySql = QrySql & " PH_PanelHogar.Ind_Activo = 1"
		QrySql = QrySql & " ORDER BY"
		QrySql = QrySql & " PH_GArea.Area,"
		QrySql = QrySql & " ss_Estado.Estado,"
		QrySql = QrySql & " PH_EncuestaHogar.Id_Hogar"
		'		
		Set rsHogares = Server.CreateObject("ADODB.recordset")
		 rsHogares.Open QrySql, conexion
		' '
		 if not rsHogares.EOF then
			arrHogares = rsHogares.GetRows()  ' Convert recordset to 2D Array
			iTotalHogares = UBound(arrHogares, 2) + 1 	  
		else
	  		iTotalHogares = 0		
		end if
		' '
		' Response.ContentType = "application/json"		
		' '
		' sTabla=vbnullstring
		
		' if IsArray(arrHogares) then
		
			' For i = 0 to ubound(arrHogares, 2)
				' sTabla = chr(123) &  chr(34) & "id"	    & chr(34) & ":" & chr(34) & arrHogares(2,i) & chr(34) & chr(44)
				' sTabla = sTabla   &  chr(34) & "nombre" & chr(34) & ":" & chr(34) & arrHogares(0,i) & " - " & arrHogares(1,i) & " - " & arrHogares(2,i) & " - " & arrHogares(3,i) & " - " & arrHogares(4,i) & " - " & arrHogares(5,i) & " - " & arrHogares(6,i) & chr(34) & chr(125) & chr(44)
				
				' sTablaJson = sTablaJson & sTabla
				' sTabla=vbnullstring
			' next
			' '
			' sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
			' JsonData = chr(91) & sTabla & chr(93) '& chr(125)
			' '
		' else
			' 'Eof()
			' sTablaJson = sTablaJson & sTabla
			' sTabla = vbnullstring
			' JsonData = chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
		' end if
		' ''
		' Response.Write(JsonData)
		' conexion.close
		' set conexion = nothing	
		'
		Response.Write iTotalHogares
		'	
	else
		Response.Write 0
	end if
%>