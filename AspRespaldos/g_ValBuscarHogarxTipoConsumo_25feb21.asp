<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	'g_ValBuscarHogarxTipoConsumo.asp
	'29dic20
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'
	Dim rsTipoConsumo, idHogar, arrTipoConsumo
	'
	idHogar = Request.QueryString("id_Hogar")
	idSemana = Request.QueryString("id_Semana")
	'
	' Buscar Los Hogares asociados al Estado
	'	
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_TipoConsumo.Id_TipoConsumo,"
	QrySql = QrySql & " PH_TipoConsumo.TipoConsumo"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo INNER JOIN PH_TipoConsumo ON PH_Consumo.id_TipoConsumo = PH_TipoConsumo.Id_TipoConsumo"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo.Id_Hogar= " & idHogar
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Id_Semana= " & idSemana
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Status_registro = 'G' AND"
	QrySql = QrySql & " PH_TipoConsumo.Ind_Activo=1"
	QrySql = QrySql & " GROUP BY PH_TipoConsumo.TipoConsumo, PH_TipoConsumo.Id_TipoConsumo"
	QrySql = QrySql & " ORDER BY PH_TipoConsumo.TipoConsumo;"	
	'	
	'
	'Response.Write QrySql '& "<BR><BR>"
	'Response.end
	'
	Set rsTipoConsumo = Server.CreateObject("ADODB.recordset")
	rsTipoConsumo.Open QrySql,conexion
	'
	if not rsTipoConsumo.EOF then
    	arrTipoConsumo = rsTipoConsumo.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	'
	'Crear Archivo Array Json
	'
	sTabla=""

    if IsArray(arrTipoConsumo) then

        For i = 0 to ubound(arrTipoConsumo, 2)
            '
            'sTabla     =  chr(123) &  chr(34) & "Id" 	 & chr(34) & ":" & chr(34) & arrTipoConsumo(0,i)  & chr(34) & chr(44)
			sTabla     =  chr(123) &  chr(34) & "Id" 	 & chr(34) & ":" & arrTipoConsumo(0,i) & chr(44)
            sTabla     =  sTabla   &  chr(34) & "Name"   & chr(34) & ":" & chr(34) & arrTipoConsumo(1,i)  & chr(34) & chr(125) & chr(44)
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
	rsTipoConsumo.Close
	Set rsTipoConsumo = Nothing
	'
	conexion.close
	set conexion = nothing
	'
%>