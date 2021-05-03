<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	'g_ValBuscarTotalConsumoxDiaSemana.asp // 29dic20 - 22ene21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"	
	'
	Dim rsConsumos, idArea, arrConsumos
	'
	idHogar 		= Request.QueryString("id_Hogar")
	idTipoConsumo 	= Request.QueryString("id_TipoConsumo")
	idSemana 		= Request.QueryString("id_Semana")	
	'
	' Buscar Los Estados asociados al Area
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " (CASE DATENAME(dw,fecha_creacion) when 'Monday' then 'LUN' when 'Tuesday' then 'MAR' when 'Wednesday' then 'MIE' when 'Thursday' then 'JUE' when 'Friday' then 'VIE' when 'Saturday' then 'SAB' when 'Sunday' then 'DOM' END) AS DIA,"
	QrySql = QrySql & " PH_Consumo.fecha_creacion," 
	QrySql = QrySql & " FORMAT (PH_Consumo.fecha_creacion, 'dd-MM-yyyy ') AS FECHA," 
	QrySql = QrySql & " Count(PH_Consumo.Fecha_Creacion) AS TOTAL_ROWS" 
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo.Id_Hogar = "& idHogar
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.id_TipoConsumo = "  & idTipoConsumo
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Id_Semana = " & idSemana
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Status_registro = 'G'"	
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Enviado_investigar = 0"	
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Resuelto = 0"	
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " PH_Consumo.Fecha_Creacion"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_Consumo.Fecha_Creacion ASC"
	'
	'Response.Write QrySql 
	'Response.End
	'
	Set rsConsumos = Server.CreateObject("ADODB.recordset")
	rsConsumos.Open QrySql,conexion
	'
	if not rsConsumos.EOF then
    	arrConsumos = rsConsumos.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	'
	'Crear Archivo Array Json
	'
	sTabla=""

    if IsArray(arrConsumos) then

        For i = 0 to ubound(arrConsumos, 2)
            '
			sTabla     =  chr(123) &  chr(34) & "Id" 	 & chr(34) & ":" & chr(34) & arrConsumos(1,i) & chr(34) & chr(44)
            sTabla     =  sTabla   &  chr(34) & "Name"   & chr(34) & ":" & chr(34) & arrConsumos(0,i) &" - " & arrConsumos(2,i) & " - (" & arrConsumos(3,i) & ")" & chr(34) & chr(125) & chr(44)
            sTablaJson =  sTablaJson & sTabla
            sTabla=""
            '
        next

    else
        'Eof()
        sTabla    =   chr(123)&  chr(34) & "Id" 	& chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
        sTabla    =   sTabla &   chr(34) & "Name"   & chr(34)& ":" & chr(34) & "No hay registros" 	& chr(34) & chr(125) & chr(44)
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
	rsConsumos.Close
	Set rsConsumos = Nothing
	'
	conexion.close
	set conexion = nothing
	'
%>