<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_MC_ValBuscarDetalleConsumoxDiaSemana.asp - 26feb21
	'
	' se le agrego que solo mostrara los consumos no investigados
	'
	Session.lcid 		= 1034
	Response.CodePage = 65001
	Response.CharSet 	= "utf-8"	
	'
	Dim rsConsumos, idArea, arrConsumos
	'
	idHogar 		= Request.QueryString("id_Hogar")
	idTipoConsumo	= Request.QueryString("id_TipoConsumo")
	idSemana 	    = Request.QueryString("id_Semana")	
	idFecha 		= Request.QueryString("id_Fecha")
	idMostrar	    = Request.QueryString("id_Mostrar")
	'
	' Buscar Los Estados asociados al Area
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT ROW_NUMBER() OVER(ORDER BY PH_Consumo.Id_Consumo ASC) AS Item,"
	QrySql = QrySql & " PH_Consumo.Id_Consumo,"
	QrySql = QrySql & " (CASE DATENAME(dw,fecha_creacion) when 'Monday' then 'LUNES' when 'Tuesday' then 'MARTES' when 'Wednesday' then 'MIERCOLES' when 'Thursday' then 'JUEVES' when 'Friday' then 'VIERNES' when 'Saturday' then 'SABADO' when 'Sunday' then 'DOMINGO' END) AS DIA,"
	QrySql = QrySql & " FORMAT (PH_Consumo.fecha_creacion, 'dd-MM-yyyy ') AS FECHA,"
	QrySql = QrySql & " PH_Consumo.Validado,"
	QrySql = QrySql & " PH_Consumo.Enviado_investigar,"
	QrySql = QrySql & " PH_Consumo.Resuelto"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Consumo"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Consumo.Id_Hogar  = " & idHogar
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Id_Semana = " & idSemana
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Id_Tipoconsumo = " & idTipoConsumo
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Fecha_Creacion = '" & idFecha & "'"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Consumo.Status_registro = 'G'"	
	QrySql = QrySql & " AND"
	'
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
	QrySql = QrySql & " PH_Consumo.Enviado_investigar='0'"	
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
			bValidado = arrConsumos(4,i)
			'			
			sTabla     =  chr(123) &  chr(34) & "Id"   & chr(34) & ":" & chr(34) & arrConsumos(1,i) & chr(34) & chr(44)
			if bValidado = "True" then
				sValidado = "Validado"
				sTabla     =  sTabla   &  chr(34) & "Name" & chr(34) & ":" & chr(34) & Right("00" & arrConsumos(0,i), 2) & ".- " & arrConsumos(2,i) & " - " & trim(arrConsumos(3,i))  & " - " & sValidado & chr(34) & chr(125) & chr(44)
			else
				sValidado = ""
				sTabla     =  sTabla   &  chr(34) & "Name" & chr(34) & ":" & chr(34) & Right("00" & arrConsumos(0,i), 2) & ".- " & arrConsumos(2,i) & " - " & trim(arrConsumos(3,i))  & chr(34) & chr(125) & chr(44)
			end if
			
			sTablaJson =  sTablaJson & sTabla
			sTabla = vbnullstring
			'			
        next

    else
			'Eof()
			sTabla    =   chr(123)&  chr(34) & "Id" 		& chr(34) & ":" & chr(34) & "0" 			 & chr(34) & chr(44)
			sTabla    =   sTabla &   chr(34) & "Name"   	& chr(34) & ":" & chr(34) & "No hay registros" & chr(34) & chr(125) & chr(44)
			'
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
    end if
	''
	sTabla 	= 	Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
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