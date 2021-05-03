<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	Encuesta=Request.QueryString("num")

	dim gDatosSol0
	dim rsx0
	set rsx0 = CreateObject("ADODB.Recordset")
	rsx0.CursorType = adOpenKeyset 
	rsx0.LockType = 2 'adLockOptimistic 
	'Buscar Area
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_GAreaEstado.Id_Estado, "
	sql = sql & " PH_GArea.Area "
	sql = sql & " FROM PH_GAreaEstado INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area "
	sql = sql & " ORDER "
	sql = sql & " BY PH_GAreaEstado.Id_Estado "
	'response.write "<br>36 sql:=" & sql
	'response.end
	rsx0.Open sql ,conexion
	if rsx0.eof then
		rsx0.close
	else
		gDatosSol0 = rsx0.GetRows
		rsx0.close
	end if
	dim gArea(50)
	for iReg = 0 to ubound(gDatosSol0,2)
		iEstado = gDatosSol0(0,iReg)
		sArea = gDatosSol0(1,iReg)
		gArea(iEstado) = sArea
	next 

	dim gDatosSol1
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_EncuestaEspecialResultados.Id_Hogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_EncuestaEspecial.EncuestaEspecial, "
	sql = sql & " PH_EncuestaEspecialDet.Pregunta, "
	sql = sql & " PH_EncuestaEspecialResultados.Id_Respuesta, "
	sql = sql & " PH_EncuestaEspecialResultados.RespuestaTexto, "
	sql = sql & " ss_Estado.id_Estado, "
	sql = sql & " PH_PanelHogar.ClaseSocial "
	sql = sql & " FROM ((((PH_EncuestaEspecialResultados INNER JOIN PH_PanelHogar ON PH_EncuestaEspecialResultados.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_EncuestaEspecial ON PH_EncuestaEspecialResultados.Id_EncuestaEspecial = PH_EncuestaEspecial.Id_EncuestaEspecial) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_EncuestaEspecialDet ON (PH_EncuestaEspecialResultados.Id_Pregunta_Encuesta = PH_EncuestaEspecialDet.Id_EncuestaEspecialDet) AND (PH_EncuestaEspecialResultados.Id_EncuestaEspecial = PH_EncuestaEspecialDet.Id_EncuestaEspecial)) INNER JOIN PH_EncuestaHogar ON (PH_EncuestaHogar.Id_Hogar = PH_EncuestaEspecialResultados.Id_Hogar) AND (PH_EncuestaEspecial.Id_EncuestaEspecial = PH_EncuestaHogar.Id_EncuestaEspecial) "
	sql = sql & " WHERE "
	sql = sql & " PH_EncuestaEspecial.Id_EncuestaEspecial = " & Encuesta
	sql = sql & " AND PH_EncuestaHogar.Ind_Realizada = 1 "
	sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
	sql = sql & " ORDER BY "
	sql = sql & " PH_EncuestaEspecialResultados.Id_Hogar, "
	sql = sql & " PH_EncuestaEspecialDet.Orden "
	'response.write "<br>36 sql:=" & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
	else
		gDatosSol1 = rsx1.GetRows
		rsx1.close
	end if
    rsx1.Open sql ,conexion
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
		iExiste = 1
	end if

	Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
	Response.ContentType = "application/vnd.ms-excel"

	Response.write "<table>"
		Response.write "<tr>"
			Response.write "<td>Id Hogar</td>"
			Response.write "<td>Hogar</td>"
			Response.write "<td>Estado</td>"
			Response.write "<td>Area</td>"
			Response.write "<td>ClaseSocial</td>"
			Response.write "<td>Encuesta</td>"
			Response.write "<td>Pregunta</td>"
			Response.write "<td>IdRespuesta</td>"
			Response.write "<td>Respuesta</td>"
		Response.write "</tr>"
		for iReg = 0 to ubound(gDatosSol1,2)
			Response.write "<tr>"
				Response.write "<td>" & gDatosSol1(0,iReg) & "</td>"
				Response.write "<td>" & gDatosSol1(1,iReg) & "</td>"
				Response.write "<td>" & gDatosSol1(2,iReg) & "</td>"
				idEstado = gDatosSol1(7,iReg)
				Response.write "<td>" & gArea(idEstado) & "</td>"
				Response.write "<td>" & gDatosSol1(8,iReg) & "</td>"
				Response.write "<td>" & gDatosSol1(3,iReg) & "</td>"
				Response.write "<td>" & gDatosSol1(4,iReg) & "</td>"
				Response.write "<td>" & gDatosSol1(5,iReg) & "</td>"
				sx = replace(gDatosSol1(6,iReg),"_"," ")
				Response.write "<td>" & sx & "</td>"
				Response.flush
			Response.write "</tr>"
		next
	Response.write "</table>"
	
%>