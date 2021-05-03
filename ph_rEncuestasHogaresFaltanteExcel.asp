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

	dim gDatosSol1
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_GArea.Area, ss_Estado.Estado, "
	sql = sql & " PH_EncuestaHogar.Id_Hogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Celular "
	sql = sql & " FROM (((PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN (PH_GArea INNER JOIN PH_GAreaEstado ON PH_GArea.Id_Area = PH_GAreaEstado.Id_Area) ON (ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado) AND (PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) "
	sql = sql & " WHERE "
	sql = sql & " PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta
	sql = sql & " AND PH_EncuestaHogar.Ind_Realizada = 0 "
	sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 " 
	sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
	sql = sql & " ORDER BY "
	sql = sql & " PH_GArea.Area, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_EncuestaHogar.Id_Hogar "
	'response.write "<br>36 sql:=" & sql
	'response.end
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
			Response.write "<td>Area</td>"
			Response.write "<td>Estado</td>"
			Response.write "<td>Id Hogar</td>"
			Response.write "<td>Hogar</td>"
			Response.write "<td>Nombre</td>"
			Response.write "<td>Apellido</td>"
			Response.write "<td>celular</td>"
		Response.write "</tr>"
		for iReg = 0 to ubound(gDatosSol,2)
			Response.write "<tr>"
				for ib = 0 to 6
					Response.write "<td>" & gDatosSol(ib,iReg) & "</td>"
				next
			Response.write "</tr>"
		next
	Response.write "</table>"
	
%>