<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	Encuesta=Request.QueryString("num")

	Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
	Response.ContentType = "application/vnd.ms-excel"

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
	dim gArea(24)
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

	dim gDatosSol2
	dim rsx2
	set rsx2 = CreateObject("ADODB.Recordset")
	rsx2.CursorType = adOpenKeyset 
	rsx2.LockType = 2 'adLockOptimistic 

	dim gDatosSol3
	dim rsx3
	set rsx3 = CreateObject("ADODB.Recordset")
	rsx3.CursorType = adOpenKeyset 
	rsx3.LockType = 2 'adLockOptimistic 
	

	Response.write "<table>"
		Response.write "<tr>"
			Response.write "<td></td>"
			Response.write "<td></td>"
			Response.write "<td>Respondidas</td>"
			Response.write "<td>Rechazadas</td>"
			Response.write "<td>Pendientes</td>"
			Response.write "<td>% de Cumplimiento</td>"
		Response.write "</tr>"
		Response.write "<tr>"
			Response.write "<td>Total Venezuela</td>"
			Response.write "<td>Total Venezuela</td>"
			'Realizadas
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Count(PH_EncuestaHogar.Id_Hogar) AS CuentaDeId_Hogar "
			sql = sql & " FROM PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar "
			sql = sql & " WHERE PH_EncuestaHogar.Ind_Realizada = 1 AND PH_PanelHogar.Ind_Activo = 1 "
			sql = sql & " GROUP BY PH_EncuestaHogar.Id_EncuestaEspecial "
			sql = sql & " HAVING PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta 
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			if rsx1.eof then
				Response.write "<td>0</td>"
				rsx1.close
			else
				gDatosSol1 = rsx1.GetRows
				rsx1.close
				Response.write "<td>" & gDatosSol1(0,0) & "</td>"
				Realizadas = cint(gDatosSol1(0,0))
			end if
			'Rechazadas
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Count(PH_EncuestaHogar.Id_Hogar) AS CuentaDeId_Hogar "
			sql = sql & " FROM PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar "
			sql = sql & " WHERE PH_EncuestaHogar.Ind_Rechazada = 1 AND PH_PanelHogar.Ind_Activo = 1 "
			sql = sql & " GROUP BY PH_EncuestaHogar.Id_EncuestaEspecial "
			sql = sql & " HAVING PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta 
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			if rsx1.eof then
				Response.write "<td>0</td>"
				rsx1.close
			else
				gDatosSol1 = rsx1.GetRows
				rsx1.close
				Response.write "<td>" & gDatosSol1(0,0) & "</td>"
				Rechazadas = gDatosSol1(0,0)
			end if
			'Pendientes
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " Count(PH_EncuestaHogar.Id_Hogar) AS CuentaDeId_Hogar "
			sql = sql & " FROM PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar "
			sql = sql & " WHERE PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta
			sql = sql & " AND PH_EncuestaHogar.Ind_Rechazada =0  AND PH_EncuestaHogar.Ind_Realizada = 0  AND PH_PanelHogar.Ind_Activo = 1 "
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			if rsx1.eof then
				Response.write "<td>0</td>"
				rsx1.close
			else
				gDatosSol1 = rsx1.GetRows
				rsx1.close
				Response.write "<td>" & gDatosSol1(0,0) & "</td>"
				Pendientes = cint(gDatosSol1(0,0))
			end if
			Total = Pendientes + Realizadas + Rechazadas
			'Cumplimiento = (Realizadas / Pendientes) * 100
			Cumplimiento = (Realizadas * 100) / Total
			Response.write "<td>" & formatnumber(Cumplimiento) & "</td>"
		Response.write "</tr>"
		Response.write "<tr>"
			'response.write "<br>277 Paso<br>"
			'Pendientes
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " ss_Estado.Estado, "
			sql = sql & " Count(PH_EncuestaHogar.Id_EncuestaHogar) AS CuentaDeId_EncuestaHogar, "
			sql = sql & " ss_Estado.Id_Estado "
			sql = sql & " FROM (PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado "
			sql = sql & " WHERE "
			sql = sql & " PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta  
			sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
			'sql = sql & " AND PH_EncuestaHogar.Ind_Rechazada = 0 "
			'sql = sql & " AND PH_EncuestaHogar.Ind_Realizada = 0 "
			sql = sql & " GROUP BY "
			sql = sql & " ss_Estado.Estado, "
			sql = sql & " ss_Estado.Id_Estado "
			sql = sql & " ORDER BY "
			sql = sql & " ss_Estado.Estado "
			'response.write "<br>36 sql:=" & sql
			'response.end
			rsx1.Open sql ,conexion
			if rsx1.eof then
				rsx1.close
				'response.write "<br>300 Paso<br>"
			else
				gDatosSol1 = rsx1.GetRows
				rsx1.close
				'response.write "<br>304 Paso<br>"
			end if
			for iReg = 0 to ubound(gDatosSol1,2)
				Response.write "<tr>" 
					Estado = gDatosSol1(0,iReg)
					idEstado = gDatosSol1(2,iReg)
					Response.write "<td>" & gDatosSol1(0,iReg) & "</td>"
					Response.write "<td>" & gArea(idEstado) & "</td>"
					'Respondidas
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Count(PH_EncuestaHogar.Id_EncuestaHogar) AS CuentaDeId_EncuestaHogar "
					sql = sql & " FROM (PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado "
					sql = sql & " WHERE "
					sql = sql & " PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta
					sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
					sql = sql & " AND ss_Estado.Estado = '" & Estado & "'" 
					sql = sql & " AND PH_EncuestaHogar.Ind_Realizada = 1 "
					'response.write "<br>36 sql:=" & sql
					'response.end
					rsx2.Open sql ,conexion
					if rsx2.eof then
						rsx2.close
						Response.write "<td></td>"
						Realizadas = 0
					else
						gDatosSol2 = rsx2.GetRows
						rsx2.close
						Response.write "<td>" & gDatosSol2(0,0) & "</td>"
						Realizadas = cint(gDatosSol2(0,0))
					end if
					'Rechazadas
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Count(PH_EncuestaHogar.Id_EncuestaHogar) AS CuentaDeId_EncuestaHogar "
					sql = sql & " FROM (PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado "
					sql = sql & " WHERE "
					sql = sql & " PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta
					sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
					sql = sql & " AND ss_Estado.Estado = '" & Estado & "'" 
					sql = sql & " AND PH_EncuestaHogar.Ind_Rechazada = 1 "
					'response.write "<br>36 sql:=" & sql
					'response.end
					rsx3.Open sql ,conexion
					if rsx3.eof then
						rsx3.close
						Response.write "<td></td>"
						Rechazadas = 0
					else
						gDatosSol3 = rsx3.GetRows
						rsx3.close
						Response.write "<td>" & gDatosSol3(0,0) & "</td>"
						Rechazadas = gDatosSol3(0,0)
					end if
					'Pendientes
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Count(PH_EncuestaHogar.Id_EncuestaHogar) AS CuentaDeId_EncuestaHogar "
					sql = sql & " FROM (PH_EncuestaHogar INNER JOIN PH_PanelHogar ON PH_EncuestaHogar.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado "
					sql = sql & " WHERE "
					sql = sql & " PH_EncuestaHogar.Id_EncuestaEspecial = " & Encuesta
					sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
					sql = sql & " AND ss_Estado.Estado = '" & Estado & "'" 
					sql = sql & " AND PH_EncuestaHogar.Ind_Rechazada = 0 "
					sql = sql & " AND PH_EncuestaHogar.Ind_Realizada= 0 "
					'response.write "<br>36 sql:=" & sql
					'response.end
					rsx3.Open sql ,conexion
					if rsx3.eof then
						rsx3.close
						'Response.write "<td></td>"
						Pendientes = 0
					else
						gDatosSol4 = rsx3.GetRows
						rsx3.close
						Pendientes = cint(gDatosSol4(0,0))
					end if
					'Response.write "<td>" & gDatosSol1(1,iReg) & "</td>"
					Response.write "<td>" & Pendientes & "</td>"
					'Pendientes = cint(gDatosSol1(1,iReg))
					Total = Pendientes + Realizadas + Rechazadas
					Cumplimiento = (Realizadas * 100) / Total
					'Cumplimiento = (Realizadas / Pendientes) * 100
					Response.write "<td>" & formatnumber(Cumplimiento) & "</td>"
					'Response.write "<td>"
					'Response.write "Pendientes:" & Pendientes
					'Response.write "Realizadas:" & Realizadas
					'Response.write "Rechazadas:" & Rechazadas
					'Response.write "</td>"
					
				Response.write "</tr>"
			next
		
		
		Response.write "</tr>"
	Response.write "</table>"
%>