<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
	'response.write "<br> Linea 6 " 
	'response.end
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idTipoConsumo, idArea, idEstado, idSemana
	'
	idTipoConsumo = cint(Request.QueryString("num"))
	idArea = Request.QueryString("are")
	idEstado = Request.QueryString("est")
	idSemana = Request.QueryString("sem")

	'response.write "<br> Linea 18 " 
	'response.end
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
	
	'response.write "<br> Linea 27 " 
	'response.end

	'# Hogares que Reportaron
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " PH_GArea.Area, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Celular "
	sql = sql & " FROM ((PH_PanelHogar INNER JOIN (PH_GAreaEstado INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar "
	sql = sql & " WHERE "
	sql = sql & " PH_PanelHogar.Ind_Activo = 1 " 
	sql = sql & " AND PH_Panelistas.ResponsablePanel =  1 "
	if idArea <> 0 then 
		sql = sql & " AND PH_GAreaEstado.Id_Area = " & idArea
	end if
	if idEstado <> 0 then
		sql = sql & " AND  PH_PanelHogar.Id_Estado = " & idEstado
	end if
	sql = sql & " GROUP BY "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " PH_GArea.Area, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Celular "
	sql = sql & " HAVING "
	sql = sql & " PH_PanelHogar.Id_PanelHogar > 1 "
	'response.write "<br> Linea 49 " & sql
	'response.end
	rsx1.Open sql ,conexion
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
	else
		gDatosSol1 = rsx1.GetRows
		rsx1.close
		iExiste = 1
	end if
	%>
	<div id="Reporte"> 
		<div class="ex1">
			<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
				<thead>
					<tr class="w3-blue">
						<th>idHogar</th>
						<th>CodHogar</th>
						<th>Area</th>
						<th>Estado</th>
						<th>Nombre</th>
						<th>Apellido</th>
						<th>Celular</th>
					</tr>
				</thead>
				<%
				for iReg = 0 to ubound(gDatosSol1,2)
					idHogar = gDatosSol1(0,iReg)
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " Count(PH_Consumo.Id_Consumo) AS Total "
					sql = sql & " FROM PH_Consumo "
					sql = sql & " WHERE "
					sql = sql & " PH_Consumo.Id_Hogar = " & idHogar
					sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
					sql = sql & " AND PH_Consumo.Id_Semana = " & idSemana
					rsx2.Open sql ,conexion
					iExiste = 0
					if rsx2.eof then
						iExiste = 0
					else
						gDatosSol2 = rsx2.GetRows
						rsx2.close
						iExiste = 1
						if cint(gDatosSol2(0,0)) = 0 then
							Response.write "<tr>"
								for iCol = 0 to 6
									Response.write "<td>" &  gDatosSol1(iCol,iReg) & "</td>"
								next 
								'Response.write "<td>" &  gDatosSol2(0,0) & "</td>"
							Response.write "</tr>"
						end if
					end if
					
				next
				%>
			</table>
		</div>
	</div> 
	<%


	'response.write "<br> Linea Final " 
	'response.end
	
%>
	
	
	
