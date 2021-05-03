<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%

  
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	dim gDatosSol1
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	dim gSemanas
	dim gTipoConsumo
	dim gHogares
	dim gCantidad
	dim idArea
	dim idEstado
	idTipoConsumo=Request.QueryString("tip")
	idArea=Request.QueryString("are")
	idEstado=Request.QueryString("est")
	'response.write "<br>idTipoConsumo:= " & idTipoConsumo
	'response.end 
	dim rsx2
	set rsx2 = CreateObject("ADODB.Recordset")
	rsx2.CursorType = adOpenKeyset 
	rsx2.LockType = 2 'adLockOptimistic 

	'Semana 
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Top 4 "
	sql = sql & " IdSemana, "
	sql = sql & " Semana "
	sql = sql & " FROM "
	sql = sql & " ss_Semana "
	sql = sql & " ORDER BY "
	sql = sql & " IdSemana DESC "
	'response.write "<br>36 sql:=" & sql
	'response.end
	rsx2.Open sql ,conexion
	if rsx2.eof then
		rsx2.close
	else 
		gSemanas = rsx2.GetRows
		rsx2.close
	end if

	'Tipos de Consumo 
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_TipoConsumo, "
	sql = sql & " TipoConsumo "
	sql = sql & " FROM "
	sql = sql & " PH_TipoConsumo "
	sql = sql & " Where "
	sql = sql & " Ind_Activo = 1 "
	sql = sql & " and Id_TipoConsumo = " & idTipoConsumo
	sql = sql & " ORDER BY "
	sql = sql & " Id_TipoConsumo "
	'response.write "<br>36 sql:=" & sql
	'response.end
	rsx2.Open sql ,conexion
	if rsx2.eof then
		rsx2.close
	else 
		gTipoConsumo = rsx2.GetRows
		rsx2.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " PH_GArea.Area, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Celular "
	sql = sql & " FROM (((PH_PanelHogar INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_GAreaEstado ON ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar "
	sql = sql & " WHERE "
	sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
	sql = sql & " AND PH_PanelHogar.Id_PanelHogar > 1 "
	sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 "
	if idArea <> 0 then
		sql = sql & " AND PH_GAreaEstado.Id_Area = " & idArea
	end if
	if idEstado <> 0 then
		sql = sql & " AND  PH_PanelHogar.Id_Estado = " & idEstado
	end if
	
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 35 "
	'response.write "<br>36 sql:=" & sql
	'response.end
	rsx2.Open sql ,conexion
	if rsx2.eof then
		rsx2.close
	else 
		gHogares = rsx2.GetRows
		rsx2.close
	end if

	Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
	Response.ContentType = "application/vnd.ms-excel"
	
	'response.write "<br>llego"
	'response.end
	%>
	<div id="DivBuscarInformación">
		<div class="ex1">
			<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
				<thead>
					<tr class="w3-blue">
						<th>IdHogar</th>
						<th>CodHogar</th>
						<th>Area</th>
						<th>Estado</th>
						<th>Nombre</th>
						<th>Apellido</th>
						<th>Celular</th>
						<th>Tipo Consumo</th>
						<%
						for iReg = 0 to ubound(gSemanas,2)
							response.write "<th>" & gSemanas(1,iReg) & "</th>"
						next 
						for iReg = 0 to ubound(gHogares,2)
							Response.flush
							response.write "<tr>"
								idHogar = gHogares(0,iReg)
								for iCol = 0 to 6
									response.write "<td>" & gHogares(iCol,iReg) & "</td>"
								next
								isw = 1
								for iReg2 = 0 to ubound(gTipoConsumo,2)
									idTipoConsumo = gTipoConsumo(0,iReg2)
									if isw = 0 then
										response.write "<td></td>"
										response.write "<td></td>"
										response.write "<td></td>"
										response.write "<td></td>"
										response.write "<td></td>"
										response.write "<td></td>"
										response.write "<td></td>"
									end if
									response.write "<td>" & gTipoConsumo(1,iReg2) & "</td>"
									for iReg3 = 0 to 3
										idSemana = gSemanas(0,iReg3)
										'response.write "<br>Semana:= " & idSemana
										'Consumos
										sql = ""
										sql = sql & " SELECT "
										sql = sql & " Count(Id_Consumo) AS Total "
										sql = sql & " FROM "
										sql = sql & " PH_Consumo "
										sql = sql & " WHERE "
										sql = sql & " Id_Semana = " & idSemana
										sql = sql & " AND Id_Hogar = " & idHogar
										sql = sql & " AND id_TipoConsumo = " & idTipoConsumo
										'response.write "<br>36 sql:=" & sql
										'response.end
										rsx1.Open sql ,conexion
										if rsx1.eof then
											rsx1.close
											Cantidad = 0
										else 
											gCantidad = rsx1.GetRows
											rsx1.close
											Cantidad = gCantidad(0,0)
										end if
										response.write "<td>" & Cantidad & "</td>"
									next 
									response.write "</tr>"
									isw = 0
								next
							response.write "</tr>"
							Response.flush
						next 
						%>
					</tr>
				</thead>

			</table>
		</div>
	</div>
	<%
	conexion.close
	%>
</body>
</html>