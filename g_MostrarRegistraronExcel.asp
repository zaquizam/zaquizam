<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
	'response.write "<br> Linea 6 " 
	'response.end
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	Server.ScriptTimeout=1000
	Response.buffer = true
	
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

	Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
	Response.ContentType = "application/vnd.ms-excel"
	select case idTipoConsumo
		case 1,8
			'# Hogares que Reportaron
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " PH_Consumo.Id_Hogar, PH_PanelHogar.CodigoHogar, PH_GArea.Area, ss_Estado.Estado, PH_Panelistas.Nombre1, PH_Panelistas.Apellido1, PH_Panelistas.Celular, PH_Medio.Medio, PH_Moneda.Moneda, PH_Canal.Canal, PH_Cadena.Cadena, PH_Consumo.Total_Compra, PH_Consumo.Id_Consumo, PH_Consumo.Fecha_Creacion  "
			sql = sql & " FROM ((((((((PH_Consumo INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) INNER JOIN PH_Medio ON PH_Consumo.Id_Medio = PH_Medio.Id_Medio) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal) LEFT JOIN PH_Cadena ON PH_Consumo.Id_Cadena = PH_Cadena.Id_Cadena) INNER JOIN PH_Moneda ON PH_Consumo.Id_Moneda = PH_Moneda.Id_Moneda "
			sql = sql & " WHERE "
			sql = sql & " PH_Consumo.Id_Hogar > 1 "
			sql = sql & " AND PH_Consumo.Id_Semana = " & idSemana 
			sql = sql & " AND PH_PanelHogar.Ind_activo = 1 "
			sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
			sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 "
			if idArea <> 0 then
				sql = sql & " AND PH_GAreaEstado.Id_Area = " & idArea
			end if
			if idEstado <> 0 then
				sql = sql & " AND  PH_PanelHogar.Id_Estado = " & idEstado
			end if
			'sql = sql & " ORDER BY PH_Consumo.Id_Hogar "
			sql = sql & " ORDER BY PH_Consumo.Fecha_Creacion desc "
			'response.write "<br> Linea 49 "  & sql
			'response.end
			rsx1.Open sql ,conexion
			iExiste = 0
			if rsx1.eof then
				iExiste = 0
				rsx1.close
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
								<th>Medio</th>
								<th>Moneda</th>
								<th>Canal</th>
								<th>Cadena</th>
								<th>Total Compra</th>
								<th>Total Prod. Reg.</th>
								<th>Total Cant * Precio</th>
								<th>idConsumo</th>
								<th>Fec. Registro</th>
							</tr>
						</thead>
						<%
						icont = 0
						for iReg = 0 to ubound(gDatosSol1,2)
							Response.write "<tr>"
								icont = icont + 1
								if icont > 50 then 
									Response.flush
									icont = 0
								end if 

								for iCol = 0 to 10
									Response.write "<td>" &  gDatosSol1(iCol,iReg) & "</td>"
								next 
								TotalCompras = gDatosSol1(11,iReg)
								Response.write "<td>" & formatnumber(TotalCompras) & "</td>"
								idConsumo = gDatosSol1(12,iReg)
								idFecConsumo = gDatosSol1(13,iReg)
								idFecConsumo = day(idFecConsumo) & "/" & month(idFecConsumo) & "/" & year(idFecConsumo)
								idHogar = gDatosSol1(0,iReg)
								sql = ""
								sql = sql & " SELECT "
								sql = sql & " Cantidad, "
								sql = sql & " Precio_producto, "
								sql = sql & " Id_Categoria "
								sql = sql & " FROM "
								sql = sql & " PH_Consumo_Detalle_Productos "
								sql = sql & " WHERE "
								sql = sql & " Id_Hogar = " & idHogar
								sql = sql & " AND Id_Consumo = " & idConsumo
								rsx2.Open sql ,conexion
								iExiste = 0
								if rsx2.eof then
									iExiste = 0
									rsx2.close
								else
									gDatosSol2 = rsx2.GetRows
									rsx2.close
									iExiste = 1
								end if
								Cuantos = 0
								Total  = 0
								SumaCuantos = 0
								SumaTotal  = 0
								for iReg1 = 0 to ubound(gDatosSol2,2)
									Cuantos = cint(gDatosSol2(0,iReg1))
									Precio  = cdbl(gDatosSol2(1,iReg1))
									iCat  = cdbl(gDatosSol2(2,iReg1))
									if iCat <> 9 then 
										Total = Cuantos * Precio
									else
										Total = Precio
									end if
									SumaCuantos = SumaCuantos + Cuantos
									SumaTotal = SumaTotal + Total
								next 
								Response.write "<td>" & SumaCuantos & "</td>"
								Response.write "<td>" & formatnumber(SumaTotal) & "</td>"
								Response.write "<td>" & idConsumo & "</td>"
								Response.write "<td>" & idFecConsumo & "</td>"
							Response.write "</tr>"
						next
						%>
					</table>
				</div>
			</div> 
			<%
		case 2 'Comida preparada
			'# Hogares que Reportaron
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " PH_Consumo.Id_Hogar, "
			sql = sql & " PH_PanelHogar.CodigoHogar, "
			sql = sql & " PH_GArea.Area, "
			sql = sql & " ss_Estado.Estado, "
			sql = sql & " PH_Panelistas.Nombre1, "
			sql = sql & " PH_Panelistas.Apellido1, "
			sql = sql & " PH_Panelistas.Celular, "
			sql = sql & " PH_Medio.Medio, "
			sql = sql & " PH_Medio_Delivery.MedioDelivery, "
			sql = sql & " PH_Moneda.Moneda, "
			sql = sql & " PH_Consumo.Tiene_Factura, "
			sql = sql & " PH_Consumo.Nombre_local, "
			sql = sql & " PH_TipoComida.Comida, "
			sql = sql & " PH_Consumo.Total_Compra "
			sql = sql & " FROM ((((((((((PH_Consumo INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) INNER JOIN PH_Medio ON PH_Consumo.Id_Medio = PH_Medio.Id_Medio) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal) LEFT JOIN PH_Cadena ON PH_Consumo.Id_Cadena = PH_Cadena.Id_Cadena) INNER JOIN PH_Moneda ON PH_Consumo.Id_Moneda = PH_Moneda.Id_Moneda) LEFT JOIN PH_TipoComida ON PH_Consumo.Id_TipoComida = PH_TipoComida.Id_TipoComida) LEFT JOIN PH_Medio_Delivery ON PH_Consumo.Id_MedioDetalle = PH_Medio_Delivery.Id_Medio_Delivery "
			sql = sql & " WHERE "
			sql = sql & " PH_Consumo.Id_Hogar > 1 "
			sql = sql & " AND PH_Consumo.Id_Semana = " & idSemana 
			sql = sql & " AND PH_PanelHogar.Ind_activo = 1 "
			sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
			sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 "
			if idArea <> 0 then
				sql = sql & " AND PH_GAreaEstado.Id_Area = " & idArea
			end if
			if idEstado <> 0 then
				sql = sql & " AND  PH_PanelHogar.Id_Estado = " & idEstado
			end if
			sql = sql & " ORDER BY PH_Consumo.Id_Hogar "
			'response.write "<br> Linea 49 "  & sql
			'response.end
			rsx1.Open sql ,conexion
			iExiste = 0
			if rsx1.eof then
				iExiste = 0
				rsx1.close
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
								<th>Medio</th>
								<th>Delivery</th>
								<th>Moneda</th>
								<th>Factura?</th>
								<th>Local</th>
								<th>Comida</th>
								<th>Total Compra</th>
							</tr>
						</thead>
						<%
						for iReg = 0 to ubound(gDatosSol1,2)
							Response.write "<tr>"
								for iCol = 0 to 9
									Response.write "<td>" &  gDatosSol1(iCol,iReg) & "</td>"
								next
								if gDatosSol1(10,iReg) = -1 then 
									Response.write "<td>Si</td>"
								else
									Response.write "<td>No</td>"
								end if
								if gDatosSol1(11,iReg) = "NoAplica" then 
									Response.write "<td></td>"
									Response.write "<td></td>"
									Response.write "<td></td>"
								else
									for iCol = 11 to 13
										Response.write "<td>" &  gDatosSol1(iCol,iReg) & "</td>"
									next
								end if
								'TotalCompras = gDatosSol1(11,iReg)
								'Response.write "<td>" & formatnumber(TotalCompras) & "</td>"
								
							Response.write "</tr>"
						next
						%>
					</table>
				</div>
			</div> 
			<%
		
		case 3,4,5,6
			'# Hogares que Reportaron
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " PH_Consumo.Id_Hogar, "
			sql = sql & " PH_PanelHogar.CodigoHogar, "
			sql = sql & " PH_GArea.Area, "
			sql = sql & " ss_Estado.Estado, "
			sql = sql & " PH_Panelistas.Nombre1, "
			sql = sql & " PH_Panelistas.Apellido1, "
			sql = sql & " PH_Panelistas.Celular, "
			sql = sql & " PH_Medio.Medio, "
			sql = sql & " PH_Moneda.Moneda, "
			sql = sql & " PH_Consumo.Tiene_Factura, "
			sql = sql & " PH_Consumo.Total_Compra "
			sql = sql & " FROM ((((((((PH_Consumo INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) INNER JOIN PH_Medio ON PH_Consumo.Id_Medio = PH_Medio.Id_Medio) INNER JOIN PH_Canal ON PH_Consumo.Id_Canal = PH_Canal.Id_Canal) LEFT JOIN PH_Cadena ON PH_Consumo.Id_Cadena = PH_Cadena.Id_Cadena) INNER JOIN PH_Moneda ON PH_Consumo.Id_Moneda = PH_Moneda.Id_Moneda "
			sql = sql & " WHERE "
			sql = sql & " PH_Consumo.Id_Hogar > 1 "
			sql = sql & " AND PH_Consumo.Id_Semana = " & idSemana 
			sql = sql & " AND PH_PanelHogar.Ind_activo = 1 "
			sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
			sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 "
			if idArea <> 0 then
				sql = sql & " AND PH_GAreaEstado.Id_Area = " & idArea
			end if
			if idEstado <> 0 then
				sql = sql & " AND  PH_PanelHogar.Id_Estado = " & idEstado
			end if
			sql = sql & " ORDER BY "
			sql = sql & " PH_Consumo.Id_Hogar "
			'response.write "<br> Linea 49 "  & sql
			'response.end
			rsx1.Open sql ,conexion
			iExiste = 0
			if rsx1.eof then
				iExiste = 0
				rsx1.close
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
								<th>Medio</th>
								<th>Moneda</th>
								<th>Factura?</th>
								<th>Total Compra</th>
							</tr>
						</thead>
						<%
						for iReg = 0 to ubound(gDatosSol1,2)
							Response.write "<tr>"
								for iCol = 0 to 8
									Response.write "<td>" &  gDatosSol1(iCol,iReg) & "</td>"
								next
								if gDatosSol1(9,iReg) = -1 then 
									Response.write "<td>Si</td>"
								else
									Response.write "<td>No</td>"
								end if
								Response.write "<td>" &  gDatosSol1(10,iReg) & "</td>"
								'TotalCompras = gDatosSol1(11,iReg)
								'Response.write "<td>" & formatnumber(TotalCompras) & "</td>"
								
							Response.write "</tr>"
						next
						%>
					</table>
				</div>
			</div> 
			<%
	
	end select

	'response.write "<br> Linea Final " 
	'response.end
	
%>
	
	
	
