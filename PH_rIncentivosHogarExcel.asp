<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%

  
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	dim idCantidadConsumos
	dim idMesPago
	dim idSemanasPago
	dim gHogares
	dim idArea
	dim idEstado
	
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

sub BuscarHogares
	'if idArea = "" then exit sub
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " PH_GArea.Area, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Titular, "
	sql = sql & " PH_Panelistas.CedulaTitular, "
	sql = sql & " PH_Banco.Banco AS Expr1, "
	sql = sql & " PH_Banco.Codigo AS Expr2, "
	sql = sql & " PH_Panelistas.NumeroCuenta "
	sql = sql & " FROM (((((PH_PanelHogar INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_GAreaEstado ON ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) LEFT JOIN PH_Nacionalidad ON PH_Panelistas.Id_Nacionalidad = PH_Nacionalidad.Id_Nacionalidad) LEFT JOIN PH_Banco ON PH_Panelistas.Id_Banco = PH_Banco.Id_Banco "
	sql = sql & " WHERE "
	sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
	sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 "
	sql = sql & " and PH_GArea.id_Area = "  & idArea
	sql = sql & " and ss_Estado.id_Estado = " & idEstado
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 706 "
	sql = sql & " Order by "
	sql = sql & " PH_GArea.Area, "
	sql = sql & " ss_Estado.Estado "
	'response.write "<br>232 sql:= " & sql
	'response.end
	rsx2.Open sql ,conexion
	if rsx2.eof then 
		rsx2.close
	else 
		gHogares = rsx2.GetRows
		rsx2.close
	end if
		'response.write "<br>241 LLEGO <br>"
		'response.end

end sub
   
  
	idMesPago = Request.QueryString("mes") 
	idCantidadConsumos = Request.QueryString("can")  
	idArea = Request.QueryString("are")  
	idEstado = Request.QueryString("est")  

    
%>
		
	<%
	
	%>
	<%
	if idMesPago = "" then idMesPago = 0
	if idCantidadConsumos = "" then idCantidadConsumos = 0
	if idArea = "" then idArea = 0 
	Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
	Response.ContentType = "application/vnd.ms-excel"
	

	'response.write "<br>LLEGO"
	'response.end
	if cint(idMesPago) <> 0 and cint(idCantidadConsumos) <> 0 and cint(idArea) <> 0 then
		'BuscarSemanas
		BuscarHogares
		'response.write "<br>140 Semanas:= " & idSemanasPago
		%>
		<table>
			<thead>
				<tr class="w3-blue">
					<th>idHogar</th>
					<th>Hogar</th>
					<th>Area</th>
					<th>Estado</th>
					<th>Nombre Panelista</th>
					<th>Apellido Panelista</th>
					<th>Nombre y Apellido Titular</th>
					<th>Cedula</th>
					<th>Banco</th>
					<th>Banco Codigo</th>
					<th>Cuenta</th>
					<th>(09) Del 01 Mar 2021 al 07 Mar 2021</th>
					<th>(10) Del 08 Mar 2021 al 14 Mar 2021</th>
					<th>(11) Del 15 Mar 2021 al 21 Mar 2021</th>
					<th>(12) Del 22 Mar 2021 al 28 Mar 2021</th>
					<th>(13) Del 29 Mar 2021 al 04 Abr 2021</th>
					<th>Prueba de producto - Limpieza de manos</th>
					<th>Cuidado Personal</th>
					<th>Pagar Incentivo</th>
					<th>Pagar Encuesta1</th>
					<th>Pagar Encuesta2</th>
				</tr>
			</thead>
			<%
			for iHog = 0 to ubound(gHogares,2)
				idHogar = gHogares(0,iHog)
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
				sql = sql & " FROM PH_Consumo "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana = 24 "
				sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
				sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
				'response.write "<br>232 sql:= " & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
					rsx2.close
					iSemana1 = 0
				else 
					gDatosSol1 = rsx2.GetRows
					rsx2.close
					iSemana1 = gDatosSol1(0,0)  
				end if
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS CuentaDeId_Consumo_Detalle_Productos "
				sql = sql & " FROM PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana = 24 "
				sql = sql & " GROUP BY "
				sql = sql & " PH_Consumo.Id_Hogar, "
				sql = sql & " PH_Consumo.id_TipoConsumo "
				sql = sql & " HAVING "
				sql = sql & " PH_Consumo.Id_Hogar = " & idHogar 
				sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
				'response.write "<br>232 sql:= " & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
					rsx2.close
					iSemana1Reg = 0
				else 
					gDatosSol1 = rsx2.GetRows
					rsx2.close
					iSemana1Reg = gDatosSol1(0,0)  
				end if

				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
				sql = sql & " FROM PH_Consumo "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana = 25 "
				sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
				sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
				'response.write "<br>232 sql:= " & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
					rsx2.close
					iSemana2 = 0
				else 
					gDatosSol1 = rsx2.GetRows
					rsx2.close
					iSemana2 = gDatosSol1(0,0)  
				end if
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS CuentaDeId_Consumo_Detalle_Productos "
				sql = sql & " FROM PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana = 25 "
				sql = sql & " GROUP BY "
				sql = sql & " PH_Consumo.Id_Hogar, "
				sql = sql & " PH_Consumo.id_TipoConsumo "
				sql = sql & " HAVING "
				sql = sql & " PH_Consumo.Id_Hogar = " & idHogar 
				sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
				'response.write "<br>232 sql:= " & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
					rsx2.close
					iSemana2Reg = 0
				else 
					gDatosSol1 = rsx2.GetRows
					rsx2.close
					iSemana2Reg = gDatosSol1(0,0)  
				end if
				

				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
				sql = sql & " FROM PH_Consumo "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana = 26 "
				sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
				sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
				'response.write "<br>232 sql:= " & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
					rsx2.close
					iSemana3 = 0
				else 
					gDatosSol1 = rsx2.GetRows
					rsx2.close
					iSemana3 = gDatosSol1(0,0)  
				end if
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS CuentaDeId_Consumo_Detalle_Productos "
				sql = sql & " FROM PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana = 26 "
				sql = sql & " GROUP BY "
				sql = sql & " PH_Consumo.Id_Hogar, "
				sql = sql & " PH_Consumo.id_TipoConsumo "
				sql = sql & " HAVING "
				sql = sql & " PH_Consumo.Id_Hogar = " & idHogar 
				sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
				'response.write "<br>232 sql:= " & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
					rsx2.close
					iSemana3Reg = 0
				else 
					gDatosSol1 = rsx2.GetRows
					rsx2.close
					iSemana3Reg = gDatosSol1(0,0)  
				end if
				
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
				sql = sql & " FROM PH_Consumo "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana = 27 "
				sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
				sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
				'response.write "<br>232 sql:= " & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
					rsx2.close
					iSemana4 = 0
				else 
					gDatosSol1 = rsx2.GetRows
					rsx2.close
					iSemana4 = gDatosSol1(0,0)  
				end if
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS CuentaDeId_Consumo_Detalle_Productos "
				sql = sql & " FROM PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana = 27 "
				sql = sql & " GROUP BY "
				sql = sql & " PH_Consumo.Id_Hogar, "
				sql = sql & " PH_Consumo.id_TipoConsumo "
				sql = sql & " HAVING "
				sql = sql & " PH_Consumo.Id_Hogar = " & idHogar 
				sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
				'response.write "<br>232 sql:= " & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
					rsx2.close
					iSemana4Reg = 0
				else 
					gDatosSol1 = rsx2.GetRows
					rsx2.close
					iSemana4Reg = gDatosSol1(0,0)  
				end if

				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
				sql = sql & " FROM PH_Consumo "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana = 28 "
				sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
				sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
				'response.write "<br>232 sql:= " & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
					rsx2.close
					iSemana5 = 0
				else 
					gDatosSol1 = rsx2.GetRows
					rsx2.close
					iSemana5 = gDatosSol1(0,0)  
				end if
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS CuentaDeId_Consumo_Detalle_Productos "
				sql = sql & " FROM PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana = 28 "
				sql = sql & " GROUP BY "
				sql = sql & " PH_Consumo.Id_Hogar, "
				sql = sql & " PH_Consumo.id_TipoConsumo "
				sql = sql & " HAVING "
				sql = sql & " PH_Consumo.Id_Hogar = " & idHogar 
				sql = sql & " AND PH_Consumo.id_TipoConsumo = 1 "
				'response.write "<br>232 sql:= " & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
					rsx2.close
					iSemana5Reg = 0
				else 
					gDatosSol1 = rsx2.GetRows
					rsx2.close
					iSemana5Reg = gDatosSol1(0,0)  
				end if

				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_EncuestaEspecial, "
				sql = sql & " Id_Hogar, "
				sql = sql & " Ind_Realizada "
				sql = sql & " FROM PH_EncuestaHogar "
				sql = sql & " WHERE "
				sql = sql & " Id_EncuestaEspecial = 29 "
				sql = sql & " and Id_Hogar = " & idHogar
				sql = sql & " AND Ind_Realizada =  1 "
				'response.write "<br>232 sql:= " & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
					rsx2.close
					iEncuesta1 = 0
				else 
					gDatosSol1 = rsx2.GetRows
					rsx2.close
					iEncuesta1 = 1
				end if

				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Id_EncuestaEspecial, "
				sql = sql & " Id_Hogar, "
				sql = sql & " Ind_Realizada "
				sql = sql & " FROM PH_EncuestaHogar "
				sql = sql & " WHERE "
				sql = sql & " Id_EncuestaEspecial = 33 "
				sql = sql & " and Id_Hogar = " & idHogar
				sql = sql & " AND Ind_Realizada =  1 "
				'response.write "<br>232 sql:= " & sql
				'response.end
				rsx2.Open sql ,conexion
				if rsx2.eof then
					rsx2.close
					iEncuesta2 = 0
				else 
					gDatosSol1 = rsx2.GetRows
					rsx2.close
					iEncuesta2 = 1
				end if

				iSemana = 0
				if iSemana1 > 0 then iSemana = iSemana + 1
				if iSemana2 > 0 then iSemana = iSemana + 1
				if iSemana3 > 0 then iSemana = iSemana + 1
				if iSemana4 > 0 then iSemana = iSemana + 1
				if iSemana5 > 0 then iSemana = iSemana + 1
				
				if cint(iSemana) >= cint(idCantidadConsumos)  or iEncuesta1 = 1 or  iEncuesta2 = 1 then
					response.write "<tr>"
						Response.flush
						response.write "<td>" & gHogares(0,iHog) & "</td>"
						response.write "<td>" & gHogares(1,iHog) & "</td>"
						response.write "<td>" & gHogares(2,iHog) & "</td>"
						response.write "<td>" & gHogares(3,iHog) & "</td>"
						response.write "<td>" & gHogares(4,iHog) & "</td>"
						response.write "<td>" & gHogares(5,iHog) & "</td>"
						response.write "<td>" & gHogares(6,iHog) & "</td>"
						iLen = len(gHogares(7,iHog))
						Cedula = ""
						if iLen = 7 then 
							Cedula = "0" & gHogares(7,iHog)
						else
							if iLen = 6 then 
								Cedula = "00" & gHogares(7,iHog)
							else
								Cedula = gHogares(7,iHog)
							end if
						end if
						response.write "<td>V" & Cedula & "</td>"
						response.write "<td>" & gHogares(8,iHog) & "</td>"
						response.write "<td>" & gHogares(9,iHog) & "</td>"
						response.write "<td>'" & gHogares(10,iHog) & "</td>"
						
						response.write "<td>'" & iSemana1 & "-" & iSemana1Reg & "</td>"
						response.write "<td>'" & iSemana2 & "-" & iSemana2Reg &"</td>"
						response.write "<td>'" & iSemana3 & "-" & iSemana3Reg &"</td>"
						response.write "<td>'" & iSemana4 & "-" & iSemana4Reg &"</td>"
						response.write "<td>'" & iSemana5 & "-" & iSemana5Reg &"</td>"
						response.write "<td>" & iEncuesta1 & "</td>"
						response.write "<td>" & iEncuesta2 & "</td>"
						
						if cint(iSemana) >= cint(idCantidadConsumos) then
							response.write "<td>Si</td>"
						else
							response.write "<td>No</td>"
						end if
						
						if iEncuesta1 = 1 then
							response.write "<td>Si</td>"
						else
							response.write "<td>No</td>"
						end if
						if iEncuesta2 = 1 then
							response.write "<td>Si</td>"
						else
							response.write "<td>No</td>"
						end if
					response.write "</tr>"
					
				end if
			next 
			
			%>
		</table>
		<%
	end if


	conexion.close
	%>

</body>
</html>