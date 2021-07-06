<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%

	Server.ScriptTimeout=1000
	Response.buffer = true
  
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	dim gDatosSol1
	dim rsx1
	dim gHogares
	
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	dim gDatosSol2
	dim rsx2
	set rsx2 = CreateObject("ADODB.Recordset")
	rsx2.CursorType = adOpenKeyset 
	rsx2.LockType = 2 'adLockOptimistic 

	idMesPago = Request.QueryString("mes") 
	idCantidadConsumos = Request.QueryString("can")  
	idArea = Request.QueryString("are")  
	idEstado = Request.QueryString("est")  
	

Sub Combos
 
	'response.write "<br>372 Combo1:=" & ed_sPar(1,0)
	'response.write " Combo2:=" & ed_sPar(2,0)
	'response.write " Combo3:=" & ed_sPar(3,0)
	'response.write " Combo4:=" & ed_sPar(4,0)
	'response.write " Combo3:=" & ed_sPar(5,0)
	'response.end 
    
	ed_iCombo = 2
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id, "
	sql = sql & " Periodo "
	sql = sql & " FROM "
	sql = sql & " ss_Periodo "
	sql = sql & " WHERE "
	sql = sql & " IdAno = 2021 "
	sql = sql & " AND IdMes = 6 "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Mes"
    ed_sCombo(1,1)=sql 
    ed_sCombo(1,2)="Seleccionar"

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Cantidad, "
	sql = sql & " Cantidad "
	sql = sql & " FROM "
	sql = sql & " ss_Cantidad "
	sql = sql & " WHERE "
	sql = sql & " Id_Cantidad < 5 "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(2,0)="Cantidad Semanas"
    ed_sCombo(2,1)=sql 
    ed_sCombo(2,2)="Seleccionar"

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Area, "
	sql = sql & " Area "
	sql = sql & " FROM PH_GArea "
	sql = sql & " Order By "
	sql = sql & " Id_Area "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(3,0)="Area"
    ed_sCombo(3,1)=sql 
    ed_sCombo(3,2)="Seleccionar"

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_GAreaEstado.Id_Estado, "
	sql = sql & " ss_Estado.Estado "
	sql = sql & " FROM PH_GAreaEstado INNER JOIN ss_Estado ON PH_GAreaEstado.Id_Estado = ss_Estado.Id_Estado "
	if ed_sPar(3,0) <> "" and ed_sPar(3,0) <> "Seleccionar" then
		sql = sql & " WHERE PH_GAreaEstado.Id_Area = " & ed_sPar(3,0)
	end if
	sql = sql & " Order By "
	sql = sql & " ss_Estado.Estado "
	'response.write "<br>372 Combo2:=" & sql
    ed_sCombo(4,0)="Estado"
    ed_sCombo(4,1)=sql 
    ed_sCombo(4,2)="Seleccionar"

	
End Sub

sub BuscarSemanas
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Semanas "
	sql = sql & " FROM "
	sql = sql & " ss_Periodo "
	sql = sql & " WHERE "
	sql = sql & " Id = " & idMesPago
	'response.write "<br>232 sql:= " & sql
	'response.end
	rsx2.Open sql ,conexion
	if rsx2.eof then
		rsx2.close
	else 
		gDatosSol2 = rsx2.GetRows
		rsx2.close
		idSemanasPago = gDatosSol2(0,0)
	end if
end sub	

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
   
    'LeePar
	'Combos
	
  
	' if ed_sPar(1,0) = "Seleccionar" then
		' idMesPago = 0 
	' else
		' idMesPago = ed_sPar(1,0)
	' end if
	' if ed_sPar(2,0) = "Seleccionar" then
		' idCantidadConsumos = 0 
	' else
		' idCantidadConsumos = ed_sPar(2,0)
	' end if

	' if ed_sPar(3,0) = "Seleccionar" then
		' idArea = 0 
	' else
		' idArea = ed_sPar(3,0)
	' end if

	' if ed_sPar(4,0) = "Seleccionar" or ed_sPar(4,0) = "" then
		' idEstado = 0 
	' else
		' idEstado = ed_sPar(4,0)
	' end if
    

	'sExcel = "mes=" & idMesPago & "&can=" & idCantidadConsumos & "&are=" & idArea & "&est=" & idEstado
    
%>
		
	<%
	'hidden
	
	'response.write "<br> Combo1:=" & ed_sPar(1,0) & "==>" & idMesPago
	'response.write "<br> Combo2:=" & ed_sPar(2,0) & "==>" & idCantidadConsumos
	'response.write "<br> Combo3:=" & ed_sPar(3,0) & "==>" & idArea
	'response.write "llego1"
	if idMesPago = "" then idMesPago = 0
	if idCantidadConsumos = "" then idCantidadConsumos = 0
	if idArea = "" then idArea = 0 
	'response.end
	'hidden 
	Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
	Response.ContentType = "application/vnd.ms-excel"

	'response.end
	'if cint(idMesPago) <> 0 and cint(idCantidadConsumos) <> 0 and cint(idArea) <> 0  and cint(idEstado) <> 0 then
	'if cint(idMesPago) <> 0 and cint(idCantidadConsumos) <> 0 then
		'BuscarSemanas
		BuscarHogares
		'response.write "<br>140 Semanas:= " & idSemanasPago
		'response.end
		%>
		<table>
			<tr>
				<td>idHogar</td>
				<td>Hogar</td>
				<td>Area</td>
				<td>Estado</td>
				<td>Nombre Panelista</td>
				<td>Apellido Panelista</td>
				<td>Nombre y Apellido Titular</td>
				<td>Cedula</td>
				<td>Banco</td>
				<td>Banco Codigo</td>
				<td>Cuenta</td>
				<th>(22) Del 31 May 2021 al 06 Jun 2021</th>
				<th>(23) Del 07 Jun 2021 al 13 Jun 2021</th>
				<th>(24) Del 14 Jun 2021 al 20 Jun 2021</th>
				<th>(25) Del 21 Jun 2021 al 27 Jun 2021</th>
				<th>Pagar Incentivo</th>
				<th>Enc. Enjuaje Bucal</th>
				<th>Enc. Alimento</th>

			</tr>
			<%
			'response.end
			for iHog = 0 to ubound(gHogares,2)
				idHogar = gHogares(0,iHog)
				
				sql = ""
				sql = sql & " SELECT "
				sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
				sql = sql & " FROM PH_Consumo "
				sql = sql & " WHERE "
				sql = sql & " PH_Consumo.Id_Semana = 37 "
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
				sql = sql & " PH_Consumo.Id_Semana = 37 "
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
				sql = sql & " PH_Consumo.Id_Semana = 38 "
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
				sql = sql & " PH_Consumo.Id_Semana = 38 "
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
				sql = sql & " PH_Consumo.Id_Semana = 39 "
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
				sql = sql & " PH_Consumo.Id_Semana = 39 "
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
				sql = sql & " PH_Consumo.Id_Semana = 40 "
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
				sql = sql & " PH_Consumo.Id_Semana = 40 "
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

				iSemana = 0
				if iSemana1 > 0 then iSemana = iSemana + 1
				if iSemana2 > 0 then iSemana = iSemana + 1
				if iSemana3 > 0 then iSemana = iSemana + 1
				if iSemana4 > 0 then iSemana = iSemana + 1
				
				if cint(iSemana) >= cint(idCantidadConsumos) then
					response.write "<tr>"

						Response.flush
						'response.write "<td>(" & iHog & ")" & gHogares(0,iHog) & "</td>"
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
						response.write "<td>'" & gHogares(9,iHog) & "</td>"
						response.write "<td>'" & gHogares(10,iHog) & "</td>"
						
						response.write "<td>'" & iSemana1 & "-" & iSemana1Reg & "</td>"
						response.write "<td>'" & iSemana2 & "-" & iSemana2Reg &"</td>"
						response.write "<td>'" & iSemana3 & "-" & iSemana3Reg &"</td>"
						response.write "<td>'" & iSemana4 & "-" & iSemana4Reg &"</td>"
						
						if cint(iSemana) >= cint(idCantidadConsumos) then
							response.write "<td>Si</td>"
						else
							response.write "<td>No</td>"
						end if
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Id_EncuestaEspecial, "
						sql = sql & " Id_Hogar, "
						sql = sql & " Ind_Realizada "
						sql = sql & " FROM PH_EncuestaHogar "
						sql = sql & " WHERE "
						sql = sql & " Id_EncuestaEspecial = 40 "
						sql = sql & " and Id_Hogar = " & idHogar
						sql = sql & " AND Ind_Realizada =  1 "
						'response.write "<br>232 sql:= " & sql
						'response.end
						rsx2.Open sql ,conexion
						if rsx2.eof then
							rsx2.close
							response.write "<td>No</td>"
						else 
							gDatosSol1 = rsx2.GetRows
							rsx2.close
							response.write "<td>Si</td>"
						end if
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Id_EncuestaEspecial, "
						sql = sql & " Id_Hogar, "
						sql = sql & " Ind_Realizada "
						sql = sql & " FROM PH_EncuestaHogar "
						sql = sql & " WHERE "
						sql = sql & " Id_EncuestaEspecial = 37 "
						sql = sql & " and Id_Hogar = " & idHogar
						sql = sql & " AND Ind_Realizada =  1 "
						'response.write "<br>232 sql:= " & sql
						'response.end
						rsx2.Open sql ,conexion
						if rsx2.eof then
							rsx2.close
							response.write "<td>No</td>"
						else 
							gDatosSol1 = rsx2.GetRows
							rsx2.close
							response.write "<td>Si</td>"
						end if

					response.write "</tr>"
				end if
			next 
			
			%>
		</table>
		<%
	'end if


	conexion.close
	%>

</body>
</html>