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
	sql = sql & " PH_Nacionalidad.Abreviatura, "
	sql = sql & " PH_Panelistas.Cedula, "
	sql = sql & " PH_Panelistas.NumeroCuenta "
	sql = sql & " FROM ((((PH_PanelHogar INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_GAreaEstado ON ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) LEFT JOIN PH_Nacionalidad ON PH_Panelistas.Id_Nacionalidad = PH_Nacionalidad.Id_Nacionalidad "
	sql = sql & " WHERE "
	sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
	sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 "
	sql = sql & " and PH_GArea.id_Area = "  & idArea
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 123 "
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

end sub
   
  
	idMesPago = Request.QueryString("mes") 
	idCantidadConsumos = Request.QueryString("can")  
	idArea = Request.QueryString("are")  

    
%>
		
	<%
	
	%>
	<%
	if idMesPago = "" then idMesPago = 0
	if idCantidadConsumos = "" then idCantidadConsumos = 0
	if idArea = "" then idArea = 0 
	'Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
	'Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-disposition","attachment; filename=tem.txt"
	Response.ContentType = "text/plain"
	

	'response.write "<br>LLEGO"
	'response.end
	if cint(idMesPago) <> 0 and cint(idCantidadConsumos) <> 0 and cint(idArea) <> 0 then
		'BuscarSemanas
		BuscarHogares
		
		'response.write "<br>140 Semanas:= " & idSemanasPago
		
		
		%>
					<%
					for iHog = 0 to ubound(gHogares,2)
						Response.flush
						response.write chr(10)
							idHogar = gHogares(0,iHog)
							sql = ""
							sql = sql & " SELECT "
							sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
							sql = sql & " FROM PH_Consumo "
							sql = sql & " WHERE "
							sql = sql & " PH_Consumo.Id_Semana = 20 "
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
							sql = sql & " PH_Consumo.Id_Semana = 20 "
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
							sql = sql & " PH_Consumo.Id_Semana = 21 "
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
							sql = sql & " PH_Consumo.Id_Semana = 21 "
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
							sql = sql & " PH_Consumo.Id_Semana = 22 "
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
							sql = sql & " PH_Consumo.Id_Semana = 22 "
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
							sql = sql & " PH_Consumo.Id_Semana = 23 "
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
							sql = sql & " PH_Consumo.Id_Semana = 23 "
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
							sql = sql & " Id_EncuestaEspecial, "
							sql = sql & " Id_Hogar, "
							sql = sql & " Ind_Realizada "
							sql = sql & " FROM PH_EncuestaHogar "
							sql = sql & " WHERE "
							sql = sql & " Id_EncuestaEspecial = 27 "
							sql = sql & " and Id_Hogar = " & idHogar
							sql = sql & " AND Ind_Realizada =  1 "
							'response.write "<br>232 sql:= " & sql
							'response.end
							rsx2.Open sql ,conexion
							if rsx2.eof then
								rsx2.close
								iEncuesta = 0
							else 
								gDatosSol1 = rsx2.GetRows
								rsx2.close
								iEncuesta = 1
							end if

							iSemana = 0
							if iSemana1 > 0 then iSemana = iSemana + 1
							if iSemana2 > 0 then iSemana = iSemana + 1
							if iSemana3 > 0 then iSemana = iSemana + 1
							if iSemana4 > 0 then iSemana = iSemana + 1
							
							if cint(iSemana) >= cint(idCantidadConsumos)  or iEncuesta = 1 then

								response.write gHogares(0,iHog) & chr(9)
								response.write gHogares(1,iHog) & chr(9)
								response.write gHogares(2,iHog) & chr(9)
								response.write gHogares(3,iHog) & chr(9)
								response.write gHogares(4,iHog) & " " & gHogares(5,iHog) & chr(9)
								iLen = len(gHogares(7,iHog))
								Cedula = gHogares(6,iHog)
								if iLen = 7 then 
									Cedula = Cedula & "0" & gHogares(7,iHog)
								else
									Cedula = Cedula & gHogares(7,iHog)
								end if
								response.write Cedula & chr(9)
								response.write gHogares(8,iHog) & chr(9)
								
								
								
								response.write "cons/reg:" & iSemana1 & "-" & iSemana1Reg & chr(9)
								response.write "cons/reg:" & iSemana2 & "-" & iSemana2Reg & chr(9)
								response.write "cons/reg:" & iSemana3 & "-" & iSemana3Reg & chr(9)
								response.write "cons/reg:" & iSemana4 & "-" & iSemana4Reg & chr(9)
								response.write iEncuesta & chr(9)
								
								if cint(iSemana) >= cint(idCantidadConsumos) then
									response.write "Si" & chr(9)
								else
									response.write "No" & chr(9)
								end if
								
								if iEncuesta = 1 then
									response.write "Si" 
								else
									response.write "No" 
								end if
								
								
							end if
								
						response.write chr(10)
					next 
					
					%>
				</table>
		<%
	end if


	conexion.close
	%>

</body>
</html>