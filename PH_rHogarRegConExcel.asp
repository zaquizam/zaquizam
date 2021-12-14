<%@language=vbscript%>

<!--#include file="Conexion.asp"-->
	
<%

	'==========================================================================================
	' Variables y Constantes
	'==========================================================================================

	Server.ScriptTimeout = 10000
	Response.buffer = true
	
	Dim idCantidadConsumos
	Dim idMesConsulta
	Dim idTipoConsumo
	Dim SemanasConsultas
	Dim gHogares
	Dim idArea
	Dim idEstado
	Dim gDatosSol1
	Dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 1 'adLockOptimistic 

	Dim gDatosSol2
	Dim rsx2
	set rsx2 = CreateObject("ADODB.Recordset")
	rsx2.CursorType = adOpenKeyset 
	rsx2.LockType = 1 'adLockOptimistic 
	
	idMesConsulta = Request.QueryString("mes") 
	idTipoConsumo = Request.QueryString("tip") 

sub BuscarSemanas
	sql = vbnullstring
	sql = sql & " SELECT IdSemana, Semana FROM ss_Semana WHERE Id_Periodo = " & idMesConsulta
	'Response.Write "<br>232 sql:= " & sql
	'response.end
	rsx2.Open sql ,conexion
	if rsx2.eof then
		rsx2.close
	else 
		SemanasConsultas = rsx2.GetRows
		rsx2.close
	end if
end sub	

sub BuscarHogares
	'if idArea = "" then exit sub
	sql = vbnullstring
	sql = sql & " SELECT PH_PanelHogar.Id_PanelHogar, PH_PanelHogar.CodigoHogar, PH_GArea.Area, ss_Estado.Estado, PH_Panelistas.Nombre1, PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Titular, PH_Panelistas.CedulaTitular, PH_Banco.Banco AS Expr1, PH_Banco.Codigo AS Expr2, PH_Panelistas.NumeroCuenta "
	sql = sql & " FROM (((((PH_PanelHogar INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN PH_GAreaEstado ON ss_Estado.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar) LEFT JOIN PH_Nacionalidad ON PH_Panelistas.Id_Nacionalidad = PH_Nacionalidad.Id_Nacionalidad) LEFT JOIN PH_Banco ON PH_Panelistas.Id_Banco = PH_Banco.Id_Banco "
	sql = sql & " WHERE PH_PanelHogar.Ind_Activo = 1 AND PH_Panelistas.ResponsablePanel = 1 "
	'sql = sql & " and PH_PanelHogar.Id_PanelHogar = 706 "
	sql = sql & " Order by PH_GArea.Area, ss_Estado.Estado "
	'Response.Write "<br>232 sql:= " & sql
	'response.end
	rsx2.Open sql ,conexion
	if rsx2.eof then
		rsx2.close
	else 
		gHogares = rsx2.GetRows
		rsx2.close
	end if
	'Response.Write "<br>241 LLEGO <br>"
	'response.end

end sub
   	
	'Response.Write "<br> Combo1:=" & idMesConsulta
	'Response.Write "<br> Combo2:=" & idTipoConsumo
	'Response.Write "llego1"
	'response.end
	'hidden 
	'Response.Write "<br>241 LLEGO <br>"
	'
	if CInt(idMesConsulta) <> 0 and CInt(idTipoConsumo) <> 0 then
		
		BuscarSemanas
		Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
		Response.ContentType = "application/vnd.ms-excel"
		BuscarHogares
%>
		<table>
			<tr>
				<td>idHogar</td>
				<td>Hogar</td>
				<td>Area</td>
				<td>Estado</td>
				<td>Nombre Panelista</td>
				<td>Apellido Panelista</td>
				<%
					for iSem = 0 to ubound(SemanasConsultas,2)
						Response.Write "<td>" & SemanasConsultas(1,iSem)  & "</td>"
					next
				%>
			</tr>
			<%
			
			for iHog = 0 to ubound(gHogares,2)
				idHogar = gHogares(0,iHog)
				Response.Flush
				Response.Write "<tr>"
					Response.Write "<td>" & gHogares(0,iHog) & "</td>"
					Response.Write "<td>" & gHogares(1,iHog) & "</td>"
					Response.Write "<td>" & gHogares(2,iHog) & "</td>"
					Response.Write "<td>" & gHogares(3,iHog) & "</td>"
					Response.Write "<td>" & gHogares(4,iHog) & "</td>"
					Response.Write "<td>" & gHogares(5,iHog) & "</td>"
					for iSem = 0 to ubound(SemanasConsultas,2)
						iSemana = SemanasConsultas(0,iSem)
						sql = vbnullstring
						sql = sql & " SELECT Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo FROM PH_Consumo "
						sql = sql & " WHERE PH_Consumo.Id_Semana = " & iSemana 
						sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
						sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
						'Response.Write "<br>232 sql:= " & sql
						'response.end
						rsx2.Open sql ,conexion
						if rsx2.eof then
							rsx2.close
							iSemanaT = 0
						else 
							gDatosSol1 = rsx2.GetRows
							rsx2.close
							iSemanaT = gDatosSol1(0,0)  
						end if
						sql = vbnullstring
						sql = sql & " SELECT "
						sql = sql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS CuentaDeId_Consumo_Detalle_Productos "
						sql = sql & " FROM PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) "
						sql = sql & " WHERE PH_Consumo.Id_Semana = " & iSemana
						sql = sql & " GROUP BY PH_Consumo.Id_Hogar, PH_Consumo.id_TipoConsumo "
						sql = sql & " HAVING PH_Consumo.Id_Hogar = " & idHogar 
						sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
						'Response.Write "<br>232 sql:= " & sql
						'response.end
						rsx2.Open sql ,conexion
						if rsx2.eof then
							rsx2.close
							iSemanaTReg = 0
						else 
							gDatosSol1 = rsx2.GetRows
							rsx2.close
							iSemanaTReg = gDatosSol1(0,0)  
						end if
						Response.Write "<td>'" & iSemanaT & "-" & iSemanaTReg & "</td>"
					
					next
					Response.Flush
				Response.Write "<tr>"
			next 
			Response.Flush
		%>
		</table>
	<%
	end if
	conexion.close
	%>
</body>
</html>