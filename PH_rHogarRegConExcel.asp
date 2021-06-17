<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%

  
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	dim idCantidadConsumos
	dim idMesConsulta
	dim idTipoConsumo
	dim SemanasConsultas
	dim gHogares
	dim idArea
	dim idEstado
	Server.ScriptTimeout=1000
	Response.buffer = true
	
%>
	<script>
	function GenerarExcel()
	{
		//alert("Generar Excel");
		num = document.getElementById("Excel").value;
		//alert("Generar Excel:="+ num);
		window.open("ph_rIncentivosHogarExcel.asp?"+num,"_blank");
	}
		

	</script>   
	
<%
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


sub BuscarSemanas
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " IdSemana, "
	sql = sql & " Semana "
	sql = sql & " FROM "
	sql = sql & " ss_Semana "
	sql = sql & " WHERE "
	sql = sql & " Id_Periodo = " & idMesConsulta
	'response.write "<br>232 sql:= " & sql
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
	
	idMesConsulta = Request.QueryString("mes") 
	idTipoConsumo = Request.QueryString("tip") 
    
	
	'response.write "<br> Combo1:=" & idMesConsulta
	'response.write "<br> Combo2:=" & idTipoConsumo
	'response.write "llego1"
	'response.end
	'hidden 
	'response.write "<br>241 LLEGO <br>"
	if cint(idMesConsulta) <> 0 and cint(idTipoConsumo) <> 0 then
		
		BuscarSemanas
		Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
		Response.ContentType = "application/vnd.ms-excel"
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
					response.write "<td>" & SemanasConsultas(1,iSem)  & "</td>"
				next
				%>
			</tr>
			<%
			BuscarHogares
			for iHog = 0 to ubound(gHogares,2)
				idHogar = gHogares(0,iHog)
				Response.flush
				response.write "<tr>"
					response.write "<td>" & gHogares(0,iHog) & "</td>"
					response.write "<td>" & gHogares(1,iHog) & "</td>"
					response.write "<td>" & gHogares(2,iHog) & "</td>"
					response.write "<td>" & gHogares(3,iHog) & "</td>"
					response.write "<td>" & gHogares(4,iHog) & "</td>"
					response.write "<td>" & gHogares(5,iHog) & "</td>"
					for iSem = 0 to ubound(SemanasConsultas,2)
						iSemana = SemanasConsultas(0,iSem)
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Count(PH_Consumo.Id_Consumo) AS CuentaDeId_Consumo "
						sql = sql & " FROM PH_Consumo "
						sql = sql & " WHERE "
						sql = sql & " PH_Consumo.Id_Semana = " & iSemana 
						sql = sql & " AND PH_Consumo.Id_Hogar = " & idHogar
						sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
						'response.write "<br>232 sql:= " & sql
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
						sql = ""
						sql = sql & " SELECT "
						sql = sql & " Count(PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos) AS CuentaDeId_Consumo_Detalle_Productos "
						sql = sql & " FROM PH_Consumo INNER JOIN PH_Consumo_Detalle_Productos ON (PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar) AND (PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo) "
						sql = sql & " WHERE "
						sql = sql & " PH_Consumo.Id_Semana = " & iSemana
						sql = sql & " GROUP BY "
						sql = sql & " PH_Consumo.Id_Hogar, "
						sql = sql & " PH_Consumo.id_TipoConsumo "
						sql = sql & " HAVING "
						sql = sql & " PH_Consumo.Id_Hogar = " & idHogar 
						sql = sql & " AND PH_Consumo.id_TipoConsumo = " & idTipoConsumo
						'response.write "<br>232 sql:= " & sql
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
						response.write "<td>'" & iSemanaT & "-" & iSemanaTReg & "</td>"
					
					next
				response.write "<tr>"
			next 
		%>
		</table>
	<%
	end if
	conexion.close
	%>
</body>
</html>