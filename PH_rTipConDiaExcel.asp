<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
  
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	idSemana = Request.QueryString("sem")
	idEstado = Request.QueryString("est")
	idArea = Request.QueryString("are")

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

	
%>
	<script>

	</script>   
<%

   
'==========================================================================================
' Parámetros del Manteniemiento
'==========================================================================================


%>
		
	<br>
	<div style="width:98%">
	<%
	
	%></div></center>

	<%
	'response.write "<br> Combo1:=" & ed_sPar(1,0) & "==>" & idArea
	'response.write "<br> Combo2:=" & ed_sPar(2,0) & "==>" & idEstado
	'response.write "<br> Combo3:=" & ed_sPar(3,0) & "==>" & idSemana
	Response.AddHeader "Content-disposition","attachment; filename=tem.xls"
	Response.ContentType = "application/vnd.ms-excel"
	'hidden 
	%>
	<table>
		<tr>
			<td>Dia</td>
			<td>Area</td>
			<td>Estado</td>
			<td>Tipo de Consumo</td>
			<td># Hogares que Reportaron</td>
		</tr>
		<%
		'Response.write "<br>176 idSemana:" & idSemana
		'Response.write "<br>176 idArea:" & idArea
		'Response.write "<br>176 idEstado:" & idEstado
		if idArea <>0 Then
			sql = ""
			sql = sql & " SELECT "
			sql = sql & " PH_Consumo.Fecha_Creacion, "
			sql = sql & " PH_GAreaEstado.Id_Area, "
			sql = sql & " PH_GArea.Area, "
			sql = sql & " PH_PanelHogar.Id_Estado, "
			sql = sql & " ss_Estado.Estado, "
			sql = sql & " PH_Consumo.id_TipoConsumo, "
			sql = sql & " PH_TipoConsumo.TipoConsumo "
			sql = sql & " FROM ((((PH_Consumo INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_TipoConsumo ON PH_Consumo.id_TipoConsumo = PH_TipoConsumo.Id_TipoConsumo) INNER JOIN ss_Estado ON PH_GAreaEstado.Id_Estado = ss_Estado.Id_Estado "
			sql = sql & " WHERE "
			sql = sql & " PH_Consumo.Id_Semana = " & idSemana
			sql = sql & " AND PH_Consumo.Status_registro='G' "
			sql = sql & " AND PH_Consumo.Id_Hogar > 1 "
			sql = sql & " GROUP BY "
			sql = sql & " PH_Consumo.Fecha_Creacion, "
			sql = sql & " PH_GAreaEstado.Id_Area, "
			sql = sql & " PH_GArea.Area, "
			sql = sql & " PH_PanelHogar.Id_Estado, "
			sql = sql & " ss_Estado.Estado, "
			sql = sql & " PH_Consumo.id_TipoConsumo, "
			sql = sql & " PH_TipoConsumo.TipoConsumo "
			sql = sql & " HAVING "
			sql = sql & " PH_GAreaEstado.Id_Area = " & idArea
			if idEstado <> 0 then 
				sql = sql & " AND PH_PanelHogar.Id_Estado = " & idEstado
			end if
			sql = sql & " ORDER BY "
			sql = sql & " PH_Consumo.Fecha_Creacion DESC , "
			sql = sql & " PH_GArea.Area, "
			sql = sql & " ss_Estado.Estado "
			'response.write "<br>232 sql:= " & sql
			'response.end
			rsx1.Open sql ,conexion
			if rsx1.eof then
				rsx1.close
			else 
				gDatosSol1 = rsx1.GetRows
				rsx1.close
			end if
			for iReg = 0 to ubound(gDatosSol1,2)
				Response.flush
				response.write "<tr>"
					sFecha = gDatosSol1(0,iReg)
					Dia = mid(sFecha,9,2)
					Mes = mid(sFecha,6,2)							
					Ano = mid(sFecha,1,4)
					nFecha = Dia & "/" & Mes & "/" & Ano
					response.write "<td>" & nFecha & "</td>"
					response.write "<td>" & gDatosSol1(2,iReg) & "</td>"
					response.write "<td>" & gDatosSol1(4,iReg) & "</td>"
					response.write "<td>" & gDatosSol1(6,iReg) & "</td>"
					sql = ""
					sql = sql & " SELECT "
					sql = sql & " PH_Consumo.Id_Hogar "
					sql = sql & " FROM ((((PH_Consumo INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar) INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado) INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area) INNER JOIN PH_TipoConsumo ON PH_Consumo.id_TipoConsumo = PH_TipoConsumo.Id_TipoConsumo) INNER JOIN ss_Estado ON PH_GAreaEstado.Id_Estado = ss_Estado.Id_Estado "
					sql = sql & " WHERE "
					sql = sql & " PH_Consumo.Fecha_Creacion = '" & sFecha & "'"
					sql = sql & " AND PH_GAreaEstado.Id_Area = " & idArea 
					if idEstado <> 0 then
						sql = sql & " AND PH_PanelHogar.Id_Estado = " & idEstado
					end if
					sql = sql & " AND PH_Consumo.id_TipoConsumo = " & gDatosSol1(5,iReg)
					sql = sql & " AND PH_Consumo.Id_Semana = " & idSemana
					sql = sql & " AND PH_Consumo.Status_registro = 'G' "
					sql = sql & " AND PH_Consumo.Id_Hogar > 1 "
					sql = sql & " GROUP BY "
					sql = sql & " PH_Consumo.Id_Hogar "
					'response.write "<br>232 sql:= " & sql
					rsx2.Open sql ,conexion
					if rsx2.eof then
						rsx2.close
					else 
						gDatosSol2 = rsx2.GetRows
						rsx2.close
						Total = 0
						for iReg1 = 0 to ubound(gDatosSol2,2)
							Total = Total + 1
						next
						
					end if
					response.write "<td>" & Total & "</td>"
				response.write "</tr>"
				'Response.flush
			next 
			
			
			Response.write "<tr>"
		
		end if
		%>
	</table>
    <%conexion.close%>


</body>
</html>