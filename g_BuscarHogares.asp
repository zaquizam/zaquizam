<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
'==========================================================================================
' Variables y Constantes
'==========================================================================================

	sBus=Request.QueryString("num")

	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
    sql = sql & " SELECT "
	sql = sql & " PH_PanelHogar.Id_PanelHogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " ss_Estado.Estado, "
	sql = sql & " PH_Ciudad.Ciudad, "
	sql = sql & " ss_Municipio.Municipio, "
	sql = sql & " ss_Parroquia.Parroquia, "
	sql = sql & " PH_PanelHogar.Calle, "
	sql = sql & " PH_PanelHogar.Edificio, "
	sql = sql & " PH_PanelHogar.Casa, "
	sql = sql & " PH_PanelHogar.Escalera, "
	sql = sql & " PH_PanelHogar.Piso, "
	sql = sql & " PH_PanelHogar.Apto, "
	sql = sql & " PH_PanelHogar.Barrio, "
	sql = sql & " PH_PanelHogar.TelefonoLocal, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Cedula, "
	sql = sql & " PH_Panelistas.Correo, "
	sql = sql & " PH_Panelistas.Celular, "
	sql = sql & " PH_Panelistas.Titular "
	sql = sql & " FROM ((((PH_PanelHogar INNER JOIN ss_Estado ON PH_PanelHogar.Id_Estado = ss_Estado.Id_Estado) INNER JOIN ss_Parroquia ON PH_PanelHogar.Id_Parroquia = ss_Parroquia.Id_Parroquia) INNER JOIN PH_Ciudad ON PH_PanelHogar.Id_Ciudad = PH_Ciudad.Id_Ciudad) INNER JOIN ss_Municipio ON PH_PanelHogar.Id_Municipio = ss_Municipio.Id_Municipio) LEFT JOIN PH_Panelistas ON PH_PanelHogar.Id_PanelHogar = PH_Panelistas.Id_Hogar "
	sql = sql & " WHERE "
	sql = sql & " PH_PanelHogar.Ind_Activo = 1 "
	sql = sql & " and (PH_Panelistas.ResponsablePanel = 1 "
	sql = sql & " or PH_Panelistas.ResponsablePanel is null) "
	if sBus <> "" then
		sql = sql & " and ( PH_PanelHogar.CodigoHogar like '%" & sBus & "%'"
		sql = sql & " or  ss_Estado.Estado like '%" & sBus & "%'"
		sql = sql & " or  PH_Ciudad.Ciudad  like '%" & sBus & "%'"
		sql = sql & " or  ss_Parroquia.Parroquia like '%" & sBus & "%'"
		sql = sql & " or  PH_Panelistas.Nombre1 like '%" & sBus & "%'"
		sql = sql & " or  PH_Panelistas.Apellido1 like '%" & sBus & "%'"
		sql = sql & " or  PH_PanelHogar.TelefonoLocal like '%" & sBus & "%')"
	end if
	sql = sql & " Order by "
	sql = sql & " PH_PanelHogar.Id_PanelHogar Desc "
	'response.write "<br>36 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
		sMensajeReg = "Total Registros: 0"
		%>
		</br>
		</br>	
		<%=sMensajeReg%>
		<div class="ex1">
			<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:800px; margin-left:auto; margin-right:auto;margin-top:10px ">
				<thead>
					<tr class="w3-blue">
						<th>Sel.</th>
						<th>Codigo</th>
						<th>Estado</th>
						<th>Ciudad</th>
						<th>Municipio</th>
						<th>Parroquia</th>
						<th>Calle</th>
						<th>Edificio</th>
						<th>Casa</th>
						<th>Escalera</th>
						<th>Piso</th>
						<th>Apto</th>
						<th>Barrio</th>
						<th>TelefonoLocal</th>
						<th>Nombre1</th>
						<th>Apellido1</th>
						<th>Cedula</th>
						<th>Correo</th>
						<th>Celular</th>
						<th>Titular</th>
					</tr>
				</thead>
				<%
				if iTot > 0 then
					for iReg = 0 to iTot - 1
						Response.write "<tr>"
							sx = "'" & gDatosSol(0,iReg) & "'"
							Response.write "<td>" 
							%>
							<img src="images/si.png"  style="margin-left:0px;" alt="Agregar" width="20px"' onclick="SelHogar(<%=sx%>)"/>
							<%
							Response.write "</td>"
							for ib = 1 to 19
								Response.write "<td>" & gDatosSol(ib,iReg) & "</td>"
							next
						Response.write "</tr>"
					next
				end if
				%>
			</table>
		</div>
		<%
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
		iExiste = 1
		iTot = ubound(gDatosSol,2) + 1
		sMensajeReg = "Total Registros: " & iTot
		%>
		</br> 
		</br>	
		<%=sMensajeReg%>
		<div class="ex1">
			<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:800px; margin-left:auto; margin-right:auto;margin-top:10px ">
				<thead>
					<tr class="w3-blue">
						<th>Sel.</th>
						<th>Codigo</th>
						<th>Estado</th>
						<th>Ciudad</th>
						<th>Municipio</th>
						<th>Parroquia</th>
						<th>Calle</th>
						<th>Edificio</th>
						<th>Casa</th>
						<th>Escalera</th>
						<th>Piso</th>
						<th>Apto</th>
						<th>Barrio</th>
						<th>TelefonoLocal</th>
						<th>Nombre1</th>
						<th>Apellido1</th>
						<th>Cedula</th>
						<th>Correo</th>
						<th>Celular</th>
						<th>Titular</th>
					</tr>
				</thead>
				<%
				if iTot > 0 then
					for iReg = 0 to iTot - 1
						Response.write "<tr>"
							sx = "'" & gDatosSol(0,iReg) & "'"
							Response.write "<td>" 
							%>
							<img src="images/si.png"  style="margin-left:0px;" alt="Agregar" width="20px"' onclick="SelHogar(<%=sx%>)"/>
							<%
							Response.write "</td>"
							for ib = 1 to 19
								Response.write "<td>" & gDatosSol(ib,iReg) & "</td>"
							next
						Response.write "</tr>"
					next
				end if
				%>
			</table>
		</div>
		<%



	end if
	'response.end
%>
<style>
div.ex1 {
  background-color: LightSkyBlue;
  width: 850px;
  height: 300px;
  overflow: scroll;
}

}
</style>	
