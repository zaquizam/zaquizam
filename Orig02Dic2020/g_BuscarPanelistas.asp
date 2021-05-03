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
	sql = sql & " PH_Panelistas.Id_Panelista, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Nombre2, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Apellido2, "
	sql = sql & " PH_Nacionalidad.Nacionalidad, "
	sql = sql & " PH_Panelistas.Cedula, "
	sql = sql & " PH_Panelistas.Celular, "
	sql = sql & " PH_EstadoCivil.EstadoCivil, "
	sql = sql & " PH_Panelistas.Correo, "
	sql = sql & " PH_Parentesco.Parentesco, "
	sql = sql & " PH_Panelistas.Fec_Nacimiento, "
	sql = sql & " PH_Sexo.Sexo, "
	sql = sql & " PH_Educacion.Educacion, "
	sql = sql & " PH_TipoIngreso.TipoIngreso "
	sql = sql & " FROM PH_Panelistas LEFT OUTER JOIN "
	sql = sql & " PH_TipoIngreso ON PH_Panelistas.Id_TipoIngreso = PH_TipoIngreso.Id_TipoIngreso LEFT OUTER JOIN "
	sql = sql & " PH_Educacion ON PH_Panelistas.Id_Educacion = PH_Educacion.Id_Educacion LEFT OUTER JOIN "
	sql = sql & " PH_Sexo ON PH_Panelistas.Id_Sexo = PH_Sexo.Id_Sexo LEFT OUTER JOIN "
	sql = sql & " PH_Parentesco ON PH_Panelistas.Id_Parentesco = PH_Parentesco.Id_Parentesco LEFT OUTER JOIN "
	sql = sql & " PH_EstadoCivil ON PH_Panelistas.Id_EstadoCivil = PH_EstadoCivil.Id_EstadoCivil LEFT OUTER JOIN "
	sql = sql & " PH_Nacionalidad ON PH_Panelistas.Id_Nacionalidad = PH_Nacionalidad.Id_Nacionalidad "
	sql = sql & " WHERE "
	sql = sql & " PH_Panelistas.Id_Hogar = " & sBus
	sql = sql & " AND PH_Panelistas.ResponsablePanel = 0 "
	'response.write "<br>36 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	iExiste = 0
	if rsx1.eof then
		%>
		<div id="DivBuscarPanelistas">
			<h3>Personas del Hogas Registradas 0</h3>
		</div>
		<%
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
		%>
		<div id="DivBuscarPanelistas">
			<h3>Personas del Hogas Registradas <%=ubound(gDatosSol,2)+1%></h3>
			<div class="ex1">
				<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
					<thead>
						<tr class="w3-blue">
							<th>Sel.</th>
							<th>Primer Nombre</th>
							<th>Segundo Nombre</th>
							<th>Primer Apellido</th>
							<th>Segundo Apellido</th>
							<th>Nacionalidad</th>
							<th>Cedula</th>
							<th>Celular</th>
							<th>EstadoCivil</th>
							<th>Correo</th>
							<th>Parentesco</th>
							<th>Fecha Nacimiento</th>
							<th>Sexo</th> 
							<th>Educacion</th>
							<th>Tipo Ingreso</th>
						</tr>
					</thead>
					<%
					for iReg = 0 to ubound(gDatosSol,2)
						Response.write "<tr>"
							sx = "'" & gDatosSol(0,iReg) & "'"
							Response.write "<td>" 
							%>
							<img src="images/PDF02.jpg"  style="margin-left:0px;" alt="Modificar" width="20px"' onclick="SelPanelista(<%=sx%>)"/>
							<%
							Response.write "</td>"
							for ib = 1 to 14
								if ib <> 11 then
									Response.write "<td>" & gDatosSol(ib,iReg) & "</td>"
								else
									if gDatosSol(ib,iReg) = "1900-01-01" then
										Response.write "<td></td>"
									else
										sFecha = gDatosSol(ib,iReg) 
										sAno = mid(sFecha,1,4)
										sMes = mid(sFecha,6,2)
										sDia = mid(sFecha,9,2)
										sFecha = sDia & "/" & sMes & "/" & sAno
										Response.write "<td>" & sFecha & "</td>"
									end if
								end if
							next
						Response.write "</tr>"
					next
					%>
				</table>
			</div>
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


</style>	
