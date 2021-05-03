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
	sql = sql & " ID, "
	sql = sql & " Cracha, "
	sql = sql & " nombres, "
	sql = sql & " acronimo, "
	sql = sql & " cargo_actual, "
	sql = sql & " pdtcargo "
	sql = sql & " from "
	'sql = sql & " ppi_postulantesbr "
	sql = sql & " ppi_postulantes "
	sql = sql & " where Cracha = '" & sBus & "'"
	'response.write "<br>36 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	iExiste = 0
	if rsx1.eof then
		response.write "Cracha No Existe"
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
		%>
		<div id="DivInformacion">
			<div class="ex1">
				<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:800px; margin-left:auto; margin-right:auto;margin-top:10px ">
					<thead>
						<tr class="w3-blue">
							<th>ID.</th>
							<th>Cracha</th>
							<th>Nomre y Apellido</th>
							<th>Acronimo</th>
							<th>Cargo_Actual</th>
							<th>Cargo PPI</th>
						</tr>
					</thead>
					<%
					sx = "'" & gDatosSol(0,0) & "'"
					Response.write "<td>" 
						%>
						<img src="no.png"  style="margin-left:0px;" alt="Agregar" width="20px"' onclick="eliminar(<%=sx%>)"/>
						<%
					Response.write "</td>"
					for ib = 1 to 5
						Response.write "<td>" & gDatosSol(ib,0) & "</td>"
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
  height: 100px;
  overflow: scroll;
}
</style>	


