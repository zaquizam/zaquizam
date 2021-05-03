<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ycat=Request.QueryString("cat")
	ymar=Request.QueryString("mar")
	yOpc = "0" 

	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Fabricante, "
	sql = sql & " rtrim(Fabricante)"
	sql = sql & " FROM "
	sql = sql & " PH_DataCrudaMensual "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & ycat
	Seleccionado = ""
	if ymar <> "" then
		sql = sql & " and Id_Marca In (" & ymar & ")"
		Seleccionado = "Selected"
	end if
	sql = sql & " GROUP BY "
	sql = sql & " Id_Fabricante, "
	sql = sql & " Fabricante "
	sql = sql & " HAVING "
	sql = sql & " Id_Fabricante <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Fabricante "
	'response.write "<br>372 sql1:=" & sql
	'response.end
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
		iExiste = 0
		%>
		
		<%
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
		iExiste = 1
		%>

		<!--Fabricante-->	
		<label for="fabricante"><i class="fas fa-industry"></i>&nbsp;Fabricante:</label>
		<select id="Fabricante2" multiple="multiple">				
			<% for iFra = 0 to  ubound(gDatosSol,2) %>
				<option value="<%=gDatosSol(0,iFra)%>"<%=Seleccionado%>  ><%=gDatosSol(1,iFra)%></option>
			<% next %>
		</select>
					
		<%
	end if

%>
<script type="text/javascript">
	$(document).ready(function() {
		$('#Fabricante2').multiselect();
</script>