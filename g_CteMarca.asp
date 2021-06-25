<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ycat=Request.QueryString("cat")
	yfab=Request.QueryString("fab")
	yOpc = "0" 

	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Marca, "
	sql = sql & " rtrim(Marca) "
	sql = sql & " FROM "
	sql = sql & " PH_DataCrudaMensual "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria =  " & ycat
	if yfab <> "" then
		sql = sql & " and Id_Fabricante In (" & yfab & ")"
		Seleccionado = "Selected"
	end if
	sql = sql & " GROUP BY "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca "
	sql = sql & " HAVING "
	sql = sql & " Id_Marca <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Marca "
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
	
		<!--Marca-->	
		<label for="marca"><i class="fas fa-industry"></i>&nbsp;Marca:</label>
		<select id="Marca" multiple="multiple">				
			<% for iMar = 0 to  ubound(gDatosSol,2) %>
				<option value="<%=gDatosSol(0,iMar)%>"<%=Seleccionado%>  ><%=gDatosSol(1,iMar)%></option>
			<% next %>
		</select>
					
		<%
	end if

%>

<script type="text/javascript">
	$(document).ready(function() {
		$("#Marca").multiselect('destroy');
		$('#Marca').multiselect();
	});
</script>