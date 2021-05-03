<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ynum=Request.QueryString("num")
	yciu=Request.QueryString("ciu")
	yOpc = "0"

	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " PH_OperadoraCable.Id_OperadoraCable, "
	sql = sql & " PH_OperadoraCable.OperadoraCable "
	sql = sql & " FROM PH_OperadoraCableCiudad INNER JOIN PH_OperadoraCable ON PH_OperadoraCableCiudad.Id_OperadoraCable = PH_OperadoraCable.Id_OperadoraCable "
	sql = sql & " WHERE "
	sql = sql & " PH_OperadoraCableCiudad.Id_Ciudad = " & yciu 
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br> Linea 223 " &
	'response.end
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
		iExiste = 1
	end if
	
	%>
	<div id="DivCableras2"> 
		<select name="Cableras2" id="Cableras2" >
			<option value="0">Seleccionar</option> 
			<%
			sSeleccion =""
			for iReg = 0 to ubound(gDatosSol,2)
				if int(gDatosSol(0,iReg)) = int(yOpc) and yOpc <> "" then
					sSeleccion =" selected"
				else
					sSeleccion =""
				end if
				Response.write "<option value=" &  gDatosSol(0,iReg)  & sSeleccion &">" & gDatosSol(1,iReg) & "</option>"
			next
			%>
		</select>
	</div> 
	<%
	
%>