<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<meta charset="utf-8">
<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ynum=Request.QueryString("num")
	yusu=Request.QueryString("idusu")
	yOpc = ""

	dim gDatosSol0
	dim rsx0
	set rsx0 = CreateObject("ADODB.Recordset")
	rsx0.CursorType = adOpenKeyset 
	rsx0.LockType = 2 'adLockOptimistic 

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Area "
	sql = sql & " FROM "
	sql = sql & " ss_Usuarios "
	sql = sql & " Where "
	sql = sql & " Id_Usuario = " & int(yusu)
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx0.Open sql ,conexion
	iExiste = 0
	if rsx0.eof then
		iExiste = 0
	else
		gDatosSol0 = rsx0.GetRows
		rsx0.close
		iExiste = 1
	end if
	iArea= int(gDatosSol0(0,0))

	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " ss_Estado.Id_Estado, "
	sql = sql & " ss_Estado.Estado "
	sql = sql & " FROM ss_Estado INNER JOIN ss_AreaEstado ON ss_Estado.Id_Estado = ss_AreaEstado.Id_Estado "
	sql = sql & " WHERE "
	sql = sql & " ss_Estado.Id_Pais = 1 "
	if iArea > 0 then
		sql = sql & " AND ss_AreaEstado.Id_Area = " & iArea
	end if 
	sql = sql & " ORDER BY "
	sql = sql & " ss_Estado.Estado "
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
	<div id="DivEstado"> 
		<select name="Estado" id="Estado" onchange="buscar_municipio()" style="float:left;" >
			<option value="0">Seleccionar</option> 
			<%
			sSeleccion =""
			for iReg = 0 to ubound(gDatosSol,2)
				if int(gDatosSol(0,iReg)) = yOpc and yOpc <> "" then
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