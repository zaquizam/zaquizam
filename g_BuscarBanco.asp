<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ynum=Request.QueryString("num")
	yOpc = "0"

	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " id_Banco, "
	sql = sql & " Banco, "
	sql = sql & " Codigo "
	sql = sql & " FROM "
	sql = sql & " PH_Banco "
	sql = sql & " Where "
	sql = sql & " ind_Activo = 1 "
	sql = sql & " ORDER BY "
	sql = sql & " Banco "
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
	<div id="DivBanco"> 
		<select name="Banco" id="Banco" >
			<option value="0">Seleccionar</option> 
			<%
			sSeleccion =""
			for iReg = 0 to ubound(gDatosSol,2)
				if int(gDatosSol(0,iReg)) = int(yOpc) and yOpc <> "" then
					sSeleccion =" selected"
				else
					sSeleccion =""
				end if
				Response.write "<option value=" &  gDatosSol(0,iReg)  & sSeleccion &">" & gDatosSol(1,iReg) & "(" & gDatosSol(2,iReg) & ")" & "</option>"
			next
			%>
		</select>
	</div> 
	<%
	
%>