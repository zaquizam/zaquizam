<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ynum=Request.QueryString("num")
	yreg=Request.QueryString("reg")
	yOpc = "0" 
	if ynum = "" then ynum = "0"
	
	if yreg<> "" then 
		dim gDatosSol2
		dim rsx2
		set rsx2 = CreateObject("ADODB.Recordset")
		rsx2.CursorType = adOpenKeyset 
		rsx2.LockType = 2 'adLockOptimistic 

		sql = ""
		sql = sql & " SELECT "
		sql = sql & " Id_Ciudad "
		sql = sql & " FROM "
		sql = sql & " PH_PanelHogar "
		sql = sql & " WHERE "
		sql = sql & " Id_PanelHogar = " & cint(yreg)
		'response.write "<br>220 sql:=" & sql
		'response.end
		rsx2.Open sql ,conexion
		'response.write "<br> Linea 223 " &
		'response.end
		gDatosSol2 = rsx2.GetRows
		rsx2.close
		yOpc = gDatosSol2(0,0)
		'response.write "<br>37 Ciudad:=" & yOpc
	end if
	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Ciudad, "
	sql = sql & " Ciudad "
	sql = sql & " FROM "
	sql = sql & " PH_Ciudad "
	sql = sql & " WHERE "
	sql = sql & " Id_Estado = " & cint(ynum)
	sql = sql & " and ind_activo = 1 "
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br> Linea 223 " &
	'response.end
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
		%>
			<select name="Ciudad" id="Ciudad" >
				<option value="0">Seleccionar</option> 
			</select>
		<%
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
		iExiste = 1
		%>
			<select name="Ciudad" id="Ciudad" >
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
		<%
	end if

%>