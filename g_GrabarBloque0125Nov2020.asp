
<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	'response.write "<br>LLEGO25" 
	'response.end
	ynum=Request.QueryString("num") 
	ypai=Request.QueryString("pai")
	yciu=Request.QueryString("ciu")
	yest=Request.QueryString("est")
	ymun=Request.QueryString("mun")
	ypar=Request.QueryString("par")
	ycal=Request.QueryString("cal")
	yedi=Request.QueryString("edi")
	ycas=Request.QueryString("cas")
	yesc=Request.QueryString("esc")
	ypis=Request.QueryString("pis")
	yapt=Request.QueryString("apt")
	ybar=Request.QueryString("bar")
	yref=Request.QueryString("ref")
	ytel=Request.QueryString("tel")
	yusu=Request.QueryString("usu")
	uauario = yusu
	
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	dim rsx2
	set rsx2 = CreateObject("ADODB.Recordset")
	rsx2.CursorType = adOpenKeyset 
	rsx2.LockType = 2 'adLockOptimistic 

	dim rsx3
	set rsx3 = CreateObject("ADODB.Recordset")
	rsx3.CursorType = 0
	rsx3.LockType = 3

	if ynum = "" then
		sql = ""
		sql = sql & " insert  into PH_PanelHogar"
		sql = sql & " ( "
		sql = sql & " Id_Pais, "
		sql = sql & " Id_Estado, "
		sql = sql & " Id_Ciudad, "
		sql = sql & " Id_Municipio, "
		sql = sql & " Id_Parroquia, "
		sql = sql & " Calle, "
		sql = sql & " Edificio, "
		sql = sql & " Casa, "
		sql = sql & " Escalera, "
		sql = sql & " Piso, "
		sql = sql & " Apto, "
		sql = sql & " Barrio, "
		sql = sql & " Referencia, "
		sql = sql & " TelefonoLocal, "
		sql = sql & " Ind_Activo "
		sql = sql & " ) "
		sql = sql & " values "
		sql = sql & " ( "
		sql = sql & " '" & ypai & "',"
		sql = sql & " '" & yest & "',"
		sql = sql & " '" & yciu & "',"
		sql = sql & " '" & ymun & "',"
		sql = sql & " '" & ypar & "',"
		sql = sql & " '" & ycal & "',"
		sql = sql & " '" & yedi & "',"
		sql = sql & " '" & ycas & "',"
		sql = sql & " '" & yesc & "',"
		sql = sql & " '" & ypis & "',"
		sql = sql & " '" & yapt & "',"
		sql = sql & " '" & ybar & "',"
		sql = sql & " '" & yref & "',"
		sql = sql & " '" & ytel & "',"
		sql = sql & " '1'"
		sql = sql & " ) "
		'response.write "<br>220 sql:=" & sql
		'response.end
		rsx1.Open sql ,conexion

		sql = ""
		sql = sql & " SELECT SCOPE_IDENTITY() As lastID "
		rsx2.Open sql ,conexion
		recordID = rsx2("lastID")
		rsx2.close
    else
		sql = ""
		sql = sql & " Select * from PH_PanelHogar "
		sql = sql & " Where Id_PanelHogar = " & cint(ynum)
		'response.write "<br>220 sql:=" & sql
		'response.end
		rsx3.Open sql ,conexion
		rsx3("Id_Pais") = ypai
		rsx3("Id_Estado") =	yest	
		rsx3("Id_Ciudad") =	yciu	
		rsx3("Id_Municipio") = ymun
		rsx3("Id_Parroquia") = ypar	
		rsx3("Calle") = ycal
		rsx3("Edificio") = yedi		
		rsx3("Casa") = ycas
		rsx3("Escalera") = yesc	
		rsx3("Piso") = ypis
		rsx3("Apto") = yapt	
		rsx3("Barrio") = ybar
		rsx3("Referencia") = yref		
		rsx3("TelefonoLocal") = ytel
		rsx3("Ind_Activo") = 1
		rsx3.Update
		rsx3.Close 
		'set rsx3 = nothing 
		recordID = ynum
	end if

	dim gDatosSol
	dim rsx4
	set rsx4 = CreateObject("ADODB.Recordset")
	rsx4.CursorType = adOpenKeyset 
	rsx4.LockType = 2 'adLockOptimistic 
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " count(id_PanelHogar) "
	sql = sql & " FROM "
	sql = sql & " PH_PanelHogar"
	sql = sql & " Where "
	sql = sql & " id_ciudad =  " & cint(yciu)
	sql = sql & " Group By "
	sql = sql & " id_ciudad "
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx4.Open sql ,conexion
	if rsx4.eof then
		iTotal = 0
	else
		gDatosSol = rsx4.GetRows
		iTotal = gDatosSol(0,0)
		rsx4.close
	end if
	'response.end
	ytot = cstr(iTotal)
	
	ilen = len(yciu)
	if ilen = 1 then yciu = "0" & yciu

	ilen = len(yusu)
	if ilen = 1 then yusu = "0" & yusu

	ilen = len(ytot)
	if ilen = 1 then ytot = "000" & ytot
	if ilen = 2 then ytot = "00" & ytot
	if ilen = 3 then ytot = "0" & ytot
	
	yano = year(date)
	yano = mid(yano,3,2)

	ycod = yciu & ytot & yusu & yano
	
	'response.write "<br> ciudad:" & yciu
	'response.write "<br> usuario:" & yusu
	'response.write "<br> ano:" & yano
	'response.write "<br> total:" & ytot
	'response.write "<br> codigo:" & ycod 

	sql = ""
	sql = sql & " Select * from PH_PanelHogar "
	sql = sql & " Where Id_PanelHogar = " & cint(recordID)
	'response.write "<br>220 sql:=" & sql
	'response.end
	rsx3.Open sql ,conexion
	rsx3("CodigoHogar") = ycod
	rsx3("Id_Usuario") = cint(usuario)
	rsx3("Fec_Registro") = date()
	rsx3.Update
	rsx3.Close 
	set rsx3 = nothing 
	
	%>
	<div id="DivHogar"> 
		<input type="text" name="Hogar" disabled id="Hogar" value="<%=recordID%>" size=6 style="text-align:right; background-color:#d1d1d1;">
		<input type="text" name="Hogar" disabled id="Hogar" value="<%=ycod%>" size=10 style="text-align:right; background-color:#d1d1d1;">
	</div>

	<%
	
%>