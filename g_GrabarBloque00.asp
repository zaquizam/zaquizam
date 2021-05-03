<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	'response.write "<br>LLEGO10" 
	'response.end
	ynum=Request.QueryString("num") 
	ynom1= Request.QueryString("nom1")
	ynom2 = Request.QueryString("nom2")
	yape1 = Request.QueryString("ape1")
	yape2 = Request.QueryString("ape2")
	ynaci = Request.QueryString("naci")
	ycedu = Request.QueryString("cedu")
	ycel1 = Request.QueryString("cel1")
	ycel2 = Request.QueryString("cel2")
	ynume = Request.QueryString("nume")
	ycor1 = Request.QueryString("cor1")
	ycor2 = Request.QueryString("cor2")
	ypare = Request.QueryString("pare")
	yesta = Request.QueryString("esta")
	yfech = Request.QueryString("fech")
	ysexo = Request.QueryString("sexo")
	yeduc = Request.QueryString("educ")
	ytipo = Request.QueryString("tipo")
	ynump = Request.QueryString("nump")
	yfrec = Request.QueryString("frec")
	ytitu = Request.QueryString("titu")
	ycedt = Request.QueryString("cedt")
	ybanc = Request.QueryString("banc")
	ycuen = Request.QueryString("cuen")
	ypago = Request.QueryString("pago")
	'response.write "<br>LLEGO36" 
	'response.end

	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	dim rsx3
	set rsx3 = CreateObject("ADODB.Recordset")
	rsx3.CursorType = 0
	rsx3.LockType = 3

	'response.write "<br>LLEGO49" 
	'response.end
	
	sql = ""
	sql = sql & " SELECT * "
	sql = sql & " FROM "
	sql = sql & " PH_Panelistas "
	sql = sql & " Where Id_Hogar = " & ynum
	'response.write "<br>57 sql:=" & sql
	'response.end
    rsx3.Open sql ,conexion
	
	iExiste = 0
	
	if rsx3.eof then
		rsx3.addNew
		iExiste = 0
	else
		iExiste = 1
	end if

	rsx3("Id_Hogar") = ynum
	rsx3("Nombre1") = ynom1
	rsx3("Nombre2") = ynom2
	rsx3("Apellido1") = yape1
	rsx3("Apellido2") = yape2
	rsx3("ResponsablePanel") = 1
	rsx3("Id_Nacionalidad") = ynaci
	rsx3("Cedula") = ycedu
	rsx3("Celular") = ycel1
	rsx3("CelularAdicional") = ycel2
	rsx3("NumeroCortesia") = ynume
	rsx3("Correo") = ycor1
	rsx3("CorreoAlterno") = ycor2
	rsx3("Titular") = ytitu
	rsx3("CedulaTitular") = ycedt
	rsx3("Id_Banco") = ybanc
	rsx3("NumeroCuenta") = ycuen
	rsx3("Id_PagoRapido") = ypago
	rsx3("Id_Parentesco") = ypare
	rsx3("Id_EstadoCivil") = yesta
	rsx3("Fec_Nacimiento") = yfech
	rsx3("Id_Sexo") = ysexo
	rsx3("Id_Educacion") = yeduc
	rsx3("Id_FrecuenciaCompra") = yfrec
	rsx3("Id_TipoIngreso") = ytipo
	rsx3("CantidadPersonas") = ynump
	rsx3("Ind_Activo") = 1
	rsx3.Update
	rsx3.Close 
	set rsx3 = nothing 
	
	dim rsx4
	set rsx4 = CreateObject("ADODB.Recordset")
	rsx4.CursorType = 0
	rsx4.LockType = 3

	sql = ""
	sql = sql & " SELECT * "
	sql = sql & " FROM "
	sql = sql & " PH_PanelHogar "
	sql = sql & " Where Id_PanelHogar = " & ynum
	'response.write "<br>57 sql:=" & sql
	'response.end
    rsx4.Open sql ,conexion
	
	iExiste = 0
	
	if rsx4.eof then
		iExiste = 0
	else
		iExiste = 1
	end if

	rsx4("TotalPersonas") = ynump
	rsx4.Update
	rsx4.Close 
	set rsx4 = nothing 

	
	
	%>

	
	<%
	
%>
