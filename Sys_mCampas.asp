<!DOCTYPE HTML>
<html >
<head>
	<title>Cambiar Clave</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>
<body topmargin="0">
<!--#include file="estiloscss.asp"-->
<!--#include file="meta.asp"-->
<!--#include file="nn_subN.asp"-->
<!--#include file="in_DataEN.asp"-->

<%
  
    dim PassAnt
    dim PassNue
    dim PassCon
    Dim iPass
  '  dim sUsu
    sPro="Sys_mCamPas.asp"

'Apertura
Sub cUsuario
    iPass=0
	'response.write "<br> Paso2"	
    'exit sub
	set rs 				= CreateObject("ADODB.Recordset")
	rs.CursorType 		= 0 		' CursorType = Forward-Only
	rs.LockType 		= 1			'LockType = Read-Only
	rs.CursorLocation 	= 3	'CursorLocation = adUseClient
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Usuario, "
	sql = sql & " Pass_Clave "
	sql = sql & " FROM "
	sql = sql & " ss_Usuarios "
	sql = sql & " WHERE "
	sql = sql & " Usuario = '" & Session("Usuario") & "'"
	'response.Write "<br>52 sql:= " & sql 
	'response.end
	rs.Open sql,conexion
	
	if not rs.eof then
		sPass = rs.fields("Pass_Clave")
	else
	    'Response.Write("<br> El Usuario no se encuentra en Usuario Cliente " & Session("usu") )	
		Response.Write("<script>alert('El Usuario no se encuentra registrado');</script>") 
		iPass=1
		exit sub
	End if
	'response.Write "pass " & sPass & "Usuario:=" & sUsu & "<br>"
	rs.close
	if sPass <> PassAct Then
		'Response.Write("La contraseña actual es incorrecta " )
		Response.Write("<script>alert('La contrasea actual es incorrecta');</script>") 
		iPass=2
		exit sub
	end if
	'response.Write "<br>64 "
	'response.end
	'response.write "<br> Paso3"	
	sql = ""
	sql = sql & " UPDATE "
	sql = sql & " ss_Usuarios "
	sql = sql & " SET Pass_Clave = '" & PassNue & "'"
	sql = sql & " WHERE "
	sql = sql & " Usuario= '" & Session("Usuario") & "'"
	rs.Open sql,conexion
	session("password")=PassNue
	'Response.Write("<br>La Contraseña fue cambiada con Éxito")
	Response.Write("<script>alert('La Contrase&ntilde;a fue cambiada con Éxito');</script>") 	
	iPass=0    
end sub

sub ValidarPassword
	'Response.Write("La contraseña nueva ")
	iErrorPass = 0
	'Falta Contraseña Actual
	if PassAct = "" then
		'Response.Write("Debe incluir la contraseña actual")
		Response.Write("<script>alert('Debe incluir la contrase&ntilde;a actual');</script>") 
		exit sub
	end if
	'Falta Contraseña Nueva
	if PassNue = "" then
		'Response.Write("Debe incluir la contraseña nueva")
		Response.Write("<script>alert('Debe incluir la contrase&ntilde;a nueva');</script>") 
		exit sub
	end if
	'Falta Contraseña de Confirmación
	if PassCon = "" then
		'Response.Write("Debe incluir la confirmacion de la contraseña")
		Response.Write("<script>alert('Debe incluir la confirmacion de la contrase&ntilde;a');</script>") 
		exit sub
	end if
	'Contraseña Nueva <> Contraseña de Confirmación
	if PassNue  <> PassCon  then
		'Response.Write("La contraseña Nueva debe ser igual a la confirmación")
		Response.Write("<script>alert('La contrase&ntilde;a Nueva debe ser igual a la confirmación');</script>") 
		exit sub
	end if
	'Contraseña Nueva debe tener mas de 4 caracteres
	' if len(PassNue) < 5 then
		' 'Response.Write("La contraseña Nueva debe tener más de 4 Caracteres")
		' Response.Write("<script>alert('La contraseña Nueva debe tener más de 4 Caracteres');</script>") 
		' exit sub
	' end if
	'response.write "<br> Paso1"
	cUsuario
  

end Sub
	PassAct = request.Form("pass1")
	PassNue = request.Form("pass2")
	PassCon = request.Form("pass3")
	'response.write "<br>146 pass1:= " & PassAct
	if isnull(PassAct) then PassAct = ""
	if isnull(PassNue) then PassNue = ""
	if isnull(PassCon) then PassCon = ""

	iErrorPass=0


'==========================================================================================
' Plano Principal
'==========================================================================================
	
	sUsu=Session("email")
	Sub ParDat
	end Sub
    Apertura
    LeePar
	%>
	<!--#include file="encabezado.asp"-->		
	<%
	Encabezado
	ed_CalPar 1,ed_iCla,ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
	%>
    <table width="98%" border="0" align="center" bgcolor="#ffffff">
	
    
	
	<%if PassAct <> "" or PassNue <> "" or PassCon <> "" then 	%>
	    <center>
	    <div style="font-family:Verdana; color:#ff0000">
	    <%ValidarPassword%>
	    </div>
	    </center>
	 <%end if %>  
	 <%
	 Calpar
	 sPar = sPar & "&perusu=" & Request.QueryString("perusu")
	 
	 
	 
	 %>
<div class="w3-card-4 w3-half w3-display-topmiddle " style="margin-top:200px">
  <div class="w3-container w3-theme ">
    <h4>Cambio de Contrase&ntilde;a</h4>
  </div>
  <form class="w3-container" method="post" action="<%=sPar%>" >
    <div class="w3-container">
    <p>      
    <label class="w3-text-theme"><b>Contrase&ntilde;a actual</b></label>
   
    <input type="password" name="pass1" size="20" maxlength="20" value ='<%=PassAct%>' class="w3-input w3-border w3-theme-l4" ID="Password1" />
    </p>
    <p>      
    <label class="w3-text-theme"><b>Contrase&ntilde;a nueva</b></label>
    <input type="password" name="pass2" size="20" maxlength="20" value ='<%=PassNue%>' class="w3-input w3-border w3-theme-l4" ID="Password5" />
    </p>
        <p>      
    <label class="w3-text-theme"><b>Confirmar contrase&ntilde;a</b></label>
    <input type="password" name="pass3" size="20" maxlength="20" value ='<%=PassCon%>'  class="w3-input w3-border w3-theme-l4" ID="Password6"  >
    <p>
    <input  type="submit" value="Cambiar" ID="Submit1" NAME="Submit1"  class="w3-button w3-theme-action"></p>
    </div>
  </form>
</div>



</body>
</html>