<!DOCTYPE HTML>
<html >
<head>
	<title>nn_subN</title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />     
    <link  href="w3.css" rel="stylesheet">
</head>
    
<style type="text/css">

.campas 	  {font-family:verdana;color:#000000;font-size: 9pt;text-decoration: none;padding:0px 0px 0px 0px}
.campas a     {font-family:verdana;color:#000000;font-size: 9pt;text-decoration: none;padding:0px 0px 0px 0px}
.campas:hover {font-weight: bold;color:#800000; cursor:pointer}

.e_boton11 	    {border-radius: 5px;font-family:verdana;border: 1px solid #9eb6ce;  font-weight: bold;color:#000000;width:300px; background-color: #ffcc00;font-size: 10pt;text-align: center;text-decoration: none;padding:3px 100px 3px 100px}
.e_boton11:hover {border-radius: 5px;                    border: 1px solid #9eb6ce;  font-weight: bold;color:#000000;background-color:#ffcc00; cursor:pointer}

.e_boton12 	    {border-radius: 5px;font-family:verdana;border: 1px solid #9eb6ce;  font-weight: bold;color:#000000;width:100%; background-color: #ffcc00;font-size: 10pt;text-align: center;text-decoration: none;padding:3px 5px 3px 5px}
.e_boton12:hover {border-radius: 5px;                    border: 1px solid #9eb6ce;  font-weight: bold;color:#000000;background-color:#ffcc00; cursor:pointer}

.e_boton13 	      {border-radius: 10px;width:400px;background-color:#ffcc00;color:#000000 ;padding:8px;text-decoration: none;}
.e_boton13 a      {border-radius: 10px;width:400px;background-color:#ffcc00; padding:8px;text-decoration: none;}
.ed_l2				{font-family:verdana;font-weight:normal;background-color: #ffffff;font-size: 9pt;text-align: center;color: #000000;text-decoration: none;margin-left:10px}
.ed_l2 a			{font-family:verdana;font-weight:normal;color: #000000;text-decoration: none;}
.ed_l2 a:hover		{font-weight:bold;font-size: 9pt;color: #800000	;text-decoration: none;}

.e_boton14 	    {border-radius: 5px;width:100px;background-color:#ffcc00; padding:3px}
.e_boton15	    {border-radius: 10px;width:100px;background-color:#ffcc00; padding:8px}
.tdd {padding:5px 10px 0px 10px;text-align:right;font-family:verdana;font-weight:normal;color: #000000;text-decoration: none;}
.tdd a {padding:5px 10px 0px 10px;text-align:right;font-family:verdana;font-weight:normal;color: #000000;text-decoration: none;}
.tdd a:hover {padding:5px 10px 0px 10px;text-align:right;font-family:verdana;font-weight:normal;color: #800000;text-decoration: none;}
.tdi{padding:5px 10px 0px 10px;text-align:left;font-family:verdana;font-weight:normal;color: #000000;text-decoration: none;}
.tdi a {padding:5px 10px 0px 10px;text-align:left;font-family:verdana;font-weight:normal;color: #000000;text-decoration: none;}
.tdi a:hover {padding:5px 10px 0px 10px;text-align:left;font-family:verdana;font-weight:normal;color: #800000;text-decoration: none;}
.tdc {padding:5px 10px 0px 10px;text-align:center;font-family:verdana;font-weight:normal;color: #000000;text-decoration: none;}

</style>

<%
   
Session.Timeout =150
'session("urlant")=Request.ServerVariables("URL")
'==========================================================================================
' Variables y Constantes
'==========================================================================================
' Variables de Sesion
    dim iPerUsu
	dim iSession
	dim idUsu	' ID del Usuario		
	dim sUsu
	dim idEmpleado
' Variables de base de datos    
    dim rse ' Empleado
    dim rsp ' Periodo
    dim rc  ' Consorcio
    dim conexion
	dim sql    
    dim sqlPer ' Query de periodo
    dim sqlCom ' Query de Compañia o EMpresa
    dim sqlPro ' Query de Proveedor
    dim sqlQui
    
' Variables de parametros	
    dim ipCon ' Consorcio
	dim ipPer
	dim ipQui
    dim ipCom  ' Empresa	
    dim sPar
    dim sxPar
    dim ixPar
    dim stPar ' texto de parametros 
    dim spMen ' texto del menu
' Variables Varias
	dim isw
	dim iMes
	dim iAno
	dim iPerMes
    dim iCierre ' 1=Si	
    dim inCombo
	

' Dudosas
	dim sIP
    dim d ' Datos del cuestionario
    dim p
    Set d=Server.CreateObject("Scripting.Dictionary")
    Set p=Server.CreateObject("Scripting.Dictionary") ' Diccionario de Preguntas


	Const adOpenKeyset = 1
	Const adLockOptimistic = 1
	Response.Buffer=True 



'==========================================================================================
' Leer variable de session
'==========================================================================================
iSession=Session("idSesion")
sUsu = session("usu")
IdUsu=session("idusu")
'iPerUsu=Session("PerUsu")

ipCon=Session("Consorcio")
idEmpleado=Session("idEmpleado")
iPerUsu = Session("idPerfil")

if session("idusu")="" then 
    sPro="default.asp"
    Calpar
   ' response.write "<br> " & sPAr
   ' response.end
    response.Redirect(sPar)
    
end if    
    
'==========================================================================================
' Parametros
'==========================================================================================
Sub LeePar

	iMontoRec = ipPer=Request.Cookies("montorec")	
	ipPer=Request.QueryString("cc_p1")
    if ipPer="" then ipPer=Request.Cookies("cc_p1")	
    if ipPer="" then 
        im=month(now)
        iy=Year(now())
        iPerMes=iy*12+im
        ipPer=iPerMes
    end if
	
    iAno=2019 'int(ipPer/12)
    iMes=5 'ipPer-iAno*12
    if imes=0 then iMes=12:iAno=iAno-1
	ipCom=Request.QueryString("cc_p2")
    if ipCom="" then ipCom=Request.Cookies("cc_p2")		
	if ipCom="" then ipCom=Session("idEmpresa")
	if ipCom="" then ipCOm=1
'Response.write "<br>125 ipCom:=" & ipCom


    ipQui=Request.QueryString("cc_p3")
    if ipQui="" then ipQui=Request.Cookies("cc_p3")	    
    if ipQui="" then ipQui=1
    
    spMen=Request.QueryString("smenu")   

	ipTie=Request.QueryString("p_tie")
	icla=Request.QueryString("edcla")
   ' if icla="999999" then icla=""
        
'==========================================================================================
' Sql
'==========================================================================================    
    sqlCom= " SELECT id_Empresa, Empresa FROM  ss_Empresa WHERE Fec_Inactivo is Null And id_Consorcio=" & ipCon & " Order By Empresa"
    SqlPer = "SELECT IdPeriodo, Periodo FROM ss_Periodo WHERE (fec_inactivo is null) ORDER BY idPeriodo DESC" 
    sqlQui = "SELECT CodMQS, Descripcion FROM ss_MQS WHERE (((ss_MQS.FrecMQS)=2));"
    sqlPro= " SELECT id_Proveedor, Proveedor FROM  ac_Proveedor WHERE Fec_Inactivo is Null And id_Empresa=" & ipCom & " Order By Proveedor"    
End Sub

Sub Calpar
    sPar=sPro & "?x=1"
    sPar=sPar & sxPar & ixPar    
    sPar= sPar & "&smenu=" & spMen
	
    
End SUb





'==========================================================================================
' Abrir Base de Datos 
'==========================================================================================   
Sub Apertura
	Set conexion = Server.CreateObject("ADODB.Connection")
	
	'conexion.ConnectionString= "Provider=SQLOLEDB.1;Password='PHaa11..**';Persist Security Info=True;User ID=Atenas_PH;Initial Catalog=Atenas_PH;Data Source=199.79.62.22"
	'conexion.ConnectionString= "Provider=SQLOLEDB.1;Password='profit';Persist Security Info=True;User ID=profit;Initial Catalog=cacevedo_atenas;Data Source=REYESHUERTA-PC\SQL2014"
	conexion.ConnectionString= "Provider=SQLOLEDB.1;Password='PHaa11..**';Persist Security Info=True;User ID=cacevedo_atenas;Initial Catalog=cacevedo_atenas;Data Source=192.185.6.37"

	
	conexion.mode = 3
	conexion.Open


	'if isnull(idEMpleado) then gIdEmpleado
	'gAcceso
	'response.write "<br>paso Apertura"
End Sub

Sub lConsorcio
    sql = " SELECT * from ss_Consorcio WHERE (Id_Consorcio=" & Session("Consorcio") & ") "
    set rc = server.CreateObject("ADODB.Recordset")
	rc.CursorType = 0
	rc.LockType = 1
'response.write "<br>213 " & sql	
    rc.open sql,conexion

end sub	
Sub mError (sErr)
%>
        <div style="color:#ff0000; font-size:32; font-family:Verdana;  font-weight:bold; border-radius: 10px; padding:10px 10px 10px 10px; border: solid 1px #cccccc; width:50%; margin-top:40px; text-align:center">
            <%=sErr %>
        </div>
   
<%  response.end
end Sub
Sub mAdver (sErr)
%>
        <div style="color:#ff0000; font-size:32; font-family:Verdana;  font-weight:bold; border-radius: 10px; padding:10px 10px 10px 10px; border: solid 1px #cccccc; width:50%; margin-top:40px; text-align:center">
            <%=sErr %>
        </div>
   
<%end Sub

Sub gCookies
    for i=1 to 9
        sx="cc_p"&i
'response.write "<br> ix2:=" & sx            
        ix=Request.QueryString(sx)  
        if ix<>"" then
            Response.Cookies(sx)=ix
            Response.Cookies(sx).Expires=now()+365
        end if        
    next    
end sub

Sub chkCierre

    sql = " SELECT * from nn_Cierre WHERE (Id_Empresa=" & ipCom & ") AND (id_Periodo=" & ipPer & ") AND (FrecMQS=2) AND (CODMqs="& ipQui & ")"
    set rs = server.CreateObject("ADODB.Recordset")
	rs.CursorType = 0
	rs.LockType = 1
'response.write "<br>213 " & sql	
    rs.open sql,conexion
    
    iCierre=0
    if rs.eof then
    else
        if rs("fec_Cierre")<>"" then iCierre=1
    end if
    rs.close
end Sub

Sub gIdEmpleado


	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = adOpenKeyset 
	rs.LockType = 1 'adLockOptimistic 
	sql = " SELECT  Id_Empleado "
    sql = sql & " FROM nn_Empleado"
    sql = sql & " WHERE "
    sql = sql & "(Correo='" & Session("Usuario") & "') "
'    response.write "<br>205 " & sql
    rs.open sql,conexion
	if rs.eof then exit sub

 ' response.write "<br>Grabar Empleado" &  rs1("IdEmpleado") 	
	set rs1 = CreateObject("ADODB.Recordset")
	rs1.CursorType = adOpenKeyset 
	rs1.LockType = 3 'adLockOptimistic 
    sql =" SELECT "
	sql = sql & " * "
	sql = sql & " FROM "
	sql = sql & " ss_U_Usuarios "
    sql = sql & " WHERE "
    sql = sql & " Usuario='" & Session("Usuario") & "'"
    rs1.open sql,conexion
    if rs1.eof then exit sub
    rs1("IdEmpleado")=rs("Id_Empleado")
    idEmpleado=rs1("IdEmpleado")
    Session("idEmpleado")=rs1("IdEmpleado")
	rs1("IP")=Request.ServerVariables("REMOTE_ADDR")
	rs1("Fec_Ult_Mod")=Now()
	rs1("Usr")=Session("Usuario")
	rs1("idSession")=ed_iSession
	rs1.update
	rs1.close
	rs.close

end sub
' Actualizar inasistencias con dias no laborales
Sub aInasistencias

    ix=ipPer-iPerMes
    if ix<-2 then exit sub
  '  response.write "<br>Actualizar Dias no Laborales" & ix
    
	set rs1 = CreateObject("ADODB.Recordset")
	rs1.CursorType = adOpenKeyset 
	rs1.LockType = 3 'adLockOptimistic 


	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = adOpenKeyset 
	rs.LockType = 1 'adLockOptimistic 
	
    if ipCom = "" then ipCom = 1
    
    sql = " SELECT  * "
    sql = sql & " FROM ss_DiasNOLAborales "
    sql = sql & " WHERE "
    sql = sql & "(id_Empresa=" & ipCom & ") AND (id_Periodo=" & ipPer & ")" 
	
'response.write "<br>" & Sql    
	rs.open sql,conexion
	if rs.eof then exit sub


    while not rs.eof
      '  response.write "<br>" & rs("dia")	

        sql = " SELECT  nn_Inasistencia.* "
        sql = sql & " FROM nn_Empleado INNER JOIN nn_Inasistencia ON nn_Empleado.Id_Empleado = nn_Inasistencia.Id_Empleado "
        sql = sql & " WHERE "
        sql = sql & "((nn_Empleado.Id_Empresa)=" & ipCom & ") AND ((id_Periodo) =" & ipPer & ") AND  (Day(nn_Inasistencia.Fec_Desde)=" & rs("dia") & ")  AND (day(nn_Inasistencia.Fec_Hasta)=" & rs("dia") & ") AND ((Id_TipoInasistencia)=1)"
        'response.write "<br>" & sql

	    rs1.open sql,conexion
	    if not(rs1.eof) then
	        while not rs1.eof
	            if rs("Ind_Empresa")=true then
	                rs1("Id_TipoInasistencia")=9
	                rs1("motivo")=rs("descripcion")
	                'response.write "<br>123 " & ipPer & " dia:=" & dia & " Descripcion " & Empresa
	            else
	                rs1("Id_TipoInasistencia")=7
	                rs1("motivo")=rs("descripcion")
	               ' response.write "<br>123 " & ipPer & " dia:=" & dia & " Descripcion " & festivo
	            end if 
	            rs1("IP")=Request.ServerVariables("REMOTE_ADDR")
		        rs1("Fec_Ult_Mod")=Now()
		    	rs1("Usr")=Session("Usuario")
		        rs1("idSession")=ed_iSession
		        rs1.update
	            rs1.movenext
	        wend
	    end if
	    rs1.close
	    rs.movenext
	wend
	rs.close
	
end sub
' Actualizar Empleados de cada Supervisor
Sub aSupervisor
    set rs1 = CreateObject("ADODB.Recordset")
	rs1.CursorType = adOpenKeyset 
	rs1.LockType = 3 'adLockOptimistic 


	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = adOpenKeyset 
	rs.LockType = 1 'adLockOptimistic 
	
    
    
    sql = " SELECT  * "
    sql = sql & " FROM nn_Empleado "
    sql = sql & " WHERE "
    sql = sql & "(fec_inactivo is null)"
    
	rs.open sql,conexion
	if rs.eof then exit sub


	

    while not rs.eof
     
 ' Chequar si esta grabado el mismo
        sql = " SELECT  * "
        sql = sql & " FROM nn_Supervisor_Empleado "
        sql = sql & " WHERE "
        sql = sql & "((Id_Supervisor)=" &  rs("Id_Empleado")& " ) AND ((id_Empleado)=" & rs("Id_Empleado") & ")"
        rs1.open sql,conexion


        if rs1.eof then
            rs1.addnew
	        rs1("Id_Supervisor")=rs("Id_Empleado")
	        rs1("Id_Empleado")=rs("Id_Empleado")
            rs1("IP")=Request.ServerVariables("REMOTE_ADDR")
	        rs1("Fec_Ult_Mod")=Now()
	   	    rs1("Usr")=Session("Usuario")
	        rs1("idSession")=ed_iSession
	        rs1.update
	    end if
	    rs1.close  
	
        sql = " SELECT  * "
        sql = sql & " FROM nn_Supervisor_Empleado "
        sql = sql & " WHERE "
        sql = sql & "((Id_Supervisor)=" & rs("Id_Supervisor")& " ) AND ((id_Empleado)=" & rs("Id_Empleado") & ")"
        'response.write "<br>" & sql
	
	    rs1.open sql,conexion
	    if rs1.eof then
	            rs1.addnew
	            rs1("Id_Supervisor")=rs("Id_Supervisor")
	            rs1("Id_Empleado")=rs("Id_Empleado")
	            rs1("IP")=Request.ServerVariables("REMOTE_ADDR")
		        rs1("Fec_Ult_Mod")=Now()
		    	rs1("Usr")=Session("Usuario")
    	        rs1("idSession")=ed_iSession
		        rs1.update
	    end if
	    rs1.close
	    rs.movenext
	wend
	rs.close
	
end sub
Sub cCombo
   ' dim ed_sCombo(10,6) ' 0 = Titulo , 1 Sql , 2 <>Null es un total, 3 Filtro, 4 width tabla, 5 width del titulo , 6 width del combo
   
' Periodo   
    ed_iCombo=3 
    ed_sCombo(1,0)="Periodo"
    
    sx = "SELECT IdPeriodo, Periodo FROM ss_Periodo WHERE ((fec_inactivo is null) "
    ides=ipPer-3
    ihas=iPper+2
    sx = sx & "AND (idPeriodo>=" & ides & " AND idPeriodo<=" & ihas & "))"
    sx = sx & " ORDER BY idPeriodo DESC"     
    ed_sCombo(1,1)= sx
    
    
    ed_sCombo(2,0)="Empresa"
    ed_sCombo(2,1)="SELECT id_empresa, empresa FROM ss_Empresa WHERE  id_consorcio=" & ipCon & " ORDER by Empresa "
    if IdUsu=3 then
        ed_sCombo(2,1)="SELECT id_empresa, empresa FROM ss_Empresa WHERE  (id_consorcio=" & ipCon & ") AND (Id_Empresa=1) ORDER by Empresa "
    end if
    
    ed_sCombo(3,0)="Quincena"
	sql="SELECT CodMQS, Descripcion FROM ss_MQS WHERE (((ss_MQS.FrecMQS)=2));"
    ed_sCombo(3,1)=sql
    if inCombo<>4 then exit sub
    ed_iCombo=4
    sql= " SELECT nn_Empleado.Id_Empleado,"
   	sql = sql & " nn_Empleado.PrimerNombre+' ' " & "+"
	'sql = sql & " nn_Empleado.SegundoNombre+' ' " & "+"
	sql = sql & " nn_Empleado.PrimerApellido "
    sql = sql & " FROM nn_Supervisor_Empleado INNER JOIN nn_Empleado ON nn_Supervisor_Empleado.Id_Empleado = nn_Empleado.Id_Empleado "
    sql = sql & " WHERE (((nn_Supervisor_Empleado.Id_Supervisor)=" & idEmpleado & ") AND ((nn_Empleado.Id_Empresa)=" & ipCom & ")); "
    ed_sCombo(4,0)="Empleado"
    ed_sCombo(4,1)=sql
end sub
Sub mDrow (ixDro)
    set rs = server.CreateObject("ADODB.Recordset")
	rs.CursorType = 1
	rs.LockType = 1
    rs.open ed_sCombo(ixDro,1),conexion
    if rs.eof then exit sub
	dim gX
    gX=rs.getrows
    
    ed_sPar(ixDro,1)=gX(0,0)
    if isnull(ed_sPar(ixDro,0)) or ed_sPar(ixDro,0)="" then 
        if ed_sCombo(ixDro,2)<>"" then 
                ed_sPar(ixDro,0)=ed_sCombo(ixDro,2)
             else   
                ed_sPar(ixDro,0)=gX(0,0)
             end if   
    end if    
    rs.close   
  
    wO=0
    for j=0 to ubound(gX,2)
       sP=ed_sPar(ixDro,0)
    	ed_sPar(ixDro,0)=gx(0,j)
    	ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_iFil,ed_iMp,ed_iMs
    	ed_sPar(ixDro,0)=sP
    	    	   
		if isnumeric(ed_sPar(ixDro,0)) then
			ix=ed_sPar(ixDro,0) -gX(0,j)
	    else
	       ix=1
	       if ed_sPar(ixDro,0) =gX(0,j) then ix=0
	    end if
	    sCla= "w3-bar-item w3-button w3-small"
	    'check_circle
	    if ix=0 then 
	        'sCla= sCla & " w3-theme-l1"
	       wO=1
	    end if%>
     	    
            <a href="<%=sPar%>" class="<%=sCla %>"><%=gX(1,j) %>
            <% if ix=0 then
                stPar= stPar& gX(1,j)%>
            <i class="material-icons w3-margin-righ w3-right w3-medium" >done</i>
            <% end if %>
            </a>   	    
    <%next%>
    
    <%if wO=0 then 
        if ed_sCombo(ixDro,2)<>"" then 
            ed_sPar(ixDro,0)=ed_sCombo(ixDro,2)
        else   
	        ed_sPar(ixDro,0)=ed_sPar(ixDro,1) 
        end if    
     end if
        
end sub

Sub Encabezado
	iPerUsu = Session("perusu")

    if ed_iPas=4 then exit sub
'response.write "<br><br><br>Paso-1" 
'response.end
	%>
	<!-- Sidebar -->
	<div class="w3-sidebar w3-bar-block  w3-border-right w3-medium" style="display:none; width:25% " id="mySidebar">
		<button onclick="w3_close()" class="w3-bar-item  w3-button     w3-large">Close &times;</button>
		<%
		'response.write "<br><br><br>Paso0" 
		sp=sPro
		'response.write "<br><br><br>Paso1" 
		mMenu1 
		%>
	</div>
	<%
	sPro=sp 
	%>
	<!-- Page Content -->
	<div  class=" ">
		<div class="w3-container w3-theme-d2" >
			<button class="w3-button w3-theme-d2 w3-large  w3-left" onclick="w3_open()"><i class="material-icons">menu</i></button>
			<%
			sP=sPro
			sPro="pr_mInicio.asp"
			ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_iFil,ed_iMp,ed_iMs
			sPro=sP
			%>
			<a href="pr_mInicio.asp	" class="w3-button w3-theme-d2" title="Ir al Inicio" ><i class="material-icons">home</i></a>
			<div class="w3-dropdown-hover w3-theme-d2 ">
				<div class="w3-dropdown-content w3-bar-block w3-border">
					<% 
					stPar=""
					'mDrow 2 
					stPar=stPar & "-"
					%>
					<div class="w3-theme-d2" style="height:2px"></div>
					<%
					'mDrow 3 
					stPar=stPar & "-"
					%>
					<div class="w3-theme-d2" style="height:2px"></div>
					<% 
					'mDrow 1 
					%>
				</div>
			</div>    
			<%
			sP=sPro
			sPro="default.asp"
			ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_iFil,ed_iMp,ed_iMs
			sPro=sP
			%>
			<a href="<%=sPar %>" class="w3-button w3-theme-d2"  title="Salir del Sistema"><i class="material-icons">power_settings_new</i></a>
			<%
			sx="img_avatar2.png"
			%>
			<a href="Sys_mCampas.asp"  class="w3-button w3-theme-d2 " title="Cambiar Contrase&ntilde;a a  <%=Session("NomApe")%>"><img src="images/<%=sx %>"  style=" max-width:30px;" alt="<%=Session("NomApe")%>" /></a>
			<%
			'ed_vCombo 
			%>
			<%
			%>
		</div>
	</div>
	<div class="w3-container w3-theme-l1" title="<%=Session("NomApe")%>">
		<%
		=spMen 
		%>
	</div>
	<%
End Sub		

Sub EncabezadoRespaldoAbri2019
    if ed_iPas=4 then exit sub
	%>
	<!-- Sidebar -->
	<div class="w3-sidebar w3-bar-block  w3-border-right w3-medium" style="display:none; width:25% " id="mySidebar">
		<button onclick="w3_close()" class="w3-bar-item  w3-button     w3-large">Close &times;</button>
		<%
		sp=sPro 
		mMenu1 
		%>
	</div>
	<%
	sPro=sp 
	%>
	<!-- Page Content -->
	<div  class=" ">
		<div class="w3-container w3-theme-d2" >
			<button class="w3-button w3-theme-d2 w3-large  w3-left" onclick="w3_open()"><i class="material-icons">menu</i></button>
			<%
			sP=sPro
			sPro="pr_mInicio.asp"
			ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_iFil,ed_iMp,ed_iMs
			sPro=sP
			%>
			<a href="pr_mInicio.asp	" class="w3-button w3-theme-d2" title="Ir al Inicio" ><i class="material-icons">home</i></a>
			<div class="w3-dropdown-hover w3-theme-d2 ">
				<div class="w3-dropdown-content w3-bar-block w3-border">
					<% 
					stPar=""
					'mDrow 2 
					stPar=stPar & "-"
					%>
					<div class="w3-theme-d2" style="height:2px"></div>
					<%
					'mDrow 3 
					stPar=stPar & "-"
					%>
					<div class="w3-theme-d2" style="height:2px"></div>
					<% 
					'mDrow 1 
					%>
				</div>
			</div>    
			<%
			sP=sPro
			sPro="default.asp"
			ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_iFil,ed_iMp,ed_iMs
			sPro=sP
			%>
			<a href="<%=sPar %>" class="w3-button w3-theme-d2"  title="Salir del Sistema"><i class="material-icons">power_settings_new</i></a>
			<%
			sx="img_avatar2.png"
			%>
			<a href="Sys_mCampas.asp"  class="w3-button w3-theme-d2 " title="Cambiar Contrase&ntilde;a a  <%=Session("NomApe")%>"><img src="images/<%=sx %>"  style=" max-width:30px;" alt="<%=Session("NomApe")%>" /></a>
			<div class="w3-right w3-small" title="<%=iPerUsu& "c:" & ipCom %>  ">
			<%
			'ed_vCombo 
			%>
			<%
			%>
		</div>
	</div>
	<div class="w3-container w3-theme-l1" title="<%=Session("NomApe")%>">
		<%
		=spMen 
		%>
	</div>
	<%
End Sub		

Sub gDiasFestivos

	set rs1 = CreateObject("ADODB.Recordset")
	rs1.CursorType = adOpenKeyset 
	rs1.LockType = 3 'adLockOptimistic 


	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = adOpenKeyset 
	rs.LockType = 1 'adLockOptimistic 
	
    
    sql =  " SELECT  * "
    sql = sql & " FROM ss_Dias_Festivos "
    sql = sql & " WHERE "
    sql = sql & " id_Periodo=" & ipPer
	rs.open sql,conexion
	if rs.eof then exit sub
    
    while not rs.eof
	
   
        sql = " SELECT  * "
        sql = sql & " FROM ss_DiasNOLAborales "
        sql = sql & " WHERE "
        sql = sql & "(id_Empresa=" & ipCom & ") AND (id_Periodo=" & ipPer & ") AND  (dia=" & rs("dia") & ")"
   ' response.write "<br>" & sql
	    rs1.open sql,conexion
	    if rs1.eof then
	        rs1.addnew 
	        rs1("id_Empresa")=ipCom
	        rs1("id_Periodo")=ipPer
	        rs1("dia")=rs("dia")
	        rs1("descripcion")=rs("descripcion")
	        rs1("ind_Empresa")="false"
	        rs1("ind_Activo")="true"
	        rs1("IP")=Request.ServerVariables("REMOTE_ADDR")
		    rs1("Fec_Ult_Mod")=Now()
		 	rs1("Usr")=Session("Usuario")
		    rs1("idSession")=ed_iSession
		    rs1.update
	        'response.write "<br>123 " & ipPer & " dia:=" & dia & " Descripcion " & rs("descripcion")
	    end if
	    rs1.close
	    rs.movenext
	wend

end sub

Sub vCombo (ed_sql, ixPar)

    dim rst    
    dim gX
	set rst = server.CreateObject("ADODB.Recordset")
	rst.CursorType =1
	rst.LockType = 1
	
	'Combo Periodo
'response.write "<br>405 ed_sql=" & ed_sql	
	rst.open ed_sql,conexion
	if rst.eof then
	    ixPar=-1
	exit sub
	    mError "Error: No hay data para construir el combo"
	end if
    gX=rst.getrows

    if ixPar="" then ixPar=gX(0,0)
    rst.close    %>

	    	<select size="1" name="per" id="Select3" onchange ="location.href=this.options[this.selectedIndex].value" style="width:100%; font-size:12px; font-family:Verdana">
	    	
	    	<% ixP=ixPar
	    	   ixPar=0
	    	   ed_CalPar 1,ed_iCla,ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
	    	   ixPar=ixP
	    	%>
	        <option value="<%=sPar%>"  <% if ixPar=0 then response.Write"selected" %> style="width:100%; font-size:12px; font-family:Verdana; padding:5px 0 5px 5px " >
	            
				[Todos]
			</option>
    		<%
    		   sw=0
    		   for i=0 to ubound(gX,2)
    		       ixP=ixPar
    		       ixPar=gX(0,i)
    	    	   ed_CalPar 1,ed_iCla,ed_iPag,ed_sBus,ed_iCol,ed_iOrd, ed_ifil,ed_iMp,ed_iMs
    	    	   ixPar=ixP
    	    	   iSi=int(ixPar) -gX(0,i)
    	    	   if iSi=0 then sW=1
    	    	   'sPar=sPar & sVar & gX(0,i)
    	    	   
    	    %>
			<option value="<%=sPar%>"  <% if iSi=0 then response.Write"selected" %> style="width:100%; font-size:12px; font-family:Verdana; padding:5px 0 5px 5px " >
				<%=  gX(1,i) %>
			</option>
			<%next 
			if sw=0 then ixPar =gX(0,0)%>
	    </select>
<%	    
End sub



Sub lTabla(sql,e_Opc)
    'response.write "<br>563 sql:=" & sql
    sql=replace(sql,"**","")
    
    dim gX
	set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = adOpenKeyset 
	rs.LockType = 1 'adLockOptimistic 
	rs.open sql,conexion
	if rs.eof then exit sub
	gX=rs.getrows
	
	inum=0
	e_Opc(0,0)=0
	for i=0 to Ubound(gx,2)
		e_Opc(0,0)=e_Opc(0,0)+1
	    inum=iNum+1
	    e_Opc(inum,0)=gx(0,i)
	    e_Opc(inum,1)=gx(1,i)
	next 
	
end sub
    


sub GraPie
%>
  <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
    <script type="text/javascript">
      google.charts.load('current', {'packages':['corechart']});
      google.charts.setOnLoadCallback(drawChart);
      function drawChart() {

        var data = google.visualization.arrayToDataTable([
          ['<%=gGraPie(0,0) %>', '<%=gGraPie(0,1) %>'],
          <%for ig=1 to iGrapie %>
          ['<%=gGraPie(ig,0) %>',    <%=gGraPie(ig,1) %>],
          <%Next %>
        ]);

        var options = {
          title: '<%=sGraPie %>', 
          is3D: true

        };

        var chart = new google.visualization.PieChart(document.getElementById('piechart'));

        chart.draw(data, options);
      }
    </script>
 
   
<%
end sub

Sub GraVer
%>
   <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
    <script type="text/javascript">
      google.charts.load('current', {'packages':['bar']});
      google.charts.setOnLoadCallback(drawChart);
      function drawChart() {
        var data = google.visualization.arrayToDataTable([
          ['Year', 'Sales', 'Expenses', 'Profit'],
          ['2014', 1000, 400, 200],
          ['2015', 1170, 460, 250],
          ['2016', 660, 1120, 300],
          ['2017', 1030, 540, 350]
        ]);

        var options = {
          chart: {
            title: 'Company Performance',
            subtitle: 'Sales, Expenses, and Profit: 2014-2017'
          },
          bars: 'horizontal' // Required for Material Bar Charts.
        };

        var chart = new google.charts.Bar(document.getElementById('barchart_material'));

        chart.draw(data, options);
      }
    </script>

<%
end sub
Sub gAcceso

    dim rs4
    sPro=Request.ServerVariables("URL")

	' Grabar Entrada
	set rs4 = CreateObject("ADODB.Recordset")
	rs4.CursorType = adOpenKeyset 
	rs4.LockType = 3 'adLockOptimistic     
	ixDia=Year(date)*365+Month(date)*30+day(date)
  	
  	sIP=Request.ServerVariables("REMOTE_ADDR")
	sql=" SELECT * FROM ss_AccesosUsuarios "
	sql = sql & " WHERE id_Usuario=" & Session("idUsu") & " AND Num_Dia=" & ixDia & " AND Pagina='" & sPro & "'"&  " AND IP='" & sIP & "'"
	
	'response.write "<br>1256 sql:= " & sql
	rs4.open sql,conexion
	if rs4.eof then
	    rs4.addnew
	    sx=Request.ServerVariables("SERVER_NAME")
	    'rs4("Software")=sx  
	    rs4("Id_Empresa")=Session("idEmpresa")
	    rs4("Id_Usuario")= Session("idUsu")
	    rs4("Pagina")=sPro
	    rs4("Num_Accesos")=1
	    rs4("Fec_Acc")=Now()
	    rs4("Fec_Ult_Acc")=Now()
        rs4("Num_Dia")=ixDia
 	    rs4("IP")=sIP
 	    rs4("USR")=Session("Usuario")
        rs4("Fec_Ult_Mod")=Now()
        rs4("idSession")=iSession 
        'sx=Request.ServerVariables ("HTTP_REFERER") 	    
        if sx<>"" then rs4("Referer")=Mid(sx,1,255)
		'response.write "<br>808 Paso"
	'response.end
	else
		'response.write "<br>811 Paso"
		'response.end
	    rs4("Num_Accesos")=rs4("Num_Accesos")+1
        rs4("idSession")=iSession	    
	    rs4("Fec_Ult_Acc")=Now()
        rs4("IP")=sIp
	end if    
    rs4.update
	rs4.close


	
End sub	

Sub lPeriodo
	set rsp = server.CreateObject("ADODB.Recordset")
	rsp.CursorType = 0
	rsp.LockType = 1
    sql = " SELECT   * "
    sql = sql & " FROM ss_Periodo "	
    sql = sql & "WHERE idPeriodo= " & ipPer
	'sql = sql & " Order By Periodo Desc "
	rsp.open sql,conexion
	if rsp.eof then
        ierr1=1%>
        <div style=" text-align:center">
            <div style="font-size:12px; border: solid 1px #cccccc; font-weight:bolder; border-radius:10px; width:70%; padding:20px 10px 20px 10px; color:#c40000; text-align:center; vertical-align:middle">
                 Error: No Hay Indicadores Financieros el periodo <%=IpPer %><br />
                 No se procesó la nómina
            </div>
       </div><%         
       response.end
       exit sub
     end if  
    iMaxSalSSO=rsp("Salario_Minimo")*5
    iMaxSalBA=rsp("Salario_Minimo")*3 
end sub 

Function dtexto (sPag,sIni,sFin)
    sx=""
    i1=instr(sPag,sIni)
    if i1<>0 then
       i1=i1+len(sIni)
       ix1=instr(i1,sPag,sFin)
       ix=ix1-i1
       sx=mid(sPag,i1,ix)
    end if 
    dTexto=sx  
end function

Sub gIp
    dim xml
    sUrl="http://ip-api.com/xml/"& sIp&"?fields=country,countryCode,city,lat,lon,timezone,isp,mobile,proxy,query,status,message"
 'response.write "<br>1548 " & sUrl   
    Set xml =Server.CreateObject("Microsoft.XMLHTTP")
	xml.Open "GET",sUrl, False
		
	'on error resume next
	xml.Send
	iSta = xml.status
	Select Case iSta
    	case 200,302,304
			sPag=""
			sPag = xml.responseText
			isw=0
		case 404,401,500
		    isw=1
		case else
			isw=1
	end select
    if isw=1 then exit sub
    
    
' leer Estatus    
    ix=instr(sPag,"</status>")
    sx=mid(sPag,1,ix)
    ix=instr(sx,"success")
    if ix =0 then exit sub
' Quitar Textos    
    sx=Replace (sPag,"<![CDATA[","")
    sPag=sx
    sx=Replace (sPag,"]]>","")
    sPag=sx
    
' Abrir Tabla
	set rs1 = CreateObject("ADODB.Recordset")
	rs1.CursorType = adOpenKeyset 
	rs1.LockType = 3 'adLockOptimistic   
	sql=" SELECT * FROM S_IP "
	rs1.open sql,conexion
    rs1.addnew
 
 'response.write "<br>1586 " & sIP  
 
' Buscar Country
    rs1("IPx")=sIP
    sx=dTexto (sPag,"<country>", "</country>")
    rs1("Country")=sx
    sx=dTexto (sPag,"<countryCode>", "</countryCode>")
    rs1("CountryCode")=sx
    sx=dTexto (sPag,"<city>", "</city>")
    rs1("City")=sx
    sx=dTexto (sPag,"<lat>", "</lat>")
    rs1("Latitud")=sx
    sx=dTexto (sPag,"<lon>", "</lon>")
    rs1("Longitud")=sx
    sx=dTexto (sPag,"<timezone>", "</timezone>")
    rs1("timezone")=sx  
    sx=dTexto (sPag,"<isp>", "</isp>")
    rs1("Isp")=sx   
    sx=dTexto (sPag,"<mobile>", "</mobile>")
    rs1("Ind_Mobile")=sx   
    sx=dTexto (sPag,"<proxy>", "</proxy>")
    rs1("Ind_Proxy")=sx 
    rs1("idSession")=iSession
    rs1("Usr")=Session("Usuario")
    rs1("IP")=Request.ServerVariables("REMOTE_ADDR")
    rs1("Fec_Ult_Mod")=Now()
      
    rs1.update
    rs1.close
    Set xml = nothing		

End sub     
Sub CalculosBonos
    'exit sub
	set rsDataCampo = server.CreateObject("ADODB.Recordset")
    rsDataCampo.CursorType = 0
    rsDataCampo.LockType = 3

    set rs = server.CreateObject("ADODB.Recordset")
    rs.CursorType = 0
    rs.LockType = 3
	
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_DataCampo, "
	sql = sql & " Id_Periodo, "
	sql = sql & " Id_Empleado, "
	sql = sql & " TiendasPlanAudit, "
	sql = sql & " TiendasPlanEspecial, "
	sql = sql & " TiendasReaAudit, "
	sql = sql & " TiendasReaEspecial, "
	sql = sql & " Dias_Habiles_Actual, "
	sql = sql & " Dias_Descontar, "
	sql = sql & " Dias_Justificados, "
	sql = sql & " Dias_Pagar, "
	sql = sql & " Monto_BonoAlimentacion, "
	sql = sql & " Total_BonoAlimentacion, "
	sql = sql & " Monto_AyudaTransporte, "
	sql = sql & " Monto_AyudaInternet, "
	sql = sql & " Monto_Tienda, "
	sql = sql & " Total_Plan, "
	sql = sql & " Total_Vis, "
	sql = sql & " Porcentaje_Cancelar "
	sql = sql & " FROM "
	sql = sql & " nn_DataCampo "
	sql = sql & " WHERE "
	sql = sql & " Id_Periodo = " & ipPer 
	rsDataCampo.open sql,conexion
	'response.write "<br> sql:= " & sql
	while not rsDataCampo.eof
		iEmpleado = rsDataCampo("Id_Empleado")
		
		iTiePlanAud = rsDataCampo("TiendasPlanAudit")
		if isNull(iTiePlanAud) then iTiePlanAud = 0
		
		iTiePlanEsp = rsDataCampo("TiendasPlanEspecial")
		if isNull(iTiePlanEsp) then iTiePlanEsp = 0
		
		iTieVisAud = rsDataCampo("TiendasReaAudit")
		if isNull(iTieVisAud) then iTieVisAud = 0
		
		iTieVisEsp = rsDataCampo("TiendasReaEspecial")
		if isNull(iTieVisEsp) then iTieVisEsp = 0
		
		iDiasHabi = rsDataCampo("Dias_Habiles_Actual")
		if isNull(iDiasHabi) then iDiasHabi = 0
		
		iDiasDesc = rsDataCampo("Dias_Descontar")
		if isNull(iDiasDesc) then iDiasDesc = 0
		
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " nn_Empleado.Id_Empleado, "
		sql = sql & " nn_Contrato.Monto_BonoAlimenticio, "
		sql = sql & " nn_Contrato.Monto_Transporte, "
		sql = sql & " nn_Contrato.Monto_Internet "
		sql = sql & " FROM nn_Empleado INNER JOIN nn_Contrato ON nn_Empleado.Id_Contrato = nn_Contrato.Id_Contrato "
		sql = sql & " WHERE "
		sql = sql & " nn_Empleado.Id_Empleado = " & iEmpleado
		rs.open sql,conexion

		iMontoAlim = rs("Monto_BonoAlimenticio") / 30
		if isNull(iMontoAlim) then iMontoAlim = 0
		
		iMontoTran = rs("Monto_Transporte") 
		if isNull(iMontoTran) then iMontoTran = 0
		
		iMontoInte = rs("Monto_Internet") 
		if isNull(iMontoInte) then iMontoInte = 0
		rs.close

		sql = ""
		sql = sql & " SELECT "
		sql = sql & " nn_Inasistencia.Id_Periodo, "
		sql = sql & " nn_Inasistencia.Id_Empleado, "
		sql = sql & " nn_Inasistencia.Fec_Desde, "
		sql = sql & " nn_Inasistencia.Fec_Hasta, "
		sql = sql & " nn_Inasistencia.Id_TipoInasistencia, "
		sql = sql & " nn_Inasistencia.TiendasPlanAudit, "
		sql = sql & " ss_TipoInasistencia.Ind_PagoBono "
		sql = sql & " FROM ((nn_Inasistencia INNER JOIN nn_Empleado ON nn_Inasistencia.Id_Empleado = nn_Empleado.Id_Empleado) INNER JOIN ss_TipoInasistencia ON nn_Inasistencia.Id_TipoInasistencia = ss_TipoInasistencia.Id_TipoInasistencia) INNER JOIN ss_Empresa ON (ss_Empresa.Id_Consorcio = ss_TipoInasistencia.id_Consorcio) AND (nn_Empleado.Id_Empresa = ss_Empresa.Id_Empresa) "
		sql = sql & " WHERE "
		sql = sql & " nn_Inasistencia.Id_Periodo = " & ipPer - 1
		sql = sql & " AND nn_Inasistencia.Id_Empleado = " & iEmpleado
		sql = sql & " AND ss_TipoInasistencia.Ind_PagoBono = 1 " 
		rs.open sql,conexion
		iDiasJust = 0
		while not rs.eof
			iDiasJust = iDiasJust + 1
			iTiePlanJust = iTiePlanJust + rs("TiendasPlanAudit") 
			rs.movenext
			'response.write "<br>Paso"	
		wend
		if isNull(iDiasJust) then iDiasJust = 0
		if isNull(iTiePlanJust) then iTiePlanJust = 0
		
		rs.close
		rsDataCampo("Dias_Justificados") = iDiasJust
		iDiasPago = iDiasHabi - iDiasDesc + iDiasJust
		rsDataCampo("Dias_Pagar") = iDiasPago
		rsDataCampo("Monto_BonoAlimentacion") = iMontoAlim
		rsDataCampo("Total_BonoAlimentacion") = iDiasPago * iMontoAlim
		rsDataCampo("Monto_AyudaTransporte") = iTieVisAud * iMontoTran
		rsDataCampo("Monto_AyudaInternet") = iMontoInte
		rsDataCampo("Monto_Tienda") = iMontoTran
		iPlan = iTiePlanAud + iTiePlanEsp
		iVisi = iTieVisAud + iTieVisEsp
		iDescuento = (iVisi * 100) /iPlan
		iDescuento = 100 - iDescuento
		rsDataCampo("Total_Plan") = iPlan
		rsDataCampo("Total_Vis") = iVisi
		rsDataCampo("Porcentaje_Cancelar") = iDescuento
		rsDataCampo.update
		rsDataCampo.movenext
	wend
	rsDataCampo.close
	

End Sub

Sub cDataCampo


    'exit sub
	set rsDataCampo = server.CreateObject("ADODB.Recordset")
    rsDataCampo.CursorType = 0
    rsDataCampo.LockType = 3

    set rs = server.CreateObject("ADODB.Recordset")
    rs.CursorType = 0
    rs.LockType = 3
	
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_DataCampo, "
	sql = sql & " Id_Periodo, "
	sql = sql & " Id_Empleado, "
	sql = sql & " TiendasPlanAudit, "
	sql = sql & " TiendasPlanEspecial, "
	sql = sql & " TiendasReaAudit, "
	sql = sql & " TiendasReaEspecial, "
	sql = sql & " Dias_Habiles_Actual, "
	sql = sql & " Dias_Descontar, "
	sql = sql & " Dias_Justificados, "
	sql = sql & " Dias_Pagar, "
	sql = sql & " Monto_BonoAlimentacion, "
	sql = sql & " Total_BonoAlimentacion, "
	sql = sql & " Monto_AyudaTransporte, "
	sql = sql & " Monto_AyudaInternet, "
	sql = sql & " Monto_Tienda, "
	sql = sql & " Total_Plan, "
	sql = sql & " Total_Vis, "
	sql = sql & " Porcentaje_Cancelar "
	sql = sql & " FROM "
	sql = sql & " nn_DataCampo "
	sql = sql & " WHERE "
	sql = sql & " Id_Periodo = " & ipPer 

    sql= "SELECT * "
    sql = sql &" FROM (nn_DataCampo INNER JOIN nn_Empleado ON nn_DataCampo.Id_Empleado = nn_Empleado.Id_Empleado) INNER JOIN nn_Contrato ON nn_Empleado.Id_Contrato = nn_Contrato.Id_Contrato "
    sql = sql & " WHERE (((nn_DataCampo.Id_Periodo)=" & ipPer & ")) "
 'response.write "<br>989 " & sql
	rsDataCampo.open sql,conexion
    if rsDataCampo.eof then%>
    
        <div style=" text-align:center">
        <div style="font-size:12px; border: solid 1px #cccccc; font-weight:bolder; border-radius:10px; width:70%; padding:20px 10px 20px 10px; color:#c40000; text-align:center; vertical-align:middle">
             Error Grave: No esta creada la Data de Campo<br />
             No se procesó la nómina
                <br /><br />
                </div>
            </div> <%
        response.end
    End if        
            
            

	while not rsDataCampo.eof
		iEmpleado = rsDataCampo("Id_Empleado")

		iTiePlanAud = rsDataCampo("TiendasPlanAudit")
		if isNull(iTiePlanAud) then iTiePlanAud = 0
		
		iTiePlanEsp = rsDataCampo("TiendasPlanEspecial")
		if isNull(iTiePlanEsp) then iTiePlanEsp = 0
		
		iTieVisAud = rsDataCampo("TiendasReaAudit")
		if isNull(iTieVisAud) then iTieVisAud = 0
		
		iTieVisEsp = rsDataCampo("TiendasReaEspecial")
		if isNull(iTieVisEsp) then iTieVisEsp = 0
		
		iPlan=iTiePlanAud +iTiePlanEsp
		rsDataCampo("Total_Plan") = iPlan
		
		iVis=iTieVisAud + iTieVisEsp
		rsDataCampo("Total_Vis") = iVis
		
		if iVis <> 0 then 
			rsDataCampo("Porcentaje_Cancelar") = iVis/iPlan*100
		end if

' Dias a Pagar		
		iDiasHabi = rsDataCampo("Dias_Habiles_Actual")
		if isNull(iDiasHabi) then iDiasHabi = 0
		
		iDiasDesc = rsDataCampo("Dias_Descontar")
		if isNull(iDiasDesc) then iDiasDesc = 0
		
		iDiasJus = rsDataCampo("Dias_Justificados")
		if isNull(iDiasJus) then iDiasJus = 0
		
        ix=iDiasHabi-iDiasDesc+iDiasJus
		rsDataCampo("Dias_Pagar")=ix

'		
		iMontoTran = rsDataCampo("Monto_Transporte") 
		if isNull(iMontoTran) then iMontoTran = 0
		
		iMontoTranTie = rsDataCampo("Monto_TransporteTienda") 
		if isNull(iMontoTranTie) then iMontoTranTie = 0
		
		
		iMontoInte = rsDataCampo("Monto_Internet") 
		if isNull(iMontoInte) then iMontoInte = 0
		
		if iTiePlanAud = 0 Then iTiePlanAud = 1
		if iTieVisAud = 0 Then iTieVisAud = 1
	    rsDataCampo("Monto_AyudaInternet") = iMontoInte/iTiePlanAud*iTieVisAud
	    rsDataCampo("Monto_AyudaTransporte") = iTieVisAud * iMontoTranTie
       
' Bono Alimenticio
        rsDataCampo("Total_BonoAlimentacion")=rsDataCampo("Monto_BonoAlimenticio")
        rsDataCampo("Ind_CobraTienda1")=rsDataCampo("Ind_CobraTienda")
        if rsDataCampo("Ind_CobraTienda")=-1 then

            iMonto=rsDataCampo("Monto_BonoAlimenticio")* rsDataCampo("Porcentaje_Cancelar")/100
            rsDataCampo("Monto_BonoAlimentacion") = iMonto
        else            
            iMontoAlim = rsDataCampo("Monto_BonoAlimenticio") / 30
    	    rsDataCampo("Monto_BonoAlimentacion")=rsDataCampo("Dias_Pagar") *iMontoAlim
        end if
		    
	
		rsDataCampo.update
		rsDataCampo.movenext
	wend
	rsDataCampo.close
	

End Sub
Sub ReCalculoBonoAlimentacion
		'response.write "<br>1077 Paso"
		'response.flush
		set rs = CreateObject("ADODB.Recordset")
		rs.CursorType = 0
		rs.LockType = 1
		set rsAct = CreateObject("ADODB.Recordset")
		rsAct.CursorType = 0
		rsAct.LockType = 3
		sql = ""
		sql = sql & " Select  "
		sql = sql & " Id_Periodo, "
		sql = sql & " Id_Empleado, "
		sql = sql & " Dias_Habiles_Actual, "
		sql = sql & " Dias_Descontar, "
		sql = sql & " Dias_Justificados, "
		sql = sql & " Dias_Pagar, "
		sql = sql & " Monto_BonoAlimentacion, "
		sql = sql & " Total_BonoAlimentacion "
		sql = sql & " FROM "
		sql = sql & " nn_DataCampo "
		sql = sql & " WHERE "
		sql = sql & " Id_Periodo = " & ipPer
		rsAct.Open sql,conexion
		while not rsAct.eof		
			iEmpleado = rsAct("Id_Empleado") 
			iDiaHab = rsAct("Dias_Habiles_Actual") 
			iDiaDes = rsAct("Dias_Descontar") 
			iDiaJus = rsAct("Dias_Justificados")
			ix = 0
			if iDiaDes > 0 then
				sql = ""
				sql = sql & " Select  "
				sql = sql & " Count(nn_Inasistencia.Id_Inasistencia) AS CuentaDeId_Inasistencia "
				sql = sql & " FROM nn_Inasistencia INNER JOIN ss_TipoInasistencia ON nn_Inasistencia.Id_TipoInasistencia = ss_TipoInasistencia.Id_TipoInasistencia "
				sql = sql & " WHERE "
				sql = sql & " nn_Inasistencia.Id_Periodo = " & ipPer - 1
				sql = sql & " AND nn_Inasistencia.Id_Empleado = " & iEmpleado
				sql = sql & " AND ss_TipoInasistencia.Ind_PagoBono <> 0 "
				rs.Open sql,conexion
				if rs.eof then
					iDiaJus = 0
				else
					iDiaJus = rs("CuentaDeId_Inasistencia")
					ix = 1
				end if
				rs.Close
			end if
			iDiaPag = rsAct("Dias_Pagar")
			iBonPag = rsAct("Monto_BonoAlimentacion")
			iBonTot = rsAct("Total_BonoAlimentacion")
			
			iDiaPag = int(iDiaHab) - int(iDiaDes) + int(iDiaJus)
			iBonPag = (iBonTot/30) * iDiaPag
			
			rsAct("Dias_Justificados") = iDiaJus
			rsAct("Monto_BonoAlimentacion") = iBonPag
			rsAct.Update
			'if iEmpleado = 26 then
			'response.write "<br><br>Empleado:= " & iEmpleado
			'response.write "<br>iDiaHab:= " & iDiaHab
			'response.write "<br>iDiaDes:= " & iDiaDes
			'response.write "<br>iDiaJus:= " & iDiaJus
			'response.write "<br>iDiaPag:= " & iDiaPag
			'response.write "<br>iBonPag:= " & iBonPag
			'end if
			rsAct.movenext
		wend
		rsAct.Close

End Sub

'-----------------------------------------------------------
' Leer Empleados
'-----------------------------------------------------------
Sub lEmpleadoSub

'
     set rse = server.CreateObject("ADODB.Recordset")
     rse.CursorType = 0
     rse.LockType = 1
     
    sql = " SELECT "
    sql = sql &"ss_Empresa.Empresa,"
    sql = sql &"nn_Empleado.Id_Empresa,"
    sql = sql &"nn_Empleado.Id_Empleado,"
    sql = sql &"nn_Empleado.PrimerNombre,"
    sql = sql &"nn_Empleado.SegundoNombre,"
    sql = sql &"nn_Empleado.PrimerApellido,"
    sql = sql &"nn_Empleado.SegundoApellido,"
    sql = sql &"nn_Empleado.Id_TipoPersona,"
    sql = sql &"nn_Empleado.NumeroCedula,"
    sql = sql &"nn_Empleado.Fec_Nacimiento,"
    sql = sql &"nn_Empleado.Id_CiudadNacimiento,"
    sql = sql &"nn_Empleado.Id_Nacionalidad,"
    sql = sql &"nn_Empleado.Id_Sexo,"
    sql = sql &"nn_Empleado.Id_EstadoCivil,"
    sql = sql &"nn_Empleado.Direccion,"
    sql = sql &"nn_Empleado.Correo,"
    sql = sql &"nn_Empleado.TelfHabitacion,"
    sql = sql &"nn_Empleado.TelfCelular,"
    sql = sql &"nn_Empleado.img_Cedula,"
    sql = sql &"nn_Empleado.img_Foto,"
    sql = sql &"nn_Empleado.Id_NivelAcademico,"
    sql = sql &"nn_Empleado.Id_Profesion,"
    sql = sql &"nn_Empleado.Fec_Nacimiento,"
    sql = sql &"nn_Empleado.Fec_Ingreso,"
    sql = sql &"nn_Empleado.Fec_Egreso,"
    sql = sql &"nn_Empleado.Id_Contrato,"
    sql = sql &"nn_Empleado.Id_TipoNomina,"
    sql = sql &"nn_Empleado.Id_Departamento,"
    sql = sql &"nn_Empleado.Id_Cargo,"
    sql = sql &"nn_Empleado.Id_Supervisor,"
    sql = sql &"nn_Empleado.Porcentaje_Retencion_ISLR,"
    sql = sql &"nn_Empleado.CuentaBancaria,"
    sql = sql &"nn_Empleado.Ind_ActivoSSO, "
    sql = sql &"nn_Empleado.Sueldo, "
 
    sql = sql & " nn_Empleado.Fec_Contrato_Inicio, "
    sql = sql & " nn_Empleado.Fec_Contrato_Fin "
    
   
   sql = sql & " FROM ss_Empresa INNER JOIN nn_Empleado ON ss_Empresa.Id_Empresa = nn_Empleado.Id_Empresa "
    sql = sql & " WHERE      ("
    sql = sql & "((ss_Empresa.Id_Consorcio) = " & ipCon & ") AND "
    sql = sql & "  (nn_Empleado.Fec_Egreso IS NULL) "
    sql = sql & "  )  "
        
   ' response.write "<br>431 sql:= " & sql
	rse.open sql,conexion
	
'" SELECT id_Empresa, Empresa FROM  ss_Empresa WHERE Fec_Inactivo is Null And id_Consorcio=" & ipCon & " Order By Empresa"	
End sub	

Sub mMenu1
	iPerUsu = Session("perusu")
'response.write "<br><br><br>Pasaaaaa" 
'response.write "<br><br><br>1410 sql:=" 
'response.end
' Abrir Perfil
	set rs = CreateObject("ADODB.Recordset")
    sql = "SELECT  id_PerfilUsuario, PerfilUsuario, link_acceso, ocultar, idGrupo , mostrar"
    sql = sql & " FROM  ss_PerfilUsuario "
	sql = sql & " WHERE (((id_Perfilusuario)=" & iPerUsu & ") AND ((Fec_Inactivo) Is Null)) "
	'response.write "<br><br><br>1410 sql:=" & sql   
	'exit sub
'response.	write "<br><br><br>Paso" 

'response.end
	
	rs.Open sql,conexion
	e_gPerUsu=rs.GetRows
	rs.close
	set rs = nothing
'exit sub

	
' Abrir Menu
	set rs = CreateObject("ADODB.Recordset")
    sql = "SELECT  idNivel, Menu, Link , Txt_Tips, Target, Parametros, id_menu, idgrupo, Ind_Activo, ind_Oculto , icon, Path"
    sql = sql & " FROM  ss_Menu "
	sOcultar=e_gPerUsu(3,0)
	sMostrar=e_gPerUsu(5,0)
	'iGrup = e_gPerUsu(4,0)
	iGrup = 1
    sql = sql & " WHERE ((IdGrupo)=" & iGrup & ")  AND ((Fec_Inactivo) Is Null) "
	if sOcultar<>"" then sql = sql & " AND (id_Menu NOT IN (" & sOcultar& "))"    
	if sMostrar<>"" then sql = sql & " AND (id_Menu IN (" & sMostrar& "))"    
	sql = sql & " AND (ind_activo=1) And (ind_Oculto = 0 )"

    sql = sql & " ORDER BY Orden,Menu ; "
'response.write "<br>2196 sql:=" & sql    
	'response.	write "<br><br><br>Paso" & sql
	'exit sub
	rs.Open sql,conexion

 %>

 
 <div class="w3-accordion w3-theme-l3">
 
         <% 
        if isnull(rs("Path")) then
           sPro= rs("Link").value 
        else
            sPro= "/" & rs("Path") & "/" & rs("Link").value 
        end if    
         spxx=sPro
            ed_CalPar 1,ed_iCla,1,"","","", ed_ifil,rs("Id_Menu").value, rs("Id_Menu").value
            if rs("Parametros").value<>"" then sPar=sPar & rs("Parametros").value
            sPro=spxx   
      
        iNiv= rs("idNivel").value
      
        i=0
        iDiv=0
        Do While NOT rs.EOF
  
               
              
               if rs("idNivel").value=0 Then ixMen1=rs("id_Menu").value
               if rs("idNivel").value=1 Then ixMen=rs("id_Menu").value

                if ed_ims="" then ed_Ims=0            
                ix=ed_iMs-rs("id_Menu").value
                
                sD="computer"
                sIcon=rs("Icon").value
                if isnull(sIcon) then sIcon=sD
                if rs("idNivel").value<1 Then
                    
                    %>
                    <% if i<>0 then %> 
                        </div> 
                         
                    <%    iDiv=1
                        end if 
                     if iNiv1=1 then%>
                        </div><%
                        iNiv1=0
                        iDiv=1
                     end if           
                                
                    
                     if rs("idNivel").value=0 then 
                        'sCla="w3-btn-block w3-theme-l1 w3-left-align w3-medium w3-show" 
                        sCla="w3-btn-block w3-theme-l1 w3-left-align w3-medium" 
                     else 
                        scla="w3-btn-block w3-theme-l4 w3-left-align w3-medium"
                     end if   
                    'if rs("Menu").value="#" then%>    
                        <button onclick="myFunction('Mnu<%=rs("Id_Menu").value%>')" class="<%=sCla %>">
                        <i class="material-icons w3-large w3-margin-righ"><%=sIcon%></i>
                        <i class=" w3-margin-right "></i><%=rs("Menu").value%>
                        </button>              
         
      

                        <%sCla="w3-accordion-content w3-container"
                     '   response.write "<br>3418 ed_imp:=" & ed_imp & " id_menu:=" & rs("Id_Menu").value
                        if isnumeric(ed_imp) then 
                            ix=ed_imp -rs("Id_Menu").value
                        end if    
                        if ix=0 then  sCla=sCla & " w3-show " %>
                        <div id="Mnu<%=rs("Id_Menu").value%>" class="<%=sCla %>">
                        <%iMen=rs("Id_Menu").value
                        iDiv=0
                   'else%>
                     
                    <!--a href="<%=sPar%>" title="<%=rs("txt_tips").value%>"><%=rs("Menu").value %></a--><%
                  '  end if    
                    
                     %>
                <%else
                
                
                 
                         if rs("Menu").value<>"#" then
                           
                           if rs("Link").value ="#" then
                            'if  rs("idNivel").value=1 then
                                if iNiv1=1 then%>
                                    </div><%
                                    iDiv=1
                                end if
                               %>
                                <%sCla="w3-btn-block w3-theme-l4 w3-left-align w3-medium"
                                ix=ed_ims -rs("Id_Menu").value
                                if ix=0 then  sCla=sCla & " w3-show " %>
                                <button onclick="myFunction('Mnu<%=rs("Id_Menu").value%>')" class="<%=sCla %>">
                                   <!--i class="material-icons w3-large w3-margin-righ"><%=sIcon%></i-->
                                   <i class=" w3-margin-right w3-small"></i><%=rs("Menu").value%>
                                </button>
                                <%sCla="w3-accordion-content w3-container"
                                  ix=ed_ims -rs("Id_Menu").value
                                if ix=0 then  sCla=sCla & " w3-show " %>
                                <div id="Mnu<%=rs("Id_Menu").value%>" class="<%=sCla %>">
                                <%iNiv1=1 
                                iDiv=0
                                iMen1=rs("Id_Menu").value%>
                            <%else                                       
                                if isnull(rs("Path")) then
                                    sPro= rs("Link").value 
                                else
                                    sPro= "/" & rs("Path") & "/" & rs("Link").value 
                                end if    
                                spxx=sPro                                
                                sx=spMen
                                spMen=rs("Txt_Tips")
                                ed_CalPar 1,ed_iCla,1,"","","", ed_ifil,ixMen1,ixMen
                                if rs("Parametros").value<>"" then sPar=sPar & rs("Parametros").value
                                sPro=spxx                    
                                spMen=sx %>    
                         
                                <a href="<%=sPar%>" title="<%=sPro & "-" & rs("txt_tips").value%>" class="w3-medium"><%=rs("Menu").value  %></a>
                             <% end if %>
                        <%else%>
                            <div class="w3-theme-l2" style="height:1px; "></div>
                         <% iDiv=0
                            end if    
                  
                   
                
                end if%>
                <% iNiv = rs("idNivel").value
            
            rs.MoveNext
            i=i+1 
        Loop
        
        rs.close
       
        if iDiv=0 then %>
           </div>
        <%
        end if    
         %>
        
   
</div>

<%
end Sub 

%>	




