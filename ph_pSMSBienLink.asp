<!DOCTYPE HTML>
<html >
<head>
	<title>SMS Bien/Link</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="mensaje.css" rel="stylesheet" type="text/css" media="screen" />
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<script type="text/javascript" src="js/sweetalert.min.js"></script>
	<script type="text/javascript" src="js/jquery-1.12.4.min.js"></script>

</head>
<body topmargin="0">
<!--#include file="estiloscss.asp"-->
<!--#include file="meta.asp"-->
<!--#include file="encabezado.asp"-->
<!--#include file="nn_subN.asp"-->
<!--#include file="in_DataEN.asp"-->

<%

  
'==========================================================================================
' Variables y Constantes
'==========================================================================================


    Apertura
	   
	dim envMensaje
	dim envCelular
%>
<script type="text/javascript">
	function enviarsms(celular) 
	{
		//debugger;
		var celularenv = "'58" + celular + "'";
		var mensaje = "'" + document.getElementById("Mensaje").value + "'";
		//alert(celularenv);
		//alert(mensaje);
		//return;
		$.ajax({
            url: 'https://site.albertext.com/api/messages/save-message',
            type: 'post',
            dataType: 'json',
            contentType: 'application/json',
            success: function (data) {
                console.log('Success', data);
            },
            data: JSON.stringify({
			"user":"Atenas",
			"token":"XZM2tfgOW0tscbRETqh91H7TKcT19NES",
			"phone":celularenv,
			"text":mensaje
		  })
        });	
		//"phone":"584168254124",
		//	"text":"Mensaje Enviado desde la Web de Operaciones de Atenas a la 2:49pm, Por favor avisame si te llego"
	}
	function alerta(total) 
	{
		swal("Se Enviaron " + total + " SMS","Enviado","success");
		//window.open("?edpas=1&smenu=?x=1&smenu=Envio%20SMS%20Bienvenida%20y%20Link","_parent");
	}
</script>
<%

Sub Combos
 
	'response.write "<br>372 Combo1:=" & ed_sPar(1,0)
	'response.write " Combo2:=" & ed_sPar(2,0)
	'response.write " Combo3:=" & ed_sPar(3,0)
	'response.write " Combo3:=" & ed_sPar(4,0)
	'response.write " Combo3:=" & ed_sPar(5,0)
    ed_iCombo = 2 
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_SMS, "
	sql = sql & " SMS "
	sql = sql & " FROM PH_SMS "
	sql = sql & " WHERE "
	sql = sql & " Id_SMS < 3 "
	sql = sql & " Order By "
	sql = sql & " Id_SMS "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="SMS"
    ed_sCombo(1,1)=sql 
    ed_sCombo(1,2)="Seleccionar"


	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_SiNo, "
	sql = sql & " SiNo "
	sql = sql & " FROM PH_SiNo "
	sql = sql & " Order By "
	sql = sql & " Id_SiNo "
	'response.write "<br>372 Combo2:=" & sql
    ed_sCombo(2,0)="Enviar"
    ed_sCombo(2,1)=sql 
    ed_sCombo(2,2)="Seleccionar"
	
End Sub

   
'==========================================================================================
' Parámetros del Manteniemiento
'==========================================================================================
    LeePar
  
    
    if ed_iPas<>4 then 
        Encabezado
    end if    

	'response.write "llego1"
	'response.end
	'ParDat
%>
		
	<br>
	<br>
	<br>
	<div style="width:98%"></div></center>
<%
	Combos
	'response.write "paso"

%>

	<table border="0" align="right">
		<tr>
			<td>
				<%
				ed_vCombo
				%>
			</td>
		</tr>
	</table>
	</br>
	<h3>OBSERVACION = Por los Momentos se enviaran 50 SMS por proceso, espere 5 minutos antes de hacer el siguiente envio, no Cierre la pantalla porque corta el envio</h3>
	</br>
	</br>

<%
'==========================================================================================
' Variables y Constantes
'==========================================================================================

	if ed_sPar(1,0) <> "Seleccionar" then
		dim gDatosSol2
		dim rsx2
		set rsx2 = CreateObject("ADODB.Recordset")
		rsx2.CursorType = adOpenKeyset 
		rsx2.LockType = 2 'adLockOptimistic 

		sql = ""
		sql = sql & " SELECT "
		sql = sql & " Mensaje "
		sql = sql & " SMS "
		sql = sql & " FROM PH_SMS "
		sql = sql & " WHERE "
		sql = sql & " Id_SMS = " & ed_sPar(1,0)
		'response.write "<br>36 sql:=" & sql
		'response.end
		rsx2.Open sql ,conexion
		iExiste = 0
		if rsx2.eof then
		else
			gDatosSol2 = rsx2.GetRows
			rsx2.close
			envMensaje = gDatosSol2(0,0)
			%>
			<div id="" align=center>
				<h2>Mensaje = <%=envMensaje%></h2>
			</div>
			<%
		end if
	end if
	if ed_sPar(2,0) = "Seleccionar" then 
		isw = 0
	else
		isw = ed_sPar(2,0)
	end if
	

	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
    sql = sql & " SELECT "
	sql = sql & " PH_Panelistas.Id_Hogar, "
	sql = sql & " PH_PanelHogar.CodigoHogar, "
	sql = sql & " PH_Panelistas.Id_Panelista, "
	sql = sql & " PH_Panelistas.Nombre1, "
	sql = sql & " PH_Panelistas.Apellido1, "
	sql = sql & " PH_Panelistas.Celular "
	sql = sql & " FROM PH_Panelistas INNER JOIN PH_PanelHogar ON PH_Panelistas.Id_Hogar = PH_PanelHogar.Id_PanelHogar "
	sql = sql & " WHERE "
	'sql = sql & " Ind_SMS_Link = 1 "
	sql = sql & " PH_Panelistas.Id_Hogar > 1 "
	sql = sql & " AND PH_Panelistas.ResponsablePanel = 1 "
	sql = sql & " AND PH_PanelHogar.Ind_Activo = 1 "
	if ed_sPar(1,0) <> "Seleccionar" then
		if cint(ed_sPar(1,0)) = 1 then
			sql = sql & " AND (Ind_SMS_Bienvenida is null "
			sql = sql & " OR Ind_SMS_Bienvenida = 0 ) "
		else
			sql = sql & " AND (Ind_SMS_Link is null "
			sql = sql & " OR Ind_SMS_Link = 0 ) "
		end if
	else
		sql = sql & " AND PH_Panelistas.Id_Hogar is null "
	end if
	'response.write "<br>36 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	iExiste = 0
	if rsx1.eof then
		%>
		<div id="DivBuscarPanelistas">
			<h3>Personas Seleccionadas = 0</h3>
		</div>
		<%
	else
		gDatosSol = rsx1.GetRows
		rsx1.close
		%>
		<div id="DivBuscarPanelistas">
			<h3>Personas Seleccionadas = <%=ubound(gDatosSol,2)+1%></h3>
			<div class="ex1">
				<table class="w3-table w3-striped w3-bordered w3-card-4 w3-small" style="width:1000px; margin-left:auto; margin-right:auto;margin-top:10px ">
					<thead>
						<tr class="w3-blue">
							<th>Id Hogar</th>
							<th>Hogar</th>
							<th>Id Panelista</th>
							<th>Primer Nombre</th>
							<th>Primer Apellido</th>
							<th>Celular</th>
							<th>Observacion</th>
						</tr>
					</thead>
					<% 
					iTotalErr = 0
					for iReg = 0 to ubound(gDatosSol,2)
						sObservacion = ""
						envCelular = trim(gDatosSol(5,iReg))
						if Len(envCelular) <> 10 or mid(envCelular,1,1) = "0" then
							sObservacion = sObservacion & "**Celular Errado**"
							iTotalErr = iTotalErr + 1
						end if
						Response.write "<tr>"
							for ib = 0 to 5
								Response.write "<td>" & gDatosSol(ib,iReg) & "</td>"
							next
							Response.write  "<td>" & sObservacion & "</td>"
						Response.write "</tr>"
					next
					%>
				</table>
			</div>
			<h3>Celulares con Errores = <%=iTotalErr%></h3>
		</div>
		<%
	end if

	'***********ENVIO
	if ed_sPar(1,0) <> "Seleccionar" and isw = 1 then
		'response.write "<br> Enviar" 
		Total = 0
		%>
		<input type="text" name="Mensaje" id="Mensaje" value="<%=envMensaje%>" size=160 style="text-align:left; ">
		<input type="text" name="Celular" id="Celular" value="" size=20 style="text-align:left; ">
		<%
		for iReg = 0 to 30 'ubound(gDatosSol,2)
			'envCelular = "4241656449"
			'envCelular = trim(gDatosSol(5,iReg))
			'envMensaje = "(" & envCelular & ")" & envMensaje 
			envCelular = trim(gDatosSol(5,iReg))
			if Len(envCelular) <> 10 or mid(envCelular,1,1) = "0" then

			else	
				Total = Total + 1
				%>
				<script>enviarsms(<%=envCelular%>)</script>
				<%
				dim rsx3
				set rsx3 = CreateObject("ADODB.Recordset")
				rsx3.CursorType = 0
				rsx3.LockType = 3
				sql = ""
				sql = sql & " Select * from PH_SMS_Hist "
				'response.write "<br>220 sql:=" & sql
				'response.end
				
				rsx3.Open sql ,conexion
				rsx3.AddNew()
				rsx3("Id_SMS") = ed_sPar(1,0)
				rsx3("Id_PanelHogar") = trim(gDatosSol(0,iReg))
				rsx3("Id_Panelista") = 	trim(gDatosSol(2,iReg))
				rsx3("Celular") = envCelular
				rsx3("Fecha") = now()
				rsx3.Update
				rsx3.Close 

				sql = ""
				sql = sql & " Select * from PH_PanelHogar "
				sql = sql & " Where Id_PanelHogar = " & trim(gDatosSol(0,iReg))
				'response.write "<br>220 sql:=" & sql
				'response.end
				rsx3.Open sql ,conexion
				if cint(ed_sPar(1,0)) = 1 then
					rsx3("Ind_SMS_Bienvenida") = 1
				else
					rsx3("Ind_SMS_Link") = 1
				end if
				rsx3.Update
				rsx3.Close 
				 
			end if
		next 
		%>
		<script>alerta(<%=Total%>)</script> 
		<%
	end if
	'response.end 
%>
<style>


</style>	

    <%conexion.close%>
	
</body>
</html>