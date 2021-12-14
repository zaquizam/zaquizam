i<%
'Variables Globales
Session.TimeOut = 60
Server.ScriptTimeout = 5000
Dim sActivar_planificacion
Dim sActivar_tiendaspendientes
Dim sActivar_tiendasrealizadas
Dim sTitulo_modulo
Response.CharSet = "ISO-8859-1"
Response.addHeader "pragma", "no-cache"
Response.CacheControl = "Private"
Response.Expires = 0
if Session("TituloApp") = "" then
	Session("TituloApp") = "| Atenas - Acceso |"
end if
Dim borra_cookies			
For Each borra_cookies In Request.Cookies
	Response.Cookies(borra_cookies).Expires =#May 25, 2009#						
Next
%>
<!DOCTYPE html>
<html Lang="es">
<head>
	<!-- Creado: 18abr17 - Actualizado: 26jun17 -->
	<title>| Atenas |</title>       
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<!-- Los 3 primeros metas siempre deben ir en el HEAD antes que nada -->			
	<link href="modal.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="css/bootstrap.min.css" rel="stylesheet" media="screen">
	<link href="css/bootstrap-theme.min.css" rel="stylesheet" media="screen">         
	<link href="css/login.css" rel="stylesheet" type="text/css" media="screen">
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 	
	<!--<script src="https://maps.googleapis.com/maps/api/js?v=3.exp&signed_in=true&libraries=places"></script>-->
	<script type="text/javascript" src="js/jquery-1.11.3-jquery.min.js"></script>
	<script type="text/javascript" src="js/validation.min.js"></script>
	<script type="text/javascript" src="js/loginscript.js"></script>						
</head>
<!--<body onload="loadLocation();">-->
<body>
	<div id="DetallePauta">

		<div id="1" class="modalmask">
				<div class="modalbox rotate">
				<a href="#close" title="Cerrar" class="close">x</a>
				<fieldset class="det">
					<!--<legend class="det">Recuperar Contase&ntilde;a</legend>-->
					<h5>
						<div style="width:800px; align=right ; ">
							<p align="right">
								Fecha:
								<input type="text" name="Fecha" id="idFecha" align="right" size=25 value ="<%=now()%>" readonly>
							</p>
						</div>
						<h5 class="text-center"><strong>Datos</strong></h5>
						<p align="right">
							<div style="width:850px; padding:3px; align=right ;font-size: 10pt">
								<div style="width:150px; float:left;">
									<!--Correo-->
										Correo de Acceso:
								</div>
								<div id="DivDescripcion" style="width:140px; float:left;">
									<input type="text" name="CorreoSolicitar" id="CorreoSolicitar" align="right" size=50 style="text-align:left">
								</div>
							</div> 
						</p>
						<br>
						<p align="right">
							<center>
							<!--BOTON--> 
							<img src="images/si.png"  style="margin-left:0px;" alt="Agregar" width="20px" onclick="SolicitarClave()"/>
							</center>
						</p>
						<br>
						<br>
						<center>
							<span id="loader02"></span>
						</center>	
					</h5>
				</fieldset>
			</div> <!--modalbox rotate-->
		</div> <!--modal5-->

	</div> <!--DetallePauta1-->

	<div class="signin-form">	
		<form class="form-signin" Method="post" id="login-form" autocomplete="off">									
			<div class="container-fluid">
				<div class="row-fluid">
					<img src="images/logo/LogoAtenasNew01.jpeg"  alt="Logo Empresarial" class="img-responsive center-block" >					
				</div>				
			</div>									
			<hr /> 		
			<div id="error" style="font-size: 10pt">
				<!-- Mostrar posibles errores ! -->
			</div>						
			<div class="form-group">
				<div class="input-group">
					<span class="input-group-addon"><i class="glyphicon glyphicon-envelope"></i></span>
					<input type="email" class="form-control" placeholder="Email" name="email" id="email" style="text-align:left"/>
					<span id="check-e"></span>				
				</div>			
			</div>
			<div class="form-group">
				<div class="input-group">
					<span class="input-group-addon"><i class="glyphicon glyphicon-lock"></i></span>
					<input type="password" class="form-control" placeholder="Password" name="password" id="password" autocomplete="off"  style="text-align:left"/>				
				</div>			
			</div>
			<hr />			
			<div class="form-group">
					<button type="submit" class="btn btn-default" name="btn-login" id="btn-login">
					<span class="glyphicon glyphicon-log-in"></span> &nbsp; Ingresar
				</button> 
			</div>
			<div style="width:150px; float:left;font-size: 10pt">
				<!--<a href="#1" title="Recuperar Contase&ntilde;a">Recuperar Contase&ntilde;a</a>-->
			</div>
	   </form>					
	</div> <!-- class="signin-form" -->
	<script src="js/bootstrap.min.js"></script>
	<script>	
		
		//**Inicio Solitar Clave por Correo 
		function SolicitarClave() {
			var sx = document.getElementById("CorreoSolicitar").value
			if (sx == "") {
				alert("Debe Incluir un correo");
				return;
			} else {
			}
			//alert(document.getElementById("CorreoSolicitar").value);

			var sVar="";
			sVar = document.getElementById("CorreoSolicitar").value;
			//document.getElementById("Programa").value =  "SolicitarClave?" + sVar;
			//alert(sVar);
			$.ajax({
				url:'Sys_gSolicitarClave.asp?cor='+sVar,
				beforeSend: function(objeto){
					$('#loader02').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				},
				success:function(data){
					//debugger;
					console.log(data);
					$('#loader02').html('');
					//alert(data)
					//alert(data)
					if (data == "NO") {
						alert("Correo No Existe. Verifique la informacion Suministrada");
						return;
					}					
					if (data == "") {
						alert("Correo No Enviado. Intentelo más tarde");
						return;
					}
					if (data == "SI") {
						alert("Correo Enviado");
						return;
					}else {
						alert("Correo No Enviado. Intentelo más tarde");
						return;
					}
					
				}
			})
			
		}
		//**Fin Solitar Clave por Correo 
		
		
		function loadLocation () {						
			if (navigator.geolocation) { /* Si el navegador tiene geolocalizacion */
				//navigator.geolocation.getCurrentPosition(coordenadas, errores);
				navigator.geolocation.getCurrentPosition(coordenadas, MostrarError,{timeout:6000});
			}else{
				alert('Aviso! El navegador no soporta geolocalizacion. Actualicelo...!');
			}						
		}
		
		function coordenadas (position) {
			debugger;		
			var times = position.timestamp;
			var altitud = position.coords.altitude;	
			var lng = position.coords.longitude;
			var lat = position.coords.latitude;
			var exactitud = position.coords.accuracy;	
			//var link = "https://www.google.com/maps/place/8°57'00.3N+75°2652.2W/@"+lat+","+lng+",19z"
			document.cookie = "lat="+position.coords.latitude; 	/*Guardamos nuestra latitud*/		
			document.cookie = "lng="+position.coords.longitude;	/*Guardamos nuestra longitud*/		
			sessionStorage.lat= lat;
			sessionStorage.lng= lng;			
			//var div = document.getElementById("ubicacion");
			//div.innerHTML = "Timestamp: " + times + "<br>Latitud: " + latitud + "<br>Longitud: " + longitud + "<br>Altura en metros: " + altitud + "<br>Exactitud: " + exactitud;}				 
		}
		
		function MostrarError (error) {
			debugger;
			//alert(error.code);		 
			/*Controlamos los posibles errores */
			if (error.code == 0) {
				document.cookie = "lat=error_0"; 	
				document.cookie = "lng=error_0";	
				alert("Error..! de Geolocalización desconocido..!");
			}
			if (error.code == 1) {
				document.cookie = "lat=error_1"; 	
				document.cookie = "lng=error_1";	
				alert("Permiso de ubicacion negado..!");			  
			}
			if (error.code == 2) {
				document.cookie = "lat=error_2"; 	
				document.cookie = "lng=error_2";	
				alert("Hay un problema para conseguir la posición del dispositivo!");
			}
			if (error.code == 3) {
				document.cookie = "lat=error_3"; 	
				document.cookie = "lng=error_3";	
				alert("La aplicación agotó el tiempo tratando de obtener la posición del dispositivo!");
			}		 
		} 	
		
		function refrescarUbicacion() {	
			navigator.geolocation.watchPosition(coordenadas);
		}	
		//
		function showError(error) {
		  switch(error.code)
			{
				case error.PERMISSION_DENIED:
				  x.innerHTML="Negada la peticion de Geolocalización en el navegador."
				  break;
				case error.POSITION_UNAVAILABLE:
				  x.innerHTML="La información de la localización no esta disponible."
				  break;
				case error.TIMEOUT:
				  x.innerHTML="El tiempo de petición ha expirado."
				  break;
				case error.UNKNOWN_ERROR:
				  x.innerHTML="Ha ocurrido un error desconocido."
				  break;
			}
		}	
	</script>
</body>
</html>