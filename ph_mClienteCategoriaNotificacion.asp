<!DOCTYPE html>
<!-- ph_mClienteCategoriaNotificacion.asp - 15sep21 - 23ene22 -->
<html lang="es">	
<head>
	<title>| Notificaci&oacute;n |</title>
	<meta charset="UTF-8">
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="favicon.ico" rel="icon" type="image/x-icon">
	<link href="css/sweetalert.css" rel="stylesheet" type="text/css" />
	<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />	
	<link href="css/bootstrap-multiselect-0915.css" rel="stylesheet" >
	<link href="notificacion/css/notificacion.css"  rel="stylesheet" type="text/css" />
</head>

<body topmargin="0">
	
	<!--#include file="estiloscss.asp"-->
	<!--#include file="meta.asp"-->
	<!--#include file="encabezado.asp"-->
	<!--#include file="nn_subN.asp"-->
	<!--#include file="in_DataEN.asp"-->
	
	<% 
		' 29dic20 - 29ene21
		Apertura
		' Parámetros del Manteniemiento
		LeePar
		
		if ed_iPas<>4 then 
			Encabezado
		end if    	
		'	
	%>

	<div class="container-fluid" id="grad1" style="visibility:hidden;"> 
		<br>	
	
		<div class="row">
	
			<div class="col-sm-3">
				<div class="form-group">				
					<label><i class="fas fa-user"></i>&nbsp;Cliente:</label>	
					<select class="form-control" title="Seleccione el Cliente" name="cboCliente" id="cboCliente" />						
					</select>
				</div>
			</div>
												
			<div class="col-sm-3">
				<div class="form-group">				
					<label><i class="fas fa-clipboard-check"></i>&nbsp;Tipo Reporte:</label>
					<select class="form-control" title="Seleccionar Tipo Reporte" name="cboTreporte" id="cboTreporte" />
						<option value="0" selected disabled> -- Seleccione -- </option>
						<option value="General">General</option>
						<option value="Detallado">Detallado</option>
					</select>
				</div>
			</div>
			
			<div class="col-sm-3">
				<div class="form-group">				
					<label><i class="fas fa-print"></i>&nbsp;Reporte:</label>
					<select class="form-control" title="Seleccionar Reporte" name="cboReporte" id="cboReporte" />						
						' <option value="0" selected disabled> -- Seleccione -- </option>
                        <option value="Semanal">Semanal</option>
                        <option value="Mensual">Mensual</option>
					</select>
				</div>
			</div>									
													
			<div class="col-sm-3">
				<div class="form-group">				
					<label><i class="fas fa-shapes"></i>&nbsp;Categoria:</label>
					<select class="form-control" title="Seleccione Categoria" name="cboCategoria" id="cboCategoria" multiple="multiple" />						
					</select>
				</div>
			</div>							
			
		</div>
		<!---->
		<div class="row">		
		
			<div class="col-sm-3">
				<div class="form-group">				
					<label id="lblTiempo"><i class="fas fa-calendar-check"></i>&nbsp;Periodo:</label>
					<select class="form-control" title="Seleccione una opcion" name="cboTiempo" id="cboTiempo" /></select>					
				</div>				
			</div>
				
			<div class="col-sm-3" style="margin-top: 2px;">				
			
				<div class="form-group">
				
					<div class="col-sm-6">
						<div class="form-group">
							<label for="usr"></label>
							<button id="btnReset" class="btn btn-block btn-sm btn-danger" title="Borrar Pantalla" ><i class="fas fa-times"></i>&nbsp;Reset</button>
						</div>
					</div>
										
					<div class="col-sm-6">
						<div class="form-group">
							<label for="usr"></label>
							<button id="btnEnviar" class="btn btn-block btn-sm btn-success" title="Enviar Notificaci&oacute;n" type="submit"><i class="fas fa-paper-plane"></i>&nbsp;Procesar</button>
						</div>
					</div>
					
				</div>				
				
			</div>
						
		</div>			
		            								
	</div>        
	<hr>
	
	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;">
		<span ><img src="images/ajax-loader7.gif"><strong>&nbsp;Espere un momento..!</strong></span>
	</div>
	
	<div class="container-fluid text-center text-primary" id="envioEmail" style="display:none;">	
		<span ><img src="images/email/email9.gif" class="img-responsive center-block" alt="Correo enviado..!" width="500" height="400"></span>
	</div>
					
			
	<%conexion.close%>
	
</body>
</html>
<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/jquery-3.1.1.min.js"></script>
<script src="js/sweetalert.min.js"></script>
<script src="js/bootstrap.min.js"></script>
<script src="matconvivencia/js/url.js"></script>
<script src="notificacion/js/funcionesV1.js"></script>
<script src="js/bootstrap-multiselect-0915.js"></script>

<script>
	
	$(document).ready(function() {
		sessionStorage.clear();
		url();
	 	$("#cboSemana").multiselect({ buttonWidth:'450px', disableIfEmpty: true, });		
		$("#cboCategoria").multiselect({ buttonWidth:'450px', disableIfEmpty: true, });
		$("#cboTiempo").multiselect({ buttonWidth:'450px', disableIfEmpty: true,  });
		$('#cargando').css('display', 'block');
		$(function() {
			fillData();
			document.getElementById('grad1').style.visibility="visible";
		});		
	});
	
</script>

<script>
var scripts = document.getElementsByTagName('script');
//console.log(scripts);
var toRefreshs = ['funcionesV1.js']; // list of js to be refresh
var key = Math.floor((Math.random() * 10) + 1); // change this key every time you want force a refresh
for (var i = 0; i < scripts.length; i++) {
    for (var j = 0; j < toRefreshs.length; j++) {
        if (scripts[i].src && (scripts[i].src.indexOf(toRefreshs[j]) > -1)) {
            new_src = scripts[i].src.replace(toRefreshs[j], toRefreshs[j] + 'k=' + key);
            scripts[i].src = new_src; // change src in order to refresh js
        }
    }
}
</script>

