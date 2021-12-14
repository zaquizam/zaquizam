<!DOCTYPE html>
<!-- ph_mClienteCategoriaNotificacion.asp - 15sep21 - 01oct21 -->
<html>
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
	<link href="matconvivencia/css/bootstrap-multiselect-0915.css" rel="stylesheet" type="text/css"/>
	<link href="matconvivencia/css/convivencia2.css"  rel="stylesheet" type="text/css" />
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

	<div class="container-fluid" id="grad1"> 
		<br>	
	
		<div class="form-group">
	
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
					<select class="form-control" title="Seleccionar Tipo Reporte" name="cboTreporte" id="cboTreporte" >
						<option value="0" selected disabled> -- Seleccione -- </option>
						<option value="General">General</option>
						<option value="Detallado">Detallado</option>
					</select>
				</div>
			</div>
			
			<div class="col-sm-3">
				<div class="form-group">				
					<label><i class="fas fa-print"></i>&nbsp;Reporte:</label>
					<select class="form-control" title="Seleccionar Reporte" name="cboReporte" id="cboReporte" >						
						' <option value="0" selected disableds> -- Seleccione -- </option>
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
		
		<div class="form-group">		
		
			<div class="col-sm-3">
				<div class="form-group">				
					<label id="lblTiempo"><i class="fas fa-calendar-check"></i>&nbsp;Periodo:</label>
					<select class="form-control" title="Seleccione una opcion" name="cboTiempo" id="cboTiempo" />						
					</select>					
				</div>				
			</div>
							
			<div class="col-sm-3">				
			
				<div class="form-group">
				
					<div class="col-sm-6">
						<label for="usr"></label>
						<button id="btnReset"  title="Borrar Pantalla" class="btn btn-block btn-sm btn-danger"><i class="fas fa-recycle"></i>&nbsp;Reset</button>
					</div>
										
					<div class="col-sm-6">
						<label for="usr"></label>
						<button id="btnEnviar" title="Enviar Notificaci&oacute" type="submit" class="btn btn-block btn-sm btn-success"><i class="fas fa-paper-plane"></i>&nbsp;Procesar</button>
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
<script src="notificacion/js/funciones.js"></script>
<!-- MultiSelect CSS & JS library -->
<script src="matconvivencia/js/bootstrap-multiselect-0915.js"></script>

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
		});		
	});
	
</script>

