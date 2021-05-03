<!Doctype html>
<!-- PH_Cte_HomePantryConvivencia - 09abr21 - 11abr21 -->
<html >
<head>
	<title>| Convivencia |</title>
	<meta charset="UTF-8">
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="favicon.ico" rel="icon" type="image/x-icon">
	<link href="css/sweetalert.css"  rel="stylesheet" type="text/css" />
	<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />	
	<link href="css/convivencia.css"  rel="stylesheet" type="text/css" />	
		
</head>
<body topmargin="0">
	
	<!--#include file="estiloscss.asp"-->
	<!--#include file="meta.asp"-->
	<!--#include file="encabezado.asp"-->
	<!--#include file="nn_subN.asp"-->
	<!--#include file="in_DataEN.asp"-->
		
	<% 
		' 09abr21 - 11abr21
		Apertura		
		LeePar		
		if ed_iPas<>4 then 
			Encabezado
		end if    	
		'	
	%>

	<div class="container-fluid" id="grad1" >  
	
		<div class="form-group" >
	
			<div class="col-sm-3">
				<div class="form-group"  >				
					<label>Seleccione Mes:</label>
					<select class="form-control input-sm" title="Seleccionar Semana" name="cboProcesarFecha" id="cboProcesarFecha" onchange="procesarFecha();"  />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
					
		</div>		
		            								
	</div>        
		
	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;">
		<span ><img src="images/ajax-loader7.gif"><strong>&nbsp;Espere, Procesando...!</strong></span>
	</div>
	
	<div class="container-fluid" id="detallesMaestro" >
	
		<div class="form-group">
	
			<div class="col-sm-6">
				<div class="form-group" id="detallesPaso1" style="display:none;">			
					<!-- Code by w3codegenerator.com -->
					<div class="table-responsive" id="Paso1">
						<!---->									 					
					</div>		
				</div>
			</div>
			
			
			<div class="col-sm-6">
				<div class="form-group" id="detallesPaso2" style="display:none;">			
					<!-- Code by w3codegenerator.com -->
					<div class="table-responsive" id="Paso2">
						<!---->									 					
					</div>		
				</div>
			</div>
				
			<br>
			<div class="col-sm-12">
				<div class="col-sm-6">
					<div class="form-group" id="detallesPaso3" style="display:none;">						
						<table class="table table-striped text-center" style=" margin: auto; ">
							<thead>
								<tr>
									<th class="text-center text-danger"><i class='fas fa-check-double'></i><strong>&nbsp;PORCENTAJES DE HOGARES QUE COMPRARON AL MENOS UNA BEBIDA REFRESCANTE ENVASADA&nbsp;</strong></th>
								</tr>
							</thead>
							<tbody>
								<tr>
									<th class="text-center text-primary"><h4><strong><span id="Paso3"></span></strong></h4></th>
								</tr>
							</tbody>
						</table>						
					</div>		
				</div>		
				<br>
				<div class="col-sm-6">
					<div class="form-group" id="detallesPaso4" style="display:none;">			
						<!-- Code by w3codegenerator.com -->
						<div class="table-responsive" id="Paso4">
							<!---->									 					
						</div>		
					</div>
				</div>
			</div>
			
		</div>
		
	</div>		
	
	<hr>	
		
	<%conexion.close%>
	
</body>
</html>
<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/jquery-3.1.1.min.js"></script>
<script src="js/sweetalert.min.js"></script>
<script src="js/bootstrap.min.js"></script>
<script src="valconveniencia/funcionesV3.js"></script>

<script>
	
	$(document).ready(function() {				
		$(function() {
			buscarFechas();
		});					
	});	
	
</script>

