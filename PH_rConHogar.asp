<!Doctype html>
<!-- ph_rConHogar - 26feb21 -  -->
<html >
<head>
	<title>Consumos x Hogar</title>
	<meta charset="UTF-8">
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="favicon.ico" rel="icon" type="image/x-icon">
	<link href="css/sweetalert.css" rel="stylesheet" type="text/css" />
	<link href="css/factura.css" rel="stylesheet" type="text/css" />	
	<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />	
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
	
		<div class="form-group">
	
			<div class="col-sm-3">
				<div class="form-group">				
					<label>Seleccione Semana:</label><span id="loader"></span>	
					<select class="form-control input-sm" title="Seleccionar Semana" name="cboSemana" id="cboSemana" onchange="buscarArea();"  />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
												
			<div class="col-sm-2">
				<div class="form-group">				
					<label>Seleccione Area:</label>
					<select class="form-control input-sm" title="Seleccionar Area" name="cboArea" id="cboArea" onchange="buscarEstado();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
			
			<div class="col-sm-2">
				<div class="form-group">				
					<label>Seleccione Estado:</label>
					<select class="form-control input-sm" title="Seleccionar Estado" name="cboEstado" id="cboEstado" onchange="buscarHogar();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
						
			<div class="col-sm-2">
				<div class="form-group">				
					<label>Seleccione Hogar:</label>
					<select class="form-control input-sm" title="Seleccionar Hogar" name="cboHogar" id="cboHogar" onchange="buscarTipoConsumo();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
			
			<div class="col-sm-3">
				<div class="form-group">				
					<label>Seleccione Tipo Consumo:</label>
					<select class="form-control input-sm" title="Seleccionar Area" name="cboTipoConsumo" id="cboTipoConsumo" onchange="buscarTotalDiaSemana();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>		
						
		</div>
		
		<!-- 0 -->
		
		<div class="form-group">
																		
			<div class="col-sm-4">
				<div class="form-group">				
					<label>Seleccione Dia Semana:</label>
					<select class="form-control input-sm" title="Seleccionar Dia" name="cboTotalxDiaSemana" id="cboTotalxDiaSemana" onchange="buscarDetalleDiaSemana();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
	
			<div class="col-sm-4">
				<div class="form-group">				
					<label>Seleccione Fecha:</label>
					<select class="form-control input-sm" title="Seleccionar Consumo a Validar" name="cboDetallexDiaSemana" id="cboDetallexDiaSemana" onchange="buscarDetalleConsumoxDia();" />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
						
			<div class="col-sm-4">				
			
				<div class="form-group">
					<label for="usr">Reset</label>
					<button id="borrar"  title="Borrar Pantalla" type="submit" class="btn btn-block btn-xs btn-danger" onclick="Reset();"><i class="fas fa-recycle fa-2x"></i></button>						
				</div>				
				
			</div>
																		
		</div>
			
		<!-- TABLA: RESUMEN -->
		
		<div class="form-group"> 
			<table class="table table-responsive">			
				<thead>
					<tr>					
						<th>Alta del Hogar:&nbsp;<span class="label label-warning" id="altaHogar"></span></th>
						<th>Responsable Hogar:&nbsp;<span class="label label-warning" id="responsableHogar"></span></th>
						<th>Celular:&nbsp;<span class="label label-warning" id="celularHogar"></span></th>
					</tr>
				</thead>			
			</table>		 
		</div>
            								
	</div>        
	<hr>
	
	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;">
		<span ><img src="images/ajax-loader7.gif"><strong>&nbsp;Espere, Procesando...!</strong></span>
	</div>
		
	<div class="container-fluid" id="DetalleFactura" style="display:none;">
							
		<div class="form-group-row">
		
			<div class="col-sm-3">
			
				<h4 class="text-danger"><strong><i class="fas fa-file-image"></i>&nbsp;Detalle Factura:</strong></h4>
				<input type="hidden" id="tieneFactura" name="tieneFactura" value="0" />									
								
				<fieldset> 					
					
					<div class="my-funky-img-box">
					   <!--/.col-bdr.left-border-col-->
					   <div class="col-bdr left-border-col"></div>				   
					   <!--/.background-col-->					   
					   <div class="background-col">
							<img class="img-thumbnail" id="imgfactura" name="imgfactura" src="images/loader/cargador1.gif" style="width:480px; height:640px;" />
					   </div>
					   <!--/.col-bdr.left-border-col-->
					   <div class="col-bdr right-border-col"></div>				   
					</div>
					
				</fieldset> 
						
				<!---->
				<br>				
				<div class="form-group">
					<label class="control-label col-sm-3">Canal:</label>
					<div class="col-sm-9">
						<select class="form-control input-sm" name="cboCanal" id="cboCanal" readonly >
							<option value="" selected disabled >-- Seleccione -- </option>							
						</select>					  
					</div>
					<div class="error" id="canalErr"></div>
				</div>
				<br>		
				
				<div class="form-group">
					<label class="control-label col-sm-3">Cadena:</label>
					<div class="col-sm-9">
						<select class="form-control input-sm" name="cboCadena" id="cboCadena" readonly >							
							<option value="" selected disabled >-- Seleccione -- </option>							
						</select>			
					</div>
					<div class="error" id="cadenaErr"></div>	
				</div>
								
				<div class="form-group">
					<strong><p class="text-danger" id="tienefactura"></p></strong>
				</div>
				
				<div class="form-group">
					<label class="control-label col-sm-9">Total Productos Comprados:</label>
					<div class="col-sm-3">
						<input type="number" class="form-control input-sm text-right" id="totalProductos" name="totalProductos" placeholder="Cantidad Productos Comprados" readonly />					
					</div>
					<div class="text-right" id="totalProductosErr"></div>	
				</div>
			    	
				<div class="form-group">
					<label class="control-label">Tipo Moneda:</label>
					<div class="">
						<select class="form-control input-sm" name="MonedaPagoFactura" id="MonedaPagoFactura" disabled>							
							<option value="0" selected disabled >-- Seleccione -- </option>							
						</select>			
					</div>
					<div class="error" id="monedapagofacturaErr"></div>	
				</div>
				
				<div class="form-group">
					<label class="control-label"><strong>Monto Total Factura:</strong></label>
					<div>
						<input type="text" class="form-control input-sm text-right" id="totalFactura" name="totalFactura" placeholder="Monto total de la Factura" readonly />					
					</div>
					<div class="text-right" id="totalfacturaErr"></div>	
				</div>								
										
										
			</div>
			
			<!-- TABLA DE PRODUCTOS REGISTRADOS -->
			<div class="col-sm-9">					
			
				<h4 class="text-danger"><strong>Detalle Productos: </strong></h4>
						
				<div class="form-group"  id="tabla-resultados">
					<!-- // ** // -->
					<!-- Matriz de Datos Resultados -->
					<!-- // ** // -->										
				</div>	
				
				<!-- PROMEDIOS SEMANALES X TIPO DE PRODUCTO -->
				<div class="form-group">					
			
					<h4><i class="fas fa-tachometer-alt"></i><strong>&nbsp;Indicador por Semana:</strong></h4>
								
					<div class="table-responsive" id="tabla-resumen">
						<!-- // ** // -->
						<!-- Matriz de Datos Resultados -->
						<!-- // ** // -->										
					</div>		
					
				</div>
				<!-- ./PROMEDIOS SEMANALES X TIPO DE PRODUCTO -->	
				
				
			</div>
			<!-- ./ TABLA DE PRODUCTOS REGISTRADOS -->
										
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
<!--<script src="mostrarconsumo/autoNumeric-1.9.18.js"></script>-->
<script src="mostrarconsumo/utilitariosV1.js"></script>
<script src="mostrarconsumo/crudV1.js"></script>
<script src="mostrarconsumo/funcionesV1.js"></script>
<script src="mostrarconsumo/funResueltoV1.js"></script>

<script>
	
	$(document).ready(function() {
							
		$(function() {
			buscarCanal();
		});
		$(function() {			
			buscarCadena(0);
		});		
		$(function() {
			buscarCategoria();
		});						
		$(function() {
			buscarMonedaPagoFactura();
		});	
		
		/*		
		$(function() {
			 buscarTipoInvestigacion();
		 });				
		*/
		
		$(function() {
			buscarSemanas();
		});
		
		sessionStorage.setItem('validado',false );
		sessionStorage.setItem('investigado', false );
		sessionStorage.setItem('resuelto', false );							
		sessionStorage.setItem("Convalidado", false );						
						
	});	
	
</script>

