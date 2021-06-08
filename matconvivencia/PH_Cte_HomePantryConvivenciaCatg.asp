<!Doctype html>
<!-- PH_Cte_HomePantryConvivencia - 09abr21 - 11abr21 -->
<html >
<head>
	<title>| Convivencia Categoria |</title>
	<meta charset="UTF-8">
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="favicon.ico" rel="icon" type="image/x-icon">
	<link href="css/sweetalert.css"  rel="stylesheet" type="text/css" />
	<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />	
	<link href="css/convivencia2.css"  rel="stylesheet" type="text/css" />			
</head>
<body topmargin="0">
	
	<!--#include file="../estiloscss.asp"-->
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
					
		<div class="form-group">
		
			<!-- CATEGORIA A -->			
			<div class="col-sm-5">			
				<br>			
			  	<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Categoria</label>
			  	<div class="col-sm-6 col-md-6 separa">
					<select class="form-control input-sm" id="cboCategoria_A" name="cboCategoria_A" >				  		
				  		<option value="0" selected disabled >-- Seleccione -- </option>
					</select> 
			  	</div>			  
				<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Fabricante</label>
				<div class="col-sm-6 col-md-6 separa">
					<select id="cboFabricante_A" class="form-control input-sm">
					  <option value="0" selected disabled >-- Seleccione -- </option>
					</select> 
				</div>				  
				<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Marca</label>
				<div class="col-sm-6 col-md-6 separa">
					<select id="cboMarca_A" class="form-control input-sm">
					  <option value="0" selected disabled >-- Seleccione -- </option>
					</select> 
				</div>				  
				<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Segmento</label>
				<div class="col-sm-6 col-md-6 separa">
					<select id="cboSegmento_A" class="form-control input-sm">
					  <option value="0" selected disabled >-- Seleccione -- </option>
					</select> 
				</div>
				<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Tamaño</label>
				<div class="col-sm-6 col-md-6 separa">
					<select id="cboTamano_A" class="form-control input-sm">
					  <option value="0" selected disabled >-- Seleccione -- </option>
					</select> 
				</div>  
			</div>
			
			<!-- Imagen -->
			<div class="col-sm-2 coolBackground">
				<br>
    		    <img src="images/versus6.png" class="img-responsive  img-center" />	
  			</div>			
			
			<!-- CATEGORIA B -->						
			<div class="col-sm-5">			
				<br>			
			  	<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Categoria</label>
			  	<div class="col-sm-6 col-md-6 separa">
					<select id="cboCategoria_B" class="form-control input-sm">
				  		<option>-- Seleccione --</option>				  		
					</select> 
			  	</div>			  
				<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Fabricante</label>
				<div class="col-sm-6 col-md-6 separa">
					<select id="cboFabricante_B" class="form-control input-sm">
					  <option>-- Seleccione --</option>					  
					</select> 
				</div>				  
				<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Marca</label>
				<div class="col-sm-6 col-md-6 separa">
					<select id="cboMarca_B" class="form-control input-sm">
					  <option>-- Seleccione --</option>					  
					</select> 
				</div>				  
				<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Segmento</label>
				<div class="col-sm-6 col-md-6 separa">
					<select id="cboSegmento_B" class="form-control input-sm">
					  <option>-- Seleccione --</option>					  
					</select> 
				</div>		   				
				<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Tamaño</label>
				<div class="col-sm-6 col-md-6 separa">
					<select id="cboTamano_B" class="form-control input-sm">
					  <option>-- Seleccione --</option>					  
					</select> 
				</div>		   			  
			</div>
	  
		</div>
		<!-- < / class="form-group" -->	
						
		<div class="form-group" id="grad2">		
			<!-- AREA -->			
			<div class="col-sm-4">			
				<br>			
			  	<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Area</label>
			  	<div class="col-sm-6 col-md-6 separa">
					<select class="form-control input-sm" id="cboArea" name="cboArea" >				  		
				  		<option value="0" selected disabled >-- Seleccione -- </option>
					</select> 
			  	</div>			  				
			</div>
			
			<!-- AREA -->			
			<div class="col-sm-4">			
				<br>			
			  	<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Periodo</label>
			  	<div class="col-sm-6 col-md-6 separa">
					<select class="form-control input-sm" id="cboPeriodo" name="cboPeriodo" >				  		
				  		<option value="0" selected disabled >-- Seleccione -- </option>
					</select> 
			  	</div>			  				
			</div>
			
			<!-- AREA -->			
			<div class="col-sm-4">			
				<br>					
				<button id="BtnProcesar"  title="Procesar" type="submit" class="btn btn-block btn-xs btn-success" onclick="Procesar();"><i class="fas fa-check fa-2x"></i></button>
				<button id="borrar"  title="Borrar Pantalla" type="submit" class="btn btn-block btn-xs btn-danger" onclick="Reset();"><i class="fas fa-recycle fa-2x"></i></button>				
			</div>
						
			
			
	  
		</div>
		<!-- < / class="form-group" -->	


		
	</div>        
	<!-- < / class="container-fluid" id="grad1" -->
		
	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;">
		<br>
		<span ><img src="images/ajax-loader8.gif"><strong>&nbsp;Espere, Procesando...!</strong></span>
	</div>
	<br>
	
	<div class="container-fluid" id="detallesMaestro" style="display:block;" >
				
		<div class="form-group" id="detallesPaso1" style="display:none;">			
			<!-- Code by w3codegenerator.com -->
			<div class="table-responsive" id="Paso1">
				<!---->									 					
			</div>		
		</div>		
		<div class="form-group" id="detallesPaso2" style="display:none;">			
			<!-- Code by w3codegenerator.com -->
			<div class="table-responsive" id="Paso2">
				<!---->									 					
			</div>		
		</div>	
		<br>	
		<div class="form-group" id="detallesPaso3" style="display:none;">						
			<table class="table table-striped table-condensed text-center" style=" margin: auto; width: 65% !important; ">
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
		<br>	
		<div class="form-group" id="detallesPaso4" style="display:none;">			
			<!-- Code by w3codegenerator.com -->
			<div class="table-responsive" id="Paso4">
				<!---->									 					
			</div>		
		</div>
			
	</div>
	<!-- </ class="container-fluid" id="detallesMaestro" style="display:block;" > -->

	<hr>	
		
	<%conexion.close%>
	
</body>
</html>
<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/jquery-3.1.1.min.js"></script>
<script src="js/sweetalert.min.js"></script>
<script src="js/bootstrap.min.js"></script>
<script src="matconvivencia/js/funcionesV1.js"></script>

<script>
	
	$(document).ready(function() {				
		$(function() {
			//debugger;
			LlenarCombos();
		});					
	});	
	
	sessionStorage.setItem("idCliente", <%=Session("idCliente")%>);
	
</script>

