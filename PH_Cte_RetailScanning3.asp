<!Doctype html>
<!-- PH_Cte_RetailScanning.asp - 12jul21 - -->
<html >
<head>
	<title>| Retail Scanning Semanal |</title>
	<meta charset="UTF-8">
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="favicon.ico" rel="icon" type="image/x-icon">
	<link href="css/sweetalert.css"  rel="stylesheet" type="text/css" />
	<link href="matconvivencia/css/convivencia2.css"  rel="stylesheet" type="text/css" />
	<!-- MultiSelect CSS & JS library -->
	<link rel="stylesheet" href="matconvivencia/css/bootstrap-multiselect-0915.css">
</head>
<body topmargin="0">

	<!--#include file="estiloscss.asp"-->
	<!--#include file="meta.asp"-->
	<!--#include file="encabezado.asp"-->
	<!--#include file="nn_subN.asp"-->
	<!--#include file="in_DataEN.asp"-->

	<%
		' 12jul21 - 
		Apertura
		LeePar
		if ed_iPas<>4 then
			Encabezado
		end if
		'
	%>

		<!--hidden-->
		<input type="<%=sVar%>" name="Filtro" id="Filtro" align="right" size=250>
		<input type="hidden" name="Cliente" id="Cliente"  align="right" size=4 value="">
		<input type="hidden" name="Cat" id="Cat" align="right" size=4 value="">
		
		<div class="container-fluid" id="grad1" >

			<div class="form-group">

				<!-- CATEGORIA A -->
				<div class="col-sm-5">
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company"><i class="fas fa-shapes"></i>&nbsp;Categoria:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboCategoria" name="cboCategoria" >
							<option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company"><i class="fas fa-globe-americas"></i>&nbsp;Area:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboArea" name="cboArea" multiple="">
							<option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company"><i class="fas fa-globe-americas"></i>&nbsp;Zona:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboZona" name="cboZona" multiple="">
							<option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company"><i class="fas fa-home"></i>&nbsp;Canal:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboCanal" name="cboCanal" multiple="">
							<option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company"><i class="fas fa-industry"></i>&nbsp;Fabricante:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboFabricante" name="cboFabricante" multiple="">
						  <option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company"><i class="fas fa-registered"></i>&nbsp;Marca:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboMarca" name="cboMarca" class="form-control input-sm" multiple="">
						  <option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					
					
				</div>

				<!-- Imagen -->
				<div class="col-sm-2 id="image">
					<img src="images/logo/LogoReatilScanning.png" class="img-responsive img-center" />
				</div>

				<!-- CATEGORIA B -->

				<div class="col-sm-5">
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Segmento:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboSegmento" name="cboSegmento" class="form-control input-sm" multiple="">
						  <option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Tamaño:</label>
					
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboTamano" name="cboTamano" class="form-control input-sm" multiple="">
						  <option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Producto:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboProducto" name="cboProducto" class="form-control input-sm" multiple="">
							<option>-- Seleccione --</option>
						</select>
					</div>
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Indicadores:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboIndicadores" name="cboIndicadores" class="form-control input-sm" multiple="">
						  <option>-- Seleccione --</option>
						</select>
					</div>
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Semanas:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboSemanas" name="cboSemanas" class="form-control input-sm" multiple="">
						  <option>-- Seleccione --</option>
						</select>
					</div>					
					<!-- BOTONES -->
					<div class="col-sm-12 separa">
							<div class="col-sm-4 separa">
								<button id="BtnValidarProceso"  title="Procesar" type="submit" class="btn btn-block btn-xs btn-success"><i class="fas fa-check "></i>&nbsp;Procesar</button>
							</div>							
							<div class="col-sm-4 separa">
								<button id="BtnExcel"  title="Exportar a Excel" type="submit" class="btn btn-block btn-xs btn-primary" onclick="Excel();"><i class="fas fa-download"></i>&nbsp;Excel</button>
							</div>
							<div class="col-sm-4 separa">
								<button id="BtnBorrar"  title="Borrar filtros" type="submit" class="btn btn-block btn-xs btn-danger" onclick="Reset();"><i class="fas fa-recycle "></i>&nbsp;Borrar</button>
							</div>
						</div>
					</div>

			</div>
			<!-- < / class="form-group" -->


		</div>
		<!-- < / class="container-fluid" id="grad1" -->

		<div class="container-fluid barrabotones">
			
			

		</div>
		<!-- < / class="form-group" -->

	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;">
		<br>
		<span ><img src="images/ajax-loader8.gif"><strong>&nbsp;Espere, Procesando...!</strong></span>
	</div>
	<br>
	
	<hr>

	<div class="container-fluid" id="tablaResultados" style="display:none;" >

		<!-- Mostrar la tabla con los resultados -->
		
		
		

	</div>
	<!-- </ class="container-fluid" id="detallesMaestro" style="display:block;" > -->

	<%conexion.close%>

</body>
</html>
<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/jquery-3.1.1.min.js"></script>
<script src="js/sweetalert.min.js"></script>
<script src="js/bootstrap.min.js"></script>
<script src="retsemanal/js/url.js"></script>
<script src="retsemanal/js/funcionesV2.js"></script>
<script src="retsemanal/js/procesarV6.js"></script>
<!-- MultiSelect CSS & JS library -->
<script src="retsemanal/js/bootstrap-multiselect-0915.js"></script>

<script>
	$(document).ready(function() {
		sessionStorage.clear();
		sessionStorage.setItem("idCliente", <%=Session("idCliente")%>);	
		$("#Cliente").val(<%=Session("idCliente")%>);
		url();	 			
		$("#cboCategoria").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, });
		$("#cboArea").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, });
		$("#cboZona").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, });
		$("#cboCanal").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, });
		$("#cboFabricante").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, });
		$("#cboMarca").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, });
		$("#cboSegmento").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, });
		$("#cboTamano").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, });
		$("#cboProducto").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, });
		$("#cboIndicadores").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, });
		$("#cboSemanas").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, });
		
	});
</script>


<script>
	$(document).ready(function() {
		//
		$(function() {
			//debugger;
			//LlenarCombos();
			LlenarCategoria();
		});
	});
	
</script>

