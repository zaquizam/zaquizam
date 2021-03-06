<!Doctype html>
<!-- PH_Cte_HomePantryConvivencia - 09abr21 - 14jun21 -->
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
	<!--<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />-->
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
						<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Categoria</label>
						<div class="col-sm-6 col-md-6 separa">
							<select class="form-control input-sm" id="cboCategoria_A" name="cboCategoria_A" >
								<option value="0" selected disabled >-- Seleccione -- </option>
							</select>
						</div>
						<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Fabricante</label>
						<div class="col-sm-6 col-md-6 separa">
							<select class="form-control input-sm" id="cboFabricante_A" name="cboFabricante_A">
							  <option value="0" selected disabled >-- Seleccione -- </option>
							</select>
						</div>
						<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Marca</label>
						<div class="col-sm-6 col-md-6 separa">
							<select id="cboMarca_A" name="cboMarca_A" class="form-control input-sm">
							  <option value="0" selected disabled >-- Seleccione -- </option>
							</select>
						</div>
						<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Segmento</label>
						<div class="col-sm-6 col-md-6 separa">
							<select id="cboSegmento_A" name="cboSegmento_A" class="form-control input-sm">
							  <option value="0" selected disabled >-- Seleccione -- </option>
							</select>
						</div>
						<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Tama??o</label>
						<div class="col-sm-6 col-md-6 separa">
							<select id="cboRangTamanoA" name="cboRangTamanoA" class="form-control input-sm">
							  <option value="0" selected disabled >-- Seleccione -- </option>
							</select>
						</div>
					</div>

					<!-- Imagen -->
					<div class="col-sm-2 id="image">
						<img src="images/convivencia/ab2.jpg" class="img-responsive img-center" />
					</div>

					<!-- CATEGORIA B -->

					<div class="col-sm-5">
						<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Categoria</label>
						<div class="col-sm-6 col-md-6 separa">
							<select id="cboCategoria_B" name="cboCategoria_B" class="form-control input-sm">
								<option>-- Seleccione --</option>
							</select>
						</div>
						<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Fabricante</label>
						<div class="col-sm-6 col-md-6 separa">
							<select id="cboFabricante_B" name="cboFabricante_B" class="form-control input-sm">
							  <option>-- Seleccione --</option>
							</select>
						</div>
						<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Marca</label>
						<div class="col-sm-6 col-md-6 separa">
							<select id="cboMarca_B" name="cboMarca_B" class="form-control input-sm">
							  <option>-- Seleccione --</option>
							</select>
						</div>
						<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Segmento</label>
						<div class="col-sm-6 col-md-6 separa">
							<select id="cboSegmento_B" name="cboSegmento_B" class="form-control input-sm">
							  <option>-- Seleccione --</option>
							</select>
						</div>
						<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Tama??o</label>
						<div class="col-sm-6 col-md-6 separa">
							<select id="cboRangTamanoB" name="cboRangTamanoB" class="form-control input-sm">
							  <option>-- Seleccione --</option>
							</select>
						</div>
					</div>

				</div>
				<!-- < / class="form-group" -->


		</div>
		<!-- < / class="container-fluid" id="grad1" -->

		<div class="container-fluid barrabotones">
			<!-- AREA -->
			<div class="col-sm-4">
				<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Area</label>
				<div class="col-sm-6 col-md-6 separa">
					 <select id="cboArea" multiple="">
						 <!-- Combo -->
					</select>
				</div>
			</div>

			<!-- PERIODO -->
			<div class="col-sm-4">
				<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company">Periodo</label>
				<div class="col-sm-6 col-md-6 separa">
					<select id="cboPeriodo" multiple="">
						<!-- Combo -->
					</select>
				</div>
			</div>

			<!-- BOTONES -->
			<div class="col-sm-4">
				<div class="col-sm-6 col-md-6 separa">
					<button id="BtnValidarProceso"  title="Procesar" type="submit" class="btn btn-block btn-xs btn-success"><i class="fas fa-check fa-2x"></i></button>
				</div>
				<div class="col-sm-6 col-md-6 separa">
					<button id="BtnBorrar"  title="Borrar filtros" type="submit" class="btn btn-block btn-xs btn-danger" onclick="Reset();"><i class="fas fa-recycle fa-2x"></i></button>
				</div>
			</div>

		</div>
		<!-- < / class="form-group" -->

	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;">
		<br>
		<span ><img src="images/ajax-loader8.gif"><strong>&nbsp;Espere, Procesando...!</strong></span>
	</div>
	<br>

	<div class="container-fluid" id="detalleTotalHogares" style="display:none;" >
			<!-- TOTAL A -->
			<div class="col-sm-4">
				<label class="control-label col-sm-offset-2 col-sm-4 lb" for="company">Total Hogares (A):</label>
				<div class="col-sm-6 col-md-6 separa">
					<span id="totalHogaresA">0</span>
				</div>
			</div>

			<!-- TOTAL A -->
			<div class="col-sm-4">
				<label class="control-label col-sm-offset-2 col-sm-4 lb" for="company">Total Hogares:</label>
				<div class="col-sm-6 col-md-6 separa">
					<span id="totalHogaresAB">0</span>
				</div>
			</div>

			<!-- TOTAL A -->
			<div class="col-sm-4">
				<label class="control-label col-sm-offset-2 col-sm-4 lb" for="company">Total Hogares (B):</label>
				<div class="col-sm-6 col-md-6 separa">
					<span id="totalHogaresB">0</span>
				</div>
			</div>

	</div>

	<hr>

	<div class="container-fluid" id="tablaResultados" style="display:block;" >

		<div class="col-sm-6">
			<table class="table table-borderless table-striped table-condensed text-center" style=" margin: auto; width: 65% !important; ">
				<thead>
				<tr>
					<th colspan='3' class='text-center lb'><i class='fas fa-check-double'></i>&nbsp;MATRIZ DE CONVIVENCIA ( A - B )</th>
				</tr>
				  <tr>
					<th></th>
					<th scope="col" class="text-center lb">A</th>
					<th scope="col" class="text-center lb">B</th>
				  </tr>
				</thead>
				<tbody>
				  <tr>
					<td class="text-center lb">A</td>
					<td id="total_AA">100%</td>
					<td id="total_AB">15%</td>
				  </tr>
				  <tr>
					<td class="text-center lb">B</td>
					<td id="total_BA">17%</td>
					<td id="total_BB">100%</td>
				  </tr>
				</tbody>
			</table>

		</div>
		
		<div class="col-sm-6">
			<table class="table table-striped table-condensed text-center" style=" margin: auto; width: 65% !important; ">
				<thead>
				<tr>
					<th colspan='3' class='text-center lb'><i class='fas fa-check-double'></i>&nbsp;PENETRACION</th>
				</tr>				 
				</thead>
				<tbody>
				  <tr>
					<td class="text-center lb">PENETRACION A-B:</td>
					<td class="text-left">26</td>					
				  </tr>
				  <tr>
					<td class="text-center lb">PENETRACION (A):</td>
					<td class="text-center">13</td>					
				  </tr>
				  <tr>
					<td class="text-center lb">PENETRACION (B):</td>
					<td class="text-center">16</td>					
				  </tr>
				  <tr>
					<td class="text-center lb">CONVIVENCIA:</td>
					<td class="text-center">72</td>					
				  </tr>
				</tbody>
			</table>

		</div>
		

	</div>
	<!-- </ class="container-fluid" id="detallesMaestro" style="display:block;" > -->

	<%conexion.close%>

</body>
</html>
<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/jquery-3.1.1.min.js"></script>
<script src="js/sweetalert.min.js"></script>
<script src="js/bootstrap.min.js"></script>
<script src="matconvivencia/js/funcionesV2.js"></script>
<script src="matconvivencia/js/procesar.js"></script>
<!-- MultiSelect CSS & JS library -->
<script src="matconvivencia/js/bootstrap-multiselect-0915.js"></script>

<script>
	$(document).ready(function() {
	 	$("#cboArea").multiselect({	buttonWidth:'auto', disableIfEmpty: true, });
		$("#cboPeriodo").multiselect({	buttonWidth:'auto', disableIfEmpty: true, });
		//$("#cboCategoria_A").multiselect({	buttonWidth:'350px', disableIfEmpty: true, });
	});
</script>


<script>
	$(document).ready(function() {
		//
		$(function() {
			//debugger;
			LlenarCombos();
		});
	});
	sessionStorage.setItem("idCliente", <%=Session("idCliente")%>);
</script>

<script>
	function getSelectedValues() {
	  var selectedVal = $("#multiselect").val();
		for(var i=0; i<selectedVal.length; i++){
			function innerFunc(i) {
				setTimeout(function() {
					location.href = selectedVal[i];
				}, i*2000);
			}
			innerFunc(i);
		}
	}
</script>

