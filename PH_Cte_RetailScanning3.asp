<!Doctype html>
<!-- PH_Cte_RetailScanning.asp - 12jul21 - 15jul21 -->
<html lang="es" >
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
	<link href="retsemanal/css/convivencia2.css"  rel="stylesheet" type="text/css" />
	<!-- MultiSelect CSS & JS library -->
	<link rel="stylesheet" href="retsemanal/css/bootstrap-multiselect-0915.css">
	<!--===============================================================================================-->	
	<link rel="stylesheet" href="css/homePantry.css" type="text/css">
	<link rel="stylesheet" type="text/css" href="css/perfect-scrollbar.css">
	<link rel="stylesheet" type="text/css" href="css/util.css">
	<link rel="stylesheet" type="text/css" href="css/mainRS.css">	
	
</head>
<body topmargin="0">

	<!--#include file="estiloscss.asp"-->
	<!--#include file="meta.asp"-->
	<!--#include file="encabezado.asp"-->
	<!--#include file="nn_subN.asp"-->
	<!--#include file="in_DataEN.asp"-->

	<%
		' 12jul21 - 18jul21
		Apertura
		LeePar
		if ed_iPas<>4 then
			Encabezado
		end if
		dim Mostrar
		Mostrar = 0
		if Mostrar = 1 and idCliente = 1 then
			sVar = "text"
		else
			sVar = "hidden"
		end if
		'		
	%>

		<!--hidden-->
		<input type="<%=sVar%>" name="Filtro" id="Filtro" align="right" size=250>
		<input type="hidden" name="Cliente" id="Cliente"  align="right" size=4 value="">
		<input type="hidden" name="Cat" id="Cat" align="right" size=4 value="">
		
		<div class="container-fluid" id="grad3" >

			<div class="form-group">

				<!-- ZONA A -->
				<div class="col-sm-5">
				
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company"><i class="fas fa-shapes"></i>&nbsp;Categoria:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboCategoria" name="cboCategoria" style="width: 285px;" >
							<option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company"><i class="fas fa-globe-americas"></i>&nbsp;Area:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboArea" name="cboArea" multiple="multiple">
							<option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company"><i class="fas fa-globe-americas"></i>&nbsp;Zona:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboZona" name="cboZona" multiple="multiple">
							<option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company"><i class="fas fa-home"></i>&nbsp;Canal:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboCanal" name="cboCanal" multiple="multiple">
							<option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company"><i class="fas fa-industry"></i>&nbsp;Fabricante:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboFabricante" name="cboFabricante" multiple="multiple">
						  <option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-2 col-sm-2 lb" for="company"><i class="fas fa-registered"></i>&nbsp;Marca:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboMarca" name="cboMarca" class="form-control input-sm" multiple="multiple">
						  <option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>					
					
				</div>

				<!-- Imagen -->
				<div class="col-sm-2 id="image">
					<img class="img-responsive img-center" src="images/logo/LogoRetailScanning2.png" />
				</div>

				<!-- ZONA B -->
				<div class="col-sm-5">
					<label class="control-label col-sm-offset-2 col-sm-3 lb" for="company"><i class="fas fa-project-diagram"></i>&nbsp;Segmento:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboSegmento" name="cboSegmento" class="form-control input-sm" multiple="multiple">
						  <option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					<label class="control-label col-sm-offset-2 col-sm-3 lb" for="company"><i class="fas fa-ruler-combined"></i>&nbsp;Tamaño:</label>					
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboTamano" name="cboTamano" class="form-control input-sm" multiple="multiple">
						  <option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-2 col-sm-3 lb" for="company"><i class="fas fa-barcode"></i>&nbsp;Producto:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboProducto" name="cboProducto" class="form-control input-sm" multiple="multiple">
							<option>-- Seleccione --</option>
						</select>
					</div>
					<label class="control-label col-sm-offset-2 col-sm-3 lb" for="company"><i class="fas fa-tachometer-alt"></i>&nbsp;Indicadores:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboIndicadores" name="cboIndicadores" class="form-control input-sm" multiple="multiple">
						  <option>-- Seleccione --</option>
						</select>
					</div>
					<label class="control-label col-sm-offset-2 col-sm-3 lb" for="company"><i class="fas fa-calendar-plus"></i>&nbsp;Semanas:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboSemanas" name="cboSemanas" class="form-control input-sm" multiple="multiple">
						  <option>-- Seleccione --</option>
						</select>
					</div>					
					<!-- BOTONES -->
					<div class="col-sm-12 separa">
						<div class="col-sm-4 separa">
							<button id="BtnAplicarFiltro"  title="Procesar" type="submit" class="btn btn-block btn-xs btn-success"><i class="fas fa-check "></i>&nbsp;Aplicar Filtro</button>
						</div>							
						<div class="col-sm-4 separa">
							<button id="BtnExcel"  title="Exportar a Excel" type="submit" class="btn btn-block btn-xs btn-primary" onclick="Excel();"><i class="fas fa-download"></i>&nbsp;Exportar Excel</button>
						</div>
						<div class="col-sm-4 separa">
							<button id="BtnBorrar"  title="Borrar filtros" type="submit" class="btn btn-block btn-xs btn-danger" onclick="Reset();"><i class="fas fa-recycle "></i>&nbsp;Borrar Filtro</button>
						</div>
					</div>
				</div>

			</div>
			<!-- < / class="form-group" -->

		</div>
		<!-- < / class="container-fluid" id="grad1" -->
		
		<div class="container-fluid text-center text-primary" id="cargando" style="display:none;">
			<br>
			<span ><img src="images/ajax-loader8.gif"><strong>&nbsp;Procesando..., Espere!</strong></span>
		</div>
		<hr>
		<div class="container-fluid text-center text-primary" id="DivRetailScanningSem" style="display:none;" >
			<!-- Mostrar la tabla con los resultados -->			
		</div>
	

	<%conexion.close%>

</body>
</html>
<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/jquery-3.1.1.min.js"></script>
<script src="js/sweetalert.min.js"></script>
<script src="js/bootstrap.min.js"></script>
<script src="retsemanal/js/funcionesV3.js"></script>
<script src="retsemanal/js/refillCombos.js"></script>
<!-- MultiSelect CSS & JS library -->
<script src="retsemanal/js/bootstrap-multiselect-0915.js"></script>

<script src="js/perfect-scrollbar.min.js"></script>
<script>
	$('.js-pscroll').each(function(){
		var ps = new PerfectScrollbar(this);
		$(window).on('resize', function(){
			ps.update();
		})
	});	
</script>
<script src="js/main.js"></script>

<script>
	
	$(function () {
        sessionStorage.clear();
		sessionStorage.setItem("idCliente", <%=Session("idCliente")%>);			
		$("#Cliente").val(<%=Session("idCliente")%>);		
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
		$(function() {		
			ValidarCliente();			
		});				
	});
	
	$('#BtnAplicarFiltro').click(function() {
			
		//debugger;
		$("#DivRetailScanningSem").css("display", "none");		
		let categ = $("#cboCategoria").val();
		
		if(categ==null){
			swal("Alerta","Debe seleccionar una Categoria..!","error");
			return false;
		}		
		let area        = $("#cboArea :selected").map((_,e) => e.value).get();
		if(area.length==0 || area==undefined ){
			area=0;						
		}else{
			area  = area.join(); 
		}
		let zona        = $("#cboZona :selected").map((_,e) => e.value).get();
		if(zona.length==0 || zona==undefined ){
			zona=0;						
		}else{
			zona  = zona.join(); 
		}
		let canal       = $("#cboCanal :selected").map((_,e) => e.value).get();
		if(canal.length==0 || canal==undefined ){
			canal=0;						
		}else{
			canal  = canal.join(); 
		}
		let fabricante  = $("#cboFabricante :selected").map((_,e) => e.value).get();
		if(fabricante.length==0 || fabricante==undefined ){
			fabricante=0;						
		}else{
			fabricante  = fabricante.join(); 
		}
		let marca       = $("#cboMarca :selected").map((_,e) => e.value).get();
		if(marca.length==0 || marca==undefined ){
			marca=0;						
		}else{
			marca  = marca.join(); 
		}
		let segmento    = $("#cboSegmento :selected").map((_,e) => e.value).get();
		if(segmento.length==0 || segmento==undefined ){
			segmento=0;						
		}else{
			segmento  = segmento.join(); 
		}
		let tamano      = $("#cboTamano :selected").map((_,e) => e.value).get();
		if(tamano.length==0 || tamano==undefined ){
			tamano=0;						
		}else{
			tamano  = tamano.join(); 
		}
		let producto    = $("#cboProducto :selected").map((_,e) => e.value).get();
		if(producto.length==0 || producto==undefined ){
			producto=0;						
		}else{
			producto  = producto.join(); 
		}
		let indicadores = $("#cboIndicadores :selected").map((_,e) => e.value).get();
		if(indicadores.length==0 || indicadores==undefined ){
			indicadores=0;						
		}else{
			indicadores  = indicadores.join(); 
		}
		let semanas     = $("#cboSemanas :selected").map((_,e) => e.value).get();
		if (semanas.length == 0 || semanas==undefined) {		
			swal("Alerta","Seleccionar una Semana","error");
			return false;		
		}else{
			let columnastotal = 5;
			if (semanas.length > columnastotal){
				swal("Alerta","Solo se pueden Seleccionar hasta un Maximo de 5 Semanas","error");
				return false;
			}		
			semanas  = semanas.join(); 
		}				
		//		
		let ajax = { 
			cat : categ,
			are : area,
			zon : zona,
			can : canal,
			fab : fabricante,
			mar : marca,
			seg : segmento,
			tam : tamano,
			pro : producto,		
			ind : indicadores,
			sem : semanas,			
		};
				
		$('#DivRetailScanningSem').html("");
		$.ajax({
			//url:'g_CteRetailScanningSem.asp?'+stodo,
			url:'RetSem_Excel.asp',
			type:'POST',
			data: ajax,
			beforeSend: function(objeto){
				$("#cargando").css("display", "block");		
			}				
		})
		/*Si la consulta se realizo con exito*/
		.done(function(data) {			
			console.log(data);
			//debugger;
			$('#DivRetailScanningSem').html(data);
			$("#cargando").css("display", "none");		
			$("#DivRetailScanningSem").css("display", "block");						
		})
		/*Si la consulta Fallo*/
		.fail(function() {
			$("#cargando").css("display", "none");										
			swal("Algo salio mal.!","Intente de nuevo", "error");
		},'html');
		
	});
	
</script>

