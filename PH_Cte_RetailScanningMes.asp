<!Doctype html>
<!-- PH_Cte_RetailScanning.asp - 12jul21 - 27ene22 -->
<html lang="es" >
<head>
	<title>| RS Reporte Mensual |</title>
	<meta charset="UTF-8">
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="favicon.ico" rel="icon" type="image/x-icon">
	<link href="css/sweetalert.css"  rel="stylesheet" type="text/css" />
	<link href="rsrepmensual/css/retmensual.css"  rel="stylesheet" type="text/css" />
	<link href="css/bootstrap-multiselect-0915.css" rel="stylesheet" >
	<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />
	<link href="rsrepmensual/css/tablamen.css" rel="stylesheet" type="text/css" >
</head>

<body topmargin="0">

	<!--#include file="estiloscss.asp"-->
	<!--#include file="meta.asp"-->
	<!--#include file="encabezado.asp"-->
	<!--#include file="nn_subN.asp"-->
	<!--#include file="in_DataEN.asp"-->

	<%
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

		<div class="container-fluid" id="grad3" style="visibility:hidden;" >

			<div class="row">

				<!-- ZONA A -->
				<div class="col-sm-5">

					<label class="control-label col-sm-offset-1 col-sm-3 lb"><i class="fas fa-shapes"></i>&nbsp;Categoria:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboCategoria" name="cboCategoria" style="width: 285px;" >
							<option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>

					<label class="control-label col-sm-offset-1 col-sm-3 lb"><i class="fas fa-globe-americas"></i>&nbsp;Area:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboArea" name="cboArea" multiple="multiple">
							<option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>

					<label class="control-label col-sm-offset-1 col-sm-3 lb" for="company"><i class="fas fa-globe-americas"></i>&nbsp;Zona:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboZona" name="cboZona" multiple="multiple">
							<option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>

					<label class="control-label col-sm-offset-1 col-sm-3 lb" for="company"><i class="fas fa-home"></i>&nbsp;Canal:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboCanal" name="cboCanal" multiple="multiple">
							<option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>

					<label class="control-label col-sm-offset-1 col-sm-3 lb" for="company" id="tipoFabricante" ><i class="fas fa-industry"></i>&nbsp;Fabricante:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboFabricante" name="cboFabricante" multiple="multiple">
						  <option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>

					<label class="control-label col-sm-offset-1 col-sm-3 lb" for="company"><i class="fas fa-registered"></i>&nbsp;Marca:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboMarca" name="cboMarca" class="form-control input-sm" multiple="multiple">
						  <option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>

				</div>

				<!-- ZONA B -->
				<div class="col-sm-5">
				
					<label class="control-label col-sm-offset-1 col-sm-3 lb" for="company"><i class="fas fa-project-diagram"></i>&nbsp;Segmento:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboSegmento" name="cboSegmento" class="form-control input-sm" multiple="multiple">
						  <option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-1 col-sm-3 lb" for="company"><i class="fas fa-ruler-combined"></i>&nbsp;Tamaño:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboTamano" name="cboTamano" class="form-control input-sm" multiple="multiple">
						  <option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>

					<label class="control-label col-sm-offset-1 col-sm-3 lb" for="company"><i class="fas fa-barcode"></i>&nbsp;Producto:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboProducto" name="cboProducto" class="form-control input-sm" multiple="multiple">
							<option>-- Seleccione --</option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-1 col-sm-3 lb" for="company"><i class="fas fa-tachometer-alt"></i>&nbsp;Indicadores:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboIndicadores" name="cboIndicadores" class="form-control input-sm" multiple="multiple">
						  <option>-- Seleccione --</option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-1 col-sm-3 lb" for="company"><i class="fas fa-calendar-week"></i>&nbsp;Semanas:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboSemanas" name="cboSemanas" class="form-control input-sm" multiple="multiple">
						  <option>-- Seleccione --</option>
						</select>
					</div>
					
					<label class="control-label col-sm-offset-1 col-sm-3 lb" for="company"><i class="fas fa-calendar-alt"></i>&nbsp;Meses:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboMeses" name="cboMeses" class="form-control input-sm" multiple="multiple">
						  <option>-- Seleccione --</option>
						</select>
					</div>
					
					<!--
					<label class="control-label col-sm-offset-1 col-sm-3 lb" for="company"><i class="fas fa-print"></i>&nbsp;Histórico:</label>
					<div class="col-sm-6 col-md-6 separa">
						<label><input type="checkbox" value="1" id="chkReporte"></label>
					</div>
					-->

				</div>

				<!-- Imagen -->
				<div class="col-sm-2">
				
					<p class="text-center"><img class="img-fluid"  src="images/logo/LogoRetailScanning2.png" width="128" height="100"/></p>					
					
					
					<div class="btn-group btn-group btn-group-lg text-right" role="group"> 
						<button type="button" id="BtnAplicarFiltro" title="Procesar" type="submit" class="btn btn-success"><i class="fas fa-check"></i>&nbsp;</button>
						<button type="button" id="BtnExcel"  title="Exportar a Excel" type="submit" class="btn btn-primary"><i class="fas fa-download"></i>&nbsp;</button>
						<button type="button" id="BtnBorrar" title="Borrar filtros"  type="submit" class="btn btn-danger"><i class="fas fa-recycle"></i>&nbsp;</button>
					</div>	
					
				</div>				
						

			</div>
			<!-- < / class="row" -->

		</div>	<!-- < / class="container-fluid" id="grad1" -->

		<div class="container-fluid text-center text-danger" id="cargando" style="display:none;">
			<br>
			<span ><img src="images/ajax-loader8.gif"><strong>&nbsp;Espere, cargando filtros....!</strong></span>
		</div>
		
		<div class="container-fluid text-center text-primary" id="procesando" style="display:none;">
			<br>			
			<span ><img src="images/atenas518.gif" border="0" width="48" height="48" ><strong>&nbsp;Espere, procesando....!</strong></span>
		</div>
		
		<div class="container-fluid text-center text-primary" id="procesandoExcel" style="display:none;">
			<br>			
			<span ><img src="images/atenas518.gif" border="0" width="48" height="48" ><strong>&nbsp;Espere, generando Excel....!</strong></span>
		</div>
		
		<hr>
		
		<div class="container-fluid text-center text-primary" id="DivRetailScanningSem" style="display:none;" >
			<!-- Mostrar la tabla con los resultados -->
		</div>
		
		<div class="container-fluid text-center text-primary" id="DivRetailScanningExcel" style="display:none;" >
			<!-- Mostrar la tabla con los resultados Excel-->
		</div>

	<%conexion.close%>

</body>
</html>
<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/jquery-3.1.1.min.js"></script>
<script src="js/sweetalert.min.js"></script>
<script src="js/bootstrap.min.js"></script>
<script src="js/bootstrap-multiselect-0915.js"></script>
<script src="rsrepmensual/js/funcionesMenV05.js"></script>
<script src="rsrepmensual/js/refillCombosMenV05.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/amcharts/3.21.15/plugins/export/libs/FileSaver.js/FileSaver.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.15.6/xlsx.full.min.js"></script>

<script>

	$(function () {
       sessionStorage.clear();
		sessionStorage.setItem("idCliente", <%=Session("idCliente")%>);
		sessionStorage.setItem("repCompleto", 0);
		sessionStorage.setItem("eXcel", 0);
		$("#Cliente").val(<%=Session("idCliente")%>);
		$("#cboCategoria").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, maxHeight: 200,});
		$("#cboArea").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, maxHeight: 200, });
		$("#cboZona").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, maxHeight: 200, });
		$("#cboCanal").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, maxHeight: 200, });
		$("#cboFabricante").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, maxHeight: 200, });
		$("#cboMarca").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, maxHeight: 200, });
		$("#cboSegmento").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, maxHeight: 200, });
		$("#cboTamano").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, maxHeight: 200, });
		$("#cboProducto").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, maxHeight: 200, });
		$("#cboIndicadores").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, maxHeight: 200, });
		$("#cboSemanas").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, maxHeight: 200, });
		$("#cboMeses").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, maxHeight: 200, });
		document.getElementById('grad3').style.visibility="visible";
    });

</script>

<script>

	$(document).ready(function() {
		$(function() {
			ValidarCliente();
		});
	});
	//
	$(function () {

		$('#BtnBorrar').click(function() {
			event.preventDefault();
			Reset();
		});
		//
		$('#BtnAplicarFiltro').click(function() {
			//
			debugger;
			event.preventDefault();
			let ExecPrograma;
			let semanas;
			bLoquear();
			//
			$("#DivRetailScanningSem").css("display", "none");
			//
			let categ = $("#cboCategoria").val();
			if(categ==null){
				swal("Alerta","Debe seleccionar una Categoria..!","error");
				$("#cargando").css("display", "none");
				aCtivar();			
				return false;
			}
			let area        = $("#cboArea :selected").map((_,e) => e.value).get();
			if(area.length==0 || area==undefined ){
				area = 0;
			}else{
				area  = area.join();
			}
			let zona        = $("#cboZona :selected").map((_,e) => e.value).get();
			if(zona.length==0 || zona==undefined ){
				zona = 0;
			}else{
				zona  = zona.join();
			}
			let canal       = $("#cboCanal :selected").map((_,e) => e.value).get();
			if(canal.length==0 || canal==undefined ){
				canal = 0;
			}else{
				canal  = canal.join();
			}
			let fabricante  = $("#cboFabricante :selected").map((_,e) => e.value).get();
			if(fabricante.length==0 || fabricante==undefined ){
				fabricante = 0;
			}else{
				fabricante  = fabricante.join();
			}
			let marca       = $("#cboMarca :selected").map((_,e) => e.value).get();
			if(marca.length==0 || marca==undefined ){
				marca = 0;
			}else{
				marca  = marca.join();
			}
			let segmento    = $("#cboSegmento :selected").map((_,e) => e.value).get();
			if(segmento.length==0 || segmento==undefined ){
				segmento = 0;
			}else{
				segmento  = segmento.join();
			}
			let tamano      = $("#cboTamano :selected").map((_,e) => e.value).get();
			if(tamano.length==0 || tamano==undefined ){
				tamano = 0;
			}else{
				tamano  = tamano.join();
			}
			let producto    = $("#cboProducto :selected").map((_,e) => e.value).get();
			if(producto.length==0 || producto==undefined ){
				producto = '';
			}else{
				producto  = producto.join();
			}
			let indicadores = $("#cboIndicadores :selected").map((_,e) => e.value).get();
			if(indicadores.length==0 || indicadores==undefined ){
				indicadores = '';
			}else{
				indicadores = indicadores.join();
			}
			//			
			semanas = $("#cboSemanas :selected").map((_,e) => e.value).get();
			meses = $("#cboMeses :selected").map((_,e) => e.value).get();			
			//
			if ((semanas.length == 0 || semanas==undefined) && (meses.length == 0 || meses==undefined)) {
				swal("Alerta","Indique al menos una Semana o Mes","error");				
				$("#cargando").css("display", "none");
				aCtivar();			
				return false;
			}			
			//
			if (semanas.length == 0 || semanas==undefined) {				
				semanas  = 0;
			}else{
				semanas  = semanas.join();								
			}			
			//
			if (meses.length == 0 || meses==undefined) {
				meses = 0;
			}else{
				meses = meses.join();				
			}			
			ExecPrograma = 'RetMen_Datos.asp';
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
				mes : meses,
			};
			$('#DivRetailScanningSem').html("");
			$.ajax({
				url: ExecPrograma,
				type:'POST',
				data: ajax,
				beforeSend: function(objeto){
					$("#procesando").css("display", "block");
				}
			})
			/*Si la consulta se realizo con exito*/
			.done(function(data) {				
				debugger;
				console.log(data);
				$('#DivRetailScanningSem').html('');
				$('#DivRetailScanningSem').html(data);
				$("#procesando").css("display", "none");
				$("#DivRetailScanningSem").css("display", "block");
				sessionStorage.setItem("eXcel", 1);
				aCtivar();
			})
			/*Si la consulta Fallo*/
			.fail(function() {
				$("#procesando").css("display", "none");
				aCtivar();
				swal("Algo salio mal.!","Intente de nuevo", "error");
			},'html');

		});
		//
				
		$('#BtnExcel').click(function() {
			//
			debugger;
			event.preventDefault();
			let ExecPrograma;
			let semanas;
			bLoquear();
			//
			$("#DivRetailScanningExcel").css("display", "none");
			//
			let categ = $("#cboCategoria").val();
			if(categ==null){
				swal("Alerta","Debe seleccionar una Categoria..!","error");
				$("#cargando").css("display", "none");
				aCtivar();			
				return false;
			}
			let area        = $("#cboArea :selected").map((_,e) => e.value).get();
			if(area.length==0 || area==undefined ){
				area = 0;
			}else{
				area  = area.join();
			}
			let zona        = $("#cboZona :selected").map((_,e) => e.value).get();
			if(zona.length==0 || zona==undefined ){
				zona = 0;
			}else{
				zona  = zona.join();
			}
			let canal       = $("#cboCanal :selected").map((_,e) => e.value).get();
			if(canal.length==0 || canal==undefined ){
				canal = 0;
			}else{
				canal  = canal.join();
			}
			let fabricante  = $("#cboFabricante :selected").map((_,e) => e.value).get();
			if(fabricante.length==0 || fabricante==undefined ){
				fabricante = 0;
			}else{
				fabricante  = fabricante.join();
			}
			let marca       = $("#cboMarca :selected").map((_,e) => e.value).get();
			if(marca.length==0 || marca==undefined ){
				marca = 0;
			}else{
				marca  = marca.join();
			}
			let segmento    = $("#cboSegmento :selected").map((_,e) => e.value).get();
			if(segmento.length==0 || segmento==undefined ){
				segmento = 0;
			}else{
				segmento  = segmento.join();
			}
			let tamano      = $("#cboTamano :selected").map((_,e) => e.value).get();
			if(tamano.length==0 || tamano==undefined ){
				tamano = 0;
			}else{
				tamano  = tamano.join();
			}
			let producto    = $("#cboProducto :selected").map((_,e) => e.value).get();
			if(producto.length==0 || producto==undefined ){
				producto = '';
			}else{
				producto  = producto.join();
			}
			let indicadores = $("#cboIndicadores :selected").map((_,e) => e.value).get();
			if(indicadores.length==0 || indicadores==undefined ){
				indicadores = '';
			}else{
				indicadores = indicadores.join();
			}
			//			
			semanas = $("#cboSemanas :selected").map((_,e) => e.value).get();
			meses = $("#cboMeses :selected").map((_,e) => e.value).get();			
			//
			if ((semanas.length == 0 || semanas==undefined) && (meses.length == 0 || meses==undefined)) {
				swal("Alerta","Indique al menos una Semana o Mes","error");				
				$("#cargando").css("display", "none");
				aCtivar();			
				return false;
			}			
			//
			if (semanas.length == 0 || semanas==undefined) {				
				semanas  = 0;
			}else{
				semanas  = semanas.join();								
			}			
			//
			if (meses.length == 0 || meses==undefined) {
				meses = 0;
			}else{
				meses = meses.join();				
			}			
			ExecPrograma = 'RetMen_Excel.asp';
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
				mes : meses,
				catg: $('#cboCategoria option:selected').text(),
			};
			//$('#DivRetailScanningExcel').html("");
			$.ajax({
				url: ExecPrograma,
				type:'POST',
				data: ajax,
				beforeSend: function(objeto){
					$("#procesandoExcel").css("display", "block");
				}
			})
			/*Si la consulta se realizo con exito*/
			.done(function(data) {								
				console.log(data);
				debugger;
				$('#DivRetailScanningExcel').html('');
				$('#DivRetailScanningExcel').html(data);				
				//$("#DivRetailScanningExcel").css("display", "block");
				sessionStorage.setItem("eXcel", 1);
				aCtivar();
				//
				if(sessionStorage.getItem("eXcel")==1){
					sessionStorage.setItem("eXcel", 0);
					event.preventDefault();
					//html_table_to_excel('xlsx');
					var wb = XLSX.utils.table_to_book(document.getElementById('tbl_exportar_to_xls'), {
					  sheet: "Resultados",
					  raw: true
					});
					var wbout = XLSX.write(wb, {
					  bookType: 'xlsx',
					  bookSST: true,
					  type: 'binary'
					});
					//saveAs(new Blob([s2ab(wbout)], {  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8" }), 'Reporte_Mensual.xlsx');
					let fileName =  reemplazaTodo('Reporte Mensual '+$("#cboCategoria option:selected").text()," ","_");
					fileName+=".xlsx";					
					saveAs(new Blob([s2ab(wbout)], {  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8" }), fileName);
					
					$("#procesandoExcel").css("display", "none");
				}else{
					swal("Aviso.!","Sin Datos para procesar..!", "error");
					return false;
				}
			})
			/*Si la consulta Fallo*/
			.fail(function() {
				$("#procesando").css("display", "none");
				aCtivar();
				swal("Algo salio mal.!","Intente de nuevo", "error");
			},'html');
			
		});

		function s2ab(s) {
			var buf = new ArrayBuffer(s.length);
			var view = new Uint8Array(buf);
			for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
			return buf;
		}
		
		function reemplazaTodo(text, busca, reemplaza) {
			while (text.toString().indexOf(busca) != -1)
				text = text.toString().replace(busca, reemplaza);
			return text;
		}
		
	});

</script>


<script>
var scripts = document.getElementsByTagName('script');
var toRefreshs = ['funcionesMenV05.js', 'refillCombosMenV05.js']; // list of js to be refresh
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

<script>
// Reload all javascript / CSS
/*
$('script').each(function() {
    if ($(this).attr('src') != undefined && $(this).attr('src').lastIndexOf('jquery') == -1) {
        var old_src = $(this).attr('src');
        var that = $(this);
        $(this).attr('src', '');
        setTimeout(function() {
            that.attr('src', old_src + '?' + new Date().getTime());
        }, 250);
    }
});
*/
</script>