<!Doctype html>
<!-- PH_Cte_HomePantryRpMen.asp - 09feb22 - 09feb22 -->
<html lang="es" >
<head>
	<title>| HP Reporte Mensual |</title>
	<meta charset="UTF-8">
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="favicon.ico" rel="icon" type="image/x-icon">
	<link href="css/sweetalert.css"  rel="stylesheet" type="text/css" />
	<link href="repmensualhp/css/hprepmensual.css"  rel="stylesheet" type="text/css" />
	<link href="css/bootstrap-multiselect-0915.css" rel="stylesheet" >
	<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />
	<link href="repmensualhp/css/hptablamen.css" rel="stylesheet" type="text/css" >	
</head>

<body topmargin="0">
	
	<!--#include file="estiloscss.asp"-->
	<!--#include file="meta.asp"-->
	<!--#include file="encabezado.asp"-->
	<!--#include file="nn_subN.asp"-->
	<!--#include file="in_DataEN.asp"-->

	<%
		'
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

		<div class="container-fluid" id="main" style="visibility:hidden;">

			<div class="row">

				<!-- ZONA A -->
				<div class="col-sm-5">

					<label class="control-label col-sm-offset-1 col-sm-3 lb"><i class="fas fa-shapes"></i>&nbsp;Categoria:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select class="form-control input-sm" id="cboCategoria" name="cboCategoria" style="width: 275px;" >
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
					
					<label class="control-label col-sm-offset-1 col-sm-3 lb" for="company"><i class="fas fa-project-diagram"></i>&nbsp;Segmento:</label>
					<div class="col-sm-6 col-md-6 separa">
						<select id="cboSegmento" name="cboSegmento" class="form-control input-sm" multiple="multiple">
						  <option value="0" selected disabled >-- Seleccione -- </option>
						</select>
					</div>

				</div>
				
				<!-- Imagen -->
				<div class="col-sm-2">				
					<p class="text-center"><img class="img-fluid"  src="images/logo/LogoHomePantry.png" width="128" height="100"/></p>										
				</div>				

				<!-- ZONA B -->
				<div class="col-sm-5">
				
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
					
					<label class="control-label col-sm-offset-1 col-sm-3 lb" for="company"></label>
					<div class="col-sm-6 col-md-6 separa">						
						<div class="btn-group btn-group-md"> 
							<button type="button" id="BtnAplicarFiltro" title="Procesar Solicitud" type="submit" class="btn btn-success"><i class="fas fa-check"></i>&nbsp;Procesar</button>
							<button type="button" id="BtnExcel"  title="Exportar a Excel" type="submit" class="btn btn-primary"><i class="fas fa-download"></i>&nbsp;Excel</button>
							<button type="button" id="BtnBorrar" title="Borrar Filtros"  type="submit" class="btn btn-danger"><i class="fas fa-times"></i>&nbsp;Borrar</button>
						</div>							
					</div>	
										
				</div>
			
			</div>
			<!-- < / class="row" -->
		</div>	<!-- < / class="container-fluid" id="grad1" -->

		<div class="container-fluid text-center text-info" id="cargando" style="display:none;">
			<br>
			<span ><img src="images/ajax-loader8.gif"><strong>&nbsp;&nbsp;Espere, cargando filtros....!</strong></span>
			<!--<span ><img src="images/atenas518.gif" border="0" width="48" height="48" ><strong>&nbsp;Espere, cargando filtros....!</strong></span>-->
		</div>
		
		<div class="container-fluid text-center text-primary" id="procesando" style="display:none;">
			<br>			
			<span ><img src="images/atenas518.gif" border="0" width="48" height="48" ><strong>&nbsp;Espere, procesando....!</strong></span>
		</div>
		
		<div class="container-fluid text-center text-primary" id="procesandoExcel" style="display:none;">
			<br>			
			<span ><img src="images/atenas518.gif" border="0" width="48" height="48" ><strong>&nbsp;Espere, Generando Excel....!</strong></span>
		</div>
		
		<hr>
		
		<div class="container-fluid text-center text-primary" id="DivHomePantryMen" style="display:none;" >
			<!-- Mostrar la tabla con los resultados -->
		</div>
		
		<div class="container-fluid text-center text-primary" id="DivHomePantryExcel" style="display:none;" >
			excel
			<!-- Mostrar la tabla con los resultados -->
		</div>

	<%conexion.close%>

</body>
</html>

<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/jquery-3.1.1.min.js"></script>
<script src="js/sweetalert.min.js"></script>
<script src="js/bootstrap.min.js"></script>
<script src="repmensualhp/js/funcionesHpMen.js"></script>
<script src="repmensualhp/js/refillCombosHpMen.js"></script>
<script src="js/bootstrap-multiselect-0915.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/amcharts/3.21.15/plugins/export/libs/FileSaver.js/FileSaver.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.15.6/xlsx.full.min.js"></script>
<script src="js/jquery.blockUI.js"></script>

<script>
	
	$(function () {		
	
		sessionStorage.clear();
		sessionStorage.setItem("idCliente", <%=Session("idCliente")%>);
		sessionStorage.setItem("repCompleto", 0);
		sessionStorage.setItem("eXcel", 0);
		$("#Cliente").val(<%=Session("idCliente")%>);
		$("#cboCategoria").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '275px', includeSelectAllOption: true, });		
		$("#cboFabricante").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '275px', includeSelectAllOption: true, });
		$("#cboMarca").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '275px', includeSelectAllOption: true, });
		$("#cboSegmento").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '275px', includeSelectAllOption: true, });
		$("#cboIndicadores").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '275px', includeSelectAllOption: true, });		
		$("#cboSemanas").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '275px', includeSelectAllOption: true, maxHeight: 200, });
		$("#cboMeses").multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '275px', includeSelectAllOption: true, maxHeight: 200, });
		document.getElementById('main').style.visibility="visible";
		
	});
	
	var removeLoading = function() {							
		setTimeout(function() {			
			document.getElementById('main').style.visibility="visible";
			$.unblockUI();
		}, 1000);
	};
	
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
			$("#DivHomePantryMen").css("display", "none");	
			let categ = $("#cboCategoria").val();
			//
			if(categ==null){
				swal("Alerta","Debe seleccionar una Categoria..!","error");
				$("#cargando").css("display", "none");
				aCtivar();			
				return false;
			}
			//									
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
			let indicadores = $("#cboIndicadores :selected").map((_,e) => e.value).get();
			if(indicadores.length==0 || indicadores==undefined ){
				indicadores = '';
			}else{
				indicadores = indicadores.join();
			}
			//			
			semanas = $("#cboSemanas :selected").map((_,e) => e.value).get();
			if (semanas.length == 0 || semanas==undefined) {
				swal("Alerta","Seleccionar una Semana","error");				
				$("#cargando").css("display", "none");
				aCtivar();			
				return false;
			}
			semanas  = semanas.join();				
			ExecPrograma = 'PH_Cte_HomePantryRpSem_Datos.asp';						
			//
			let ajax = {
				cat : categ,				
				fab : fabricante,
				mar : marca,
				seg : segmento,				
				ind : indicadores,
				sem : semanas,
				cli : sessionStorage.getItem('idCliente'),
			};
			//
			$('#DivHomePantryMen').html("");
			$.ajax({
				url: ExecPrograma,
				type:'POST',
				data: ajax,
				beforeSend: function(objeto){
					$("#procesando").css("display", "block");
				}
			})
			/* Si la consulta se realizo con exito */
			.done(function(data) {				
				debugger;
				console.log(data);
				$('#DivHomePantryMen').html('');
				$('#DivHomePantryMen').html(data);
				$("#procesando").css("display", "none");
				$("#DivHomePantryMen").css("display", "block");
				sessionStorage.setItem("eXcel", 1);
				aCtivar();
			})
			/*Si la consulta Fallo*/
			.fail(function (jqXHR, textStatus, errorThrown) {
				console.log('Error BtnAplicarFiltro:  ' + errorThrown);										
				swal("Algo salio mal.!", errorThrown , "error");				
				$("#procesando").css("display", "none");
				aCtivar();						
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
			$("#DivHomePantryExcel").css("display", "block");
			//
			let categ = $("#cboCategoria").val();
			if(categ==null){
				swal("Alerta","Debe seleccionar una Categoria..!","error");
				$("#cargando").css("display", "none");
				aCtivar();			
				return false;
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
			let indicadores = $("#cboIndicadores :selected").map((_,e) => e.value).get();
			if(indicadores.length==0 || indicadores==undefined ){
				indicadores = '';
			}else{
				indicadores = indicadores.join();
			}
			//			
			semanas = $("#cboSemanas :selected").map((_,e) => e.value).get();
			if (semanas.length == 0 || semanas==undefined) {
				swal("Alerta","Seleccionar una Semana","error");				
				$("#cargando").css("display", "none");
				aCtivar();			
				return false;
			} else {
				semanas  = semanas.join();				
				ExecPrograma = 'PH_Cte_HomePantryRpSem_Excel.asp';
			}			
			//
			let ajax = {
				cat : categ,				
				fab : fabricante,
				mar : marca,
				seg : segmento,				
				ind : indicadores,
				sem : semanas,
				cli : sessionStorage.getItem('idCliente'),
				catg: $("#cboCategoria option:selected").text(),
			};

			$('#DivHomePantryExcel').html("");
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
				debugger;
				console.log(data);
				$('#DivHomePantryExcel').html('');
				$('#DivHomePantryExcel').html(data);
				$("#procesandoExcel").css("display", "none");
				$("#DivHomePantryExcel").css("display", "none");
				sessionStorage.setItem("eXcel", 1);
				aCtivar();				
				//		
				if(sessionStorage.getItem("eXcel")==1){
				    sessionStorage.setItem("eXcel", 0);
					event.preventDefault();					
					var wb = XLSX.utils.table_to_book(document.getElementById('tbl_exportar_to_xls'), {
					  sheet: "Resultados",
					  raw: true
					});
					var wbout = XLSX.write(wb, {
					  bookType: 'xlsx',
					  bookSST: true,
					  type: 'binary'
					});
					let fileName =  reemplazaTodo('Reporte mensual '+$("#cboCategoria option:selected").text()," ","_");
					fileName+=".xlsx";					
					saveAs(new Blob([s2ab(wbout)], {  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8" }), fileName);
				}else{
					swal("Aviso.!","No ha procesado Datos..!", "error");
					return false;
				}								
			})
			/*Si la consulta Fallo*/
			.fail(function() {				
				swal("Algo salio mal.!","Intente de nuevo", "error");
				$("#procesando").css("display", "none");
				aCtivar();
			},'html');		
			
		});

		function s2ab(s) {
			var buf = new ArrayBuffer(s.length);
			var view = new Uint8Array(buf);
			for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
			return buf;
		}
		
	});
	
	function reemplazaTodo(text, busca, reemplaza) {
		while (text.toString().indexOf(busca) != -1)
			text = text.toString().replace(busca, reemplaza);
		return text;
	}

</script>

<script>
var scripts = document.getElementsByTagName('script');
//console.log(scripts);
var toRefreshs = ['funcionesSemHp.js', 'refillCombosSemHp.js']; // list of js to be refresh
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
