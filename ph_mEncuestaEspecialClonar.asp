<!DOCTYPE HTML>
<html >
<head>
	<title>| Encuesta Duplicar |</title>
	<meta charset="utf-8">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />	
	<meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
	<link href="css/sweetalert.css" rel="stylesheet" type="text/css" />	
	<link href="css/bootstrap.min.css" rel="stylesheet" type="text/css" />	
	<link href="https://unpkg.com/tabulator-tables@4.9.3/dist/css/tabulator.min.css" rel="stylesheet">
    <link href="css/tabulator_modern.min.css" rel="stylesheet" type="text/css">
    <style type="text/css">.tabulator { font-size: 12px; } #grad1 { background-color: #f3f3f3; } </style>
	</head>
<body topmargin="0">
<!--#include file="estiloscss.asp"-->
<!--#include file="encabezado.asp"-->
<!--#include file="nn_subN.asp"-->
<!--#include file="in_DataEN.asp"-->
<%
   ' 27sep21

    Apertura

	sPar=""      
    if ed_iPas <> 4 then 
        Encabezado
    end if    
	' 	
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
		
%>		
	<div class="container-fluid" id="grad1" >  		
		<br>						
		<div class="col-sm-6">
			<div class="form-group">				
				<label><i class="fas fa-copy"></i>&nbsp;Seleccione la encuesta a copiar:</label><span id="loader"></span>	
				<select class="form-control" title="Seleccionar Encuesta" name="cboEncuestas" id="cboEncuestas" />
					<option selected disabled value='0'>-- Seleccione --</option>
				</select>
			</div>
		</div>		
		<div class="col-sm-2">				
			<div class="form-group">
				<label for="usr">Procesar:</label>
				<button id="btnDuplicar" type="submit" title="Duplicar la Encuesta" class="btn btn-block btn-sm btn-success"><i class="fas fa-check fa-2x"></i></button>
			</div>				
		</div>					
		<div class="col-sm-2">				
			<div class="form-group">
				<label for="usr">Reset:</label>
				<button id="btnReset" type="submit"  title="Borrar Datos" class="btn btn-block btn-sm btn-info"><i class="fas fa-recycle fa-2x"></i></button>
			</div>				
		</div>								
		<div class="col-sm-2">				
			<div class="form-group">
				<label for="usr">Borrar:</label>
				<button id="btnEliminar" type="submit"  title="Borrar Encuesta Copiada" class="btn btn-block btn-sm btn-danger"><i class="fas fa-times fa-2x"></i></button>
			</div>				
		</div>								
	</div>
	
	<hr>
	
	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;">
		<span ><img src="images/ajax-loader7.gif"><strong>&nbsp;Procesando, espere un momento..!</strong></span>
	</div>
	
	<div class="container-fluid text-center text-danger" id="duplicar" style="display:none;">
		<span><img src="images/loader/cargador17.gif"></span><br>
		<span><strong>Espere...., Copiando..!</strong></span>
	</div>

	<div class="container-fluid" id="main" style="display:none;">  		
	
		<div class="form-group text-primary" id="totales" style="display:none;">
			<div class="col-sm-2">
				<button id="download-xlsx" class="btn btn-block btn-primary"
					title="Exportar a Excel"><i class="fas fa-download"></i>&nbsp;Excel
				</button>		
			</div>
			<div class="col-sm-2">
				<button id="clearFilter" class="btn btn-block btn-default"
					title="Eliminar filtros"><i class="fas fa-times"></i>&nbsp;Quitar filtros
				</button>				
			</div>
			<div class="col-sm-8 text-right">
				<h4><strong><span id="total" class="text-danger"></span></strong></h4>
			</div>
		</div>		
					
		<div id="tabla-resultados">			
			<!-- // ** // -->
			<!-- Matriz de Datos Resultados -->
			<!-- // ** // -->
			...				 
		</div>				
		
	</div>  
	
	<!-- /.modal investigarConsumo -->
	<div class="modal" id="nombreSurvey" tabindex="-1" data-backdrop="static" data-keyboard="false" role="dialog" aria-labelledby="myModalLabel" data-focus-on="input:first">

		<div class="modal-dialog modal-dialog-centered"  role="document">

			<div class="modal-content">
				
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>					
					<h4 class="modal-title">Large Modal</h4>					
				</div>

				<div class="modal-body">																	
					<div class="form-group">
						 <input type="text" class="form-control input-sm" name="txtNombre" id="txtNombre" maxlength='75' placeholder ="...." />
					</div>									
				</div>
				
				<div class="modal-footer">
					<button type="button" class="btn btn-danger" data-dismiss="modal" title="Salir"><i class='fas fa-sign-out-alt'></i> Salir</button>
					<button type="button" class="btn btn-primary" title="Grabar" id="btnGrabar"><i class='fas fa-save'></i> Grabar</button>
				</div>
			</div>
			<!-- /.modal-content -->
		</div>
		<!-- /.modal-dialog -->
	</div>
    <!-- /.modal -->
	 
	<% conexion.close %>
</body>
</html>
<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="matconvivencia/js/url.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
<script src="js/bootstrap.min.js"></script>
<script src="js/sweetalert.min.js"></script>
<!--Tabulator  + Excel -->
<script src="https://unpkg.com/tabulator-tables@4.9.3/dist/js/tabulator.min.js"></script>
<script src="https://oss.sheetjs.com/sheetjs/xlsx.full.min.js"></script>

<script>
	$(document).ready(function() {
		sessionStorage.clear();
		url();	
		fillCboEncuesta();
	});
	$("#btnDuplicar").prop("disabled", true);
	$("#btnEliminar").prop("disabled", true);
</script>

<script>
	
   function fillCboEncuesta() {
		//
		$('#cargando').css('display', 'block');	
		$("#cboEncuestas").prop("disabled", true);
		//
		return $.ajax({
			  url: sessionStorage.getItem("urlApi") + "getCboDataEncuesta",
			  type: 'get',
			  success: function(response) {
				//console.log(response);			
				var len = response.data.length;		
				$("#cboEncuestas").empty();
				$("#cboEncuestas").append("<option selected disabled value='0'>-- Seleccione --</option>");
				for( var i = 0; i < len; i++){
					var id = response.data[i]['id'];
					var nombre = response.data[i]['nombre'];
					$("#cboEncuestas").append("<option value='"+id+"'>"+nombre+"</option>");
				}		
				$('#cargando').css('display', 'none');
				$("#cboEncuestas").prop("disabled", false);			
				$("#cboEncuestas").focus();
			  },
			  error: function(jqXHR, textStatus, errorThrown) {
				$('#cargando').css('display', 'none');	
				alert("Fallo () CboEncuestas");
			  }
		});
		
	}
	//	
	$("#cboEncuestas").on("change", function() {
    	event.preventDefault();	
		//		
		$('#cargando').css('display', 'block');	
		$("#cboEncuestas").prop("disabled", true);
		//
		$("#tabla-resultados").html("");
		$('#total').html('');
		$('#main').css('display', 'none');
		$('#totales').css('display', 'none');
		//		
		let idEnc = $("#cboEncuestas").val();		
		return $.ajax({
		  url: sessionStorage.getItem("urlApi")+"getShowDataEncuesta/" + idEnc +"",
		  type: 'get',
		  success: function(response) {
			console.log(response);		
			let total = response.data.length;
			let valor = Number(parseInt(total)).toLocaleString("es-ES", {minimumFractionDigits: 0});
			$('#total').html('Total preguntas: ' + valor );				
			$('#totales').css('display', 'block');						
			//
			graPhData(response.data);
		  },
		  error: function(jqXHR, textStatus, errorThrown) {
			alert("Fallo () CboEncuestas");
		  }
		});
	});
	//
	$("#btnDuplicar").click(function() {		
		event.preventDefault();			
		return false;
		$("#txtNombre").val("");		
		$("#nombreSurvey").modal("show");
		$(".modal-title").html("<i class='fas fa-edit'></i> Nombre de la encuesta:");		
	});
	//
	$("#btnReset").click(function() {
		event.preventDefault();	
		$("#tabla-resultados").html("");
		$('#total').html('');		
		$("#txtNombre").val("");	
		$('#main').css('display', 'none');
		$('#cargando').css('display', 'none');
		$('#duplicar').css('display', 'none');
		$('#totales').css('display', 'none');		
		$("#btnDuplicar").prop("disabled", true);
		$("#btnEliminar").prop("disabled", true);
		$("#cboEncuestas").prop("disabled", false);
		$("#cboEncuestas").prop("selectedIndex", 0);		
		$("#cboEncuestas").focus();
	});
	//
	$("#btnGrabar").click(function() {
		//
		event.preventDefault();
		debugger;
		$("#nombreSurvey").modal("hide");				
		let idName = $("#txtNombre").val();		
		if (idName == null || idName == "" || idName.Length == 0 || idName == undefined ) {
			swal("Aviso..!", "Debe indicar un nombre...!", "error");
			return false;	
		}
		//
		swal({
			title: "¿Deseas copiar esta Encuesta?",
			text: "",
			type: "warning",
			showCancelButton: true,
			confirmButtonColor: "#5CB85C",
			confirmButtonText: "Si",
			cancelButtonText: "No",
			closeOnConfirm: true,
			closeOnCancel: false 
			},
			function(isConfirm){
			
				if (isConfirm) {				
					//						
					$('#duplicar').css('display', 'block');	
					$("#cboEncuestas").prop("disabled", true);
					$("#btnEliminar").prop("disabled", true);
					$("#btnDuplicar").prop("disabled", true);
					//
					$("#tabla-resultados").html("");
					$('#total').html('');
					$('#main').css('display', 'none');
					$('#totales').css('display', 'none');
					//		
					let idEnc = $("#cboEncuestas").val();		
					
					return $.ajax({
						url: sessionStorage.getItem("urlApi")+"getClonarEncuesta/" + idEnc +"",
						type: 'get',
						success: function(response) {					
							console.log(response);		
							setTimeout(function () {
								$('#duplicar').css('display', 'none');;
								debugger;
								fillCboEncuesta();							
							}, 3500);												
						},
						error: function(jqXHR, textStatus, errorThrown) {
							alert("Fallo () Clonar");
						}
					});
					
					setTimeout(function () {
						$('#duplicar').css('display', 'none');
						$("#cboEncuestas").prop("disabled", false);
						$("#cboEncuestas").prop("selectedIndex", 0);		
						swal("¡Hecho!", "Encuesta copiada..!", "success");
					}, 5000);												
										
				} else {
									
					swal("¡Acción!","...Cancelada...",	"error");
				}
			});
		//
		
		
		
  });

	$("#btnEliminar").click(function() {
		swal({
		title: "¿Eliminar esta encuesta?",
		text: "No podrás deshacer este paso...",
		type: "warning",
		showCancelButton: true,
		cancelButtonText: "No",
		confirmButtonColor: "#DD6B55",
		confirmButtonText: "Si",
		closeOnConfirm: false },
		function(isConfirm){
									
			if (isConfirm) {			
			
				let idEnc = $("#cboEncuestas").val();		
				
				return $.ajax({
					url: sessionStorage.getItem("urlApi")+"getEliminarEncuesta/" + idEnc +"",
					type: 'get',
					success: function(response) {							
						console.log(response);		
						debugger;
						$("#btnReset").click();
						swal("¡Hecho!",	"Encuesta Eliminada.","success");
					},
					error: function(jqXHR, textStatus, errorThrown) {
						alert("Fallo () GenExcel");
					}
				});
			}
			
		});	
	
	});
	//
	function graPhData(jsonData) {
		//	
		//debugger;
		let table = new Tabulator('#tabla-resultados', {
			height: "100%",
			layout:"fitColumns",
			data: jsonData,
			pagination: "local",
			paginationSize: 25,
			paginationSizeSelector: [25, 50, 75, 100],
			headerFilterPlaceholder: "...",
			movableColumns: true,
			tooltips: true,
			movableRows: true,
			locale: true,
			//
			columns: [{
					title: "Orden",
					field: "Orden",
					hozAlign: "center",					
					frozen: true,
					headerFilter: true,
				},
				{
					title: "Pregunta",
					field: "Pregunta",					
					headerFilter: true,
				},
				{
					title: "Tipo Pregunta",
					field: "TipoPregunta",					
					headerFilter: true,
				},				
				{
					title: "Respuestas",
					field: "Respuesta",					
				},
				{
					title: "Respuesta Cuadro",
					field: "Respuesta_Cuadro",										
				},
				{
					title: "Salto",
					field: "Salto_cuadro",
					hozAlign: "center",					
				},
				{
					title: "Imagen",
					field: "Imagen",											
				},				
				{
					title: "Maxima Cantidad Respuesta",
					field: "Maxima_Cantidad_Respuesta",						
				},
				{
					title: "Control Porcentaje",
					field: "Control_Porcentaje",		
					hozAlign: "center",					
				},
				{
					title: "Random",
					field: "Random",										
					hozAlign: "center",
				},
				{
					title: "Otro",
					field: "Otro",
					hozAlign: "center",
				},
				{
					title: "Activo",
					field: "Activo",
					hozAlign: "center",
				},
				{
					title: "Usuario",
					field: "USR",										
				},
				{
					title: "Fecha",
					field: "Fec_Ult_Mod",										
				},				
			],
			langs: {
				"es-ar": {
					pagination: {
						page_size: "Mostrar: ",
						first: "<i class='fas fa-backward'></i>",
						first_title: "Inicio",
						last: "<i class='fas fa-forward'></i>",
						last_title: "Ultimo",
						prev: "<i class='fas fa-caret-left'></i>",
						prev_title: "Anterior",
						next: "<i class='fas fa-caret-right'></i>",
						next_title: "Siguiente",
					},
				},
			},			
			
		});		
		//
		table.setLocale('es-ar'); //set locale to spanish
		// trigger download of data.xlsx file
		document
			.getElementById('download-xlsx')
			.addEventListener('click', function() {
				table.download('xlsx', 'ListadoPreguntasEncuesta.xlsx', {
					sheetName: 'Resultados',
				});
			});
			
		document
			.getElementById('clearFilter')
			.addEventListener('click', function() {
				table.clearHeaderFilter();
			});

		$('#main').css('display', 'block');
		$('#cargando').css('display', 'none');
		$('#totales').css('display', 'block');
		$("#cboEncuestas").prop("disabled", false);
		$("#btnDuplicar").prop("disabled", false);
		$("#btnEliminar").prop("disabled", false);
		//		
	}
	//	

</script>  

 