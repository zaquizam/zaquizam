<!DOCTYPE HTML>
<html >
<head>
	<title>| Encuesta Resultados |</title>
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
' 01sep21

    Apertura

	sPar=""      
    if ed_iPas <> 4 then 
        Encabezado
    end if    
	' 	
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	
	Dim arrEncuesta
	Dim rsx1
	Set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 1 'adLockOptimistic 

	QrySql = vbnullstring
    QrySql = QrySql & " SELECT"
    QrySql = QrySql & " PH_EncuestaEspecial.Id_EncuestaEspecial AS id,"
    QrySql = QrySql & " PH_EncuestaEspecial.EncuestaEspecial AS nombre"
    QrySql = QrySql & " FROM"
    QrySql = QrySql & " PH_EncuestaEspecial"
    ' QrySql = QrySql & " WHERE"
    ' QrySql = QrySql & " PH_EncuestaEspecial.Ind_Activo = 1"
    QrySql = QrySql & " ORDER BY"
    QrySql = QrySql & " PH_EncuestaEspecial.Id_EncuestaEspecial DESC"
	'
    rsx1.Open QrySql, conexion
	'	
	iExiste = 0
	if rsx1.eof then
		iExiste = 0
	else
		arrEncuesta = rsx1.GetRows()
		rsx1.close
		iExiste = 1
	end if
	
%>		
	<div class="container-fluid" id="grad1" >  		
		<br>						
		<div class="col-sm-5">
			<div class="form-group">				
				<label>Seleccione la Encuesta a Revisar:</label><span id="loader"></span>	
				<select class="form-control input-sm" title="Seleccionar Encuesta" name="cboEncuestas" id="cboEncuestas"  onchange="genExcel();" />
					<option value="0" select>-- Seleccione --</option> 
					<%							
					for iReg = 0 to ubound(arrEncuesta,2)								
						Response.write "<option value=" &  arrEncuesta(0,iReg) &">" & arrEncuesta(1,iReg) & "</option>"
					next
					%>
				</select>
			</div>
		</div>		
		<div class="col-sm-2">				
			<div class="form-group">
				<label for="usr">Borrar:</label>
				<button id="borrar" type="submit" class="btn btn-block btn-sm btn-danger" onclick="Reset();"><i class="fas fa-times"></i></button>
			</div>				
		</div>								
	</div>
	<hr>
	
	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;">
		<span ><img src="images/ajax-loader7.gif"><strong>&nbsp;Procesando, espere un momento..!</strong></span>
	</div>

	<div class="container-fluid" id="main" style="display:none;">  		
	
		<div class="form-group text-center text-primary" id="totales" style="display:none;">
			<h4><strong><span id="total" class="text-right text-primary"></span></strong></h4>
			<h4><strong><span id="totalRegistros" class="text-right text-info"></span></strong></h4>
			
			<button id="download-xlsx" class="btn btn-primary btn-sm"
				title="Exportar a Excel"><i class="fas fa-download"></i>&nbsp;Excel
			</button>		
			<button id="clearFilter" class="btn btn-warning btn-sm"
				title="Eliminar filtros"><i class="fas fa-times"></i>&nbsp;Quitar filtros
			</button>				
		</div>		
					
		<div id="tabla-resultados">			
			<!-- // ** // -->
			<!-- Matriz de Datos Resultados -->
			<!-- // ** // -->
			...				 
		</div>				
		
	</div>        
	 
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
	});
</script>

<script>
		
	function genHogares() {		
		$("#cargando").css("display", "block");		
		let idEncuesta = $("#cboEncuestas").val();
		let ajax = { id_Opcion: 1,	id_Encuesta: $("#cboEncuestas").val() };
		//
		$.ajax({
			url: "ph_mEncuestaEspecialResultadosData.asp",
			type: "GET",			
			data:  ajax,
			beforeSend: function(){
				$("#cargando").css("display", "block");
			}
		})
		.done (function(response, textStatus, jqXHR) {
			console.log(response);
			let total = response;
			if (total == 0){				
				Reset();		
				swal("Aviso..!", "No hay Datos Disponibles para esta Encuesta..!", "info"); 				
				//$('#totales').css('display', 'block');			
				//$('#total').html('Total Hogares Participantes: ' + parseInt(total).toLocaleString("es-ES", { minimumFractionDigits: 0 } ) );				
				//$('#totales').css('display', 'block');
			}else{				
				let valor = Number(parseInt(total)).toLocaleString("es-ES", {minimumFractionDigits: 0});
				$('#total').html('Total Hogares Participantes: ' + valor );								
				$('#totales').css('display', 'block');
			}			
		})
		.fail (function(jqXHR, textStatus, errorThrown) {
			swal("Algo salio mal.!","genHogares()", "error");
		})
		.always (function(jqXHROrData, textStatus, jqXHROrErrorThrown) {
			$("#cargando").css("display", "none");			
		});
		//		
	}			
	//		
	function Reset() {
		//				
		$("#tabla-resultados").html('');
		$('#total').html('');
		$('#totalRegistros').html('');
		//
		let doc = document;			
		let mySelect = doc.getElementById('cboEncuestas');
		mySelect.selectedIndex = 0;		
		$("#cboEncuestas").focus();
		//		
		$('#main').css('display', 'none');
		$('#cargando').css('display', 'none');
		$('#totales').css('display', 'none');
		$("#cboEncuestas").prop("disabled", false);		
	}
	//
	function genExcel(){				
		//debugger;
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
		  url: sessionStorage.getItem("urlApi")+"getReporteResultadosEncuestaExcel/" + idEnc +"",
		  type: 'get',
		  success: function(response) {
			//console.log(response);					
			graPhData(response.data);
		  },
		  error: function(jqXHR, textStatus, errorThrown) {
			alert("Fallo () GenExcel");
		  }
		});
	}
	//	
	function graPhData(jsonData) {
		//		
		let valor = Number(parseInt(jsonData.length)).toLocaleString("es-ES", {minimumFractionDigits: 0});
		$('#totalRegistros').html('Total Respuestas: ' + valor );						
		//
		const table = new Tabulator('#tabla-resultados', {
			//height: "640px",
			//layout: "fitDataStretch",
			layout:"fitColumns",
			data: jsonData,
			pagination: "local",
			paginationSize: 25,
			paginationSizeSelector: [25, 50, 75, 100, 500, 1000],
			headerFilterPlaceholder: "...",
			movableColumns: true,
			tooltips: true,
			movableRows: true,
			locale: true,
			//
			columns: [{
					title: "Area",
					field: "Area",
					frozen: true,
					headerFilter: true,
				},
				{
					title: "Estado",
					field: "Estado",
					frozen: true,
					headerFilter: true,
				},
				{
					title: "Hogar",
					field: "Id_Hogar",
					hozAlign: "center",
					headerFilter: true,
				},				
				{
					title: "Código",
					field: "CodigoHogar",
					headerFilter: true,
				},
				{
					title: "Nombre",
					field: "Nombre",					
					headerFilter: true,
				},
				{
					title: "Apellido",
					field: "Apellido",					
					headerFilter: true,
				},
				{
					title: "Clase Social",
					field: "ClaseSocial",	
					hozAlign: "center",					
					headerFilter: true,
				},
				{
					title: "Orden",
					field: "Orden",	
					hozAlign: "center",					
					headerFilter: true,
				},
				{
					title: "Pregunta",
					field: "Pregunta",					
					headerFilter: true,
				},
				{
					title: "Sub Pregunta",
					field: "Sub_Pregunta",					
					headerFilter: true,
				},
				{
					title: "# Respuesta",
					field: "Id_Respuesta",					
					headerFilter: true,
				},
				{
					title: "Respuesta",
					field: "RespuestaTexto",					
					headerFilter: true,
				},
			],
			langs: {
				"es-es": {
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
			downloadReady:function(fileContents, blob){               
                /* XLSX content */
                const jsonContent = JSON.parse(fileContents);
                const ws = XLSX.utils.book_new();

                //Starting in the second row to avoid overriding and skipping headers
				//jsonContent.unshift({name :"Name",gender :"Gender",age :"Age",lorem :"Description"});
                jsonContent.unshift({Area:"Area" , Estado: "Estado" , Id_Hogar: "Hogar" , CodigoHogar: "Código" , Nombre: "Nombre" , Apellido: "Apellido" , ClaseSocial: "Clase Social" , Orden: "Orden", Pregunta: "Pregunta" , Sub_Pregunta: "Sub Pregunta", Id_Respuesta: "# Respuesta", RespuestaTexto: "Respuesta Texto" });
				
                const filename  = 'Resultados';
                const dataSheet = XLSX.utils.json_to_sheet(jsonContent, { skipHeader: true });

                XLSX.utils.book_append_sheet(ws, dataSheet, filename.replace('/', ''));
                XLSX.writeFile(ws, "Resultado Encuesta.xlsx",{ bookSST: true, compression: true, bookType: 'xlsx' });
                return null;
                //return blob; //must return a blob to proceed with the download, return false to abort download
            },			
			
		});		
		//
		table.setLocale('es-es'); //set locale to spanish
		table.setData(jsonData);

        //trigger download of data.xlsx file
        document.getElementById("download-xlsx").addEventListener("click", function(){
            //table.download("xlsx", "Resultado Encuesta.xlsx", {sheetName:"Resultados"});
			table.download("json", "Resultado Encuesta.xlsx", {sheetName:"Resultados"});
        });
		//	
		//		
		document
			.getElementById('clearFilter')
			.addEventListener('click', function() {
				table.clearHeaderFilter();
			});

		$('#main').css('display', 'block');
		$('#cargando').css('display', 'none');
		$("#cboEncuestas").prop("disabled", false);		
		//
		genHogares();
		//
	}
	//	

</script>  