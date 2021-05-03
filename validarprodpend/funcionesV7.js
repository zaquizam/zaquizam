//
// funcionesV7.js -- 02mar21 - 17abr21
//
function Reset() {
	//		
	$("#totalProductos").html("0");
	$("#totalValidados").html("0");
	$("#totalPendientes").html("0");
	$("#tabla-DetalleProductosPendientes").html("");	
	$("#tabla-DetalleProductosPendientes").css("display", "none")
	$("#cboProductosPendientes").prop("selectedIndex", 0);
	$("#DatosProductos").css("display", "none");
	//
}
//
function Reset_Detalle() {
	//		
	$("#totalProductos").html("0");
	$("#totalValidados").html("0");
	$("#totalPendientes").html("0");
	$("#DatosProductos").css("display", "none");
	$("#tabla-DetalleProductosPendientes").html("");	
	$("#tabla-DetalleProductosPendientes").css("display", "none")		
	//
}
//
function ProcesarCodigoBarras() {
	//	
	var idCodigoBarras	= $("#cboProductosPendientes").val();
	var myUrl			= "g_pPendContarProductosPendientesxValidar.asp";
	Reset_Detalle();
    $("#cargando").css("display", "block");
	//
	$.when( 
	   $.ajax({
		url: myUrl,
		type: 'POST',
		data: {id: idCodigoBarras, status: 0, }
	   }),

	   $.ajax({
		  url: myUrl,
		  type: 'POST',
		  data: {id: idCodigoBarras, status: 1, }
	   })
	 ).done(function( data1, data2) {
	  	// data1 and data2 are arguments resolved for the testtable.php and subtable.php' ajax requests, respectively.
	 	// Each argument is an array with the following structure: [ data, statusText, jqXHR ]
		console.log(data1);
		console.log(data2);			    
		valor = Number(parseInt(data1)).toLocaleString("es-ES", {minimumFractionDigits: 0});	
		$("#totalPendientes").html("TOTAL PENDIENTES: "+valor);
		valor = Number(parseInt(data2)).toLocaleString("es-ES", {minimumFractionDigits: 0});	
		$("#totalValidados").html("TOTAL MARCADOS: "+valor);
		valor = Number(parseInt(data1)+parseInt(data2)).toLocaleString("es-ES", {minimumFractionDigits: 0});	
		$("#totalProductos").html("TOTAL PRODUCTOS: "+valor);
		$("#cargando").css("display", "none");		
		//
		buscarDatosCodigoBarras(idCodigoBarras);
		buscarDetallesxProductosPendientes();
		//		
	});	
}
//
function buscarDetallesxProductosPendientes(){
	//bDxPP()
	//debugger;
	var idCodigoBarras	= $("#cboProductosPendientes").val();
	var myUrl			= "g_pPendBuscarDetallesxProductosPendientes.asp";	
	//
	let ajax = {
		idBarcode	:	idCodigoBarras,
	};
	//
	$.ajax({		
		url: myUrl,
		type: 'GET',
		cache: false,
		async: false,
		dataType: 'HTML',
		data: ajax,
		beforeSend: function(){
			$("#cargando").css("display", "block");
		}
	})
	/*Si la consulta se realizo con exito*/
	.done(function(data) {
		//debugger;
		console.log(data);		
		$("#cargando").css("display", "none");
		$("#tabla-DetalleProductosPendientes").css("display", "block")
		$("#tabla-DetalleProductosPendientes").html(data);
		//				
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		swal("ALerta Fallo","Fallo - bDxPP()","error");
		//alert("Fallo - bDxPP()");
	},'HTML');	
}
//
function MostrarDetalleRegistro(id) {
	//	
	var idConsumoDetalle =	id;			
	//	
	$("#MostrarDetalleRegistro").modal("show");
	$(".modal-title").html("<i class='fas fa-search'></i> Mostrar Detalle del Producto");	
	//
	DetalleRegistro(idConsumoDetalle);
	//			
}
//
function DetalleRegistro(id) {	
	//debugger;			 
	$.ajax({		
		url: "g_pPendBuscarDetallexProductos.asp?idConDetalle="        + id,
		cache: false,
		async: false,
		beforeSend: function(objeto){
			$("#cargando").css("display", "block");
		},
		success: function (data) {
			//debugger;
			console.log(data);							
			//			
			var barcode  = data[0].codigobar;
			var cantidad = data[0].cantidad;
			var precio   = parseFloat(preformatFloat(data[0].precio));			
			var tasa     = parseFloat(preformatFloat(data[0].tasa));
			var total    = parseFloat(preformatFloat(data[0].total));			
			var moneda   = data[0].moneda;
			var fecha    = data[0].fecha;
			//						
			$("#txtCodigoBar").val(barcode);						
			$("#txtCantidad").val(cantidad);
			$("#txtPrecio").val(Number(precio).toLocaleString("es-ES", {minimumFractionDigits: 2}));
			$("#txtTasa").val(Number(tasa).toLocaleString("es-ES",     {minimumFractionDigits: 2}));
			$("#txtTotal").val(Number(total).toLocaleString("es-ES",   {minimumFractionDigits: 2}));
			$("#txtMoneda").val(moneda);
			$("#txtFecha").val(fecha);			
			$("#cargando").css("display", "none");			
		},
	});		
}
//
function MostrarModalMaestroProductos() {
	//	
	let cmbProductosPendientes = document.getElementById("cboProductosPendientes").selectedIndex;
	if (cmbProductosPendientes == null || cmbProductosPendientes == 0 || cmbProductosPendientes < 0) {
		swal("Aviso..!", "Seleccione un Codigo de Barras..!", "error"); 
		return false;	  
	} 	
	$("#CrearProductos").modal("show");
	$(".modal-title").html("<i class='fas fa-plus-square'></i>&nbsp;Crear Productos");
	setDate();
	$("#codigoBarras").val($("#cboProductosPendientes").val());
	//
}
//
function buscarDatosCodigoBarras(id) {	
	//debugger;			 
	$.ajax({		
		url: "g_pPendbuscarDatosCodigoBarras.asp?CodigoBarras=" + id,
		cache: false,
		async: false,
		beforeSend: function(objeto){
			$("#cargando").css("display", "block");
		},
		success: function (data) {
			//debugger;
			console.log(data);							
			//				
			var categoria   = data[0].categoria;
			var descripcion = data[0].descripcion;			
			//	
			$("#DatosProductos").css("display", "block");			
			$("#categoria").html("Categoria: "+categoria);						
			$("#descripcion").html("Descripci&oacute;n: "+descripcion);
			$("#cargando").css("display", "none");			
		},
	});		
}
//
function setDate(){    
    var today = new Date();
    var dd = today.getDate();
    var mm = today.getMonth()+1; // Enero es 0!
    var yyyy = today.getFullYear();
    if(dd<10){dd='0'+dd} 
    if(mm<10){mm='0'+mm} 
    today = dd+'-'+mm+'-'+yyyy;     
    $("#fechaCreacion").val(today);
}
//
function showMostrarMasivoPrecios() {
	//
	// debugger;
	//	
	let cmbProductosPendientes = document.getElementById("cboProductosPendientes").selectedIndex;
	if (cmbProductosPendientes == null || cmbProductosPendientes == 0 || cmbProductosPendientes < 0) {
		swal("Aviso..!", "Seleccione un Codigo de Barras..!", "error"); 
		return false;	  
	} 	
	//	
	$("#txtBarcode").val($("#cboProductosPendientes").val());
	$("#txtPrecioMasivo").val("");
	//	
	$("#MasivoPrecios").modal("show");
	$(".modal-title").html("<i class='fas fa-edit'></i> Actualizar Precios Masivo");		
	//			
}
//