//
// CRUDV11.JS - 11mar21 - 17abr21
//
function preformatFloat(float){
   if(!float){
      return '';
   };
   
   //Index of first comma
   const posC = float.indexOf(',');

   if(posC === -1){
      //No commas found, treat as float
      return float;
   };

   //Index of first full stop
   const posFS = float.indexOf('.');

   if(posFS === -1){
      //Uses commas and not full stops - swap them (e.g. 1,23 --> 1.23)
      return float.replace(/\,/g, '.');
   };

   //Uses both commas and full stops - ensure correct order and remove 1000s separators
   return ((posC < posFS) ? (float.replace(/\,/g,'')) : (float.replace(/\./g,'').replace(',', '.')));
};
//
function Reset_FormCrearProductos(){
	$('#cboCategoria').find('option:not(:first)').remove();
	$('#cboFabricante').find('option:not(:first)').remove();
	$('#cboMarcas').find('option:not(:first)').remove();
	$('#cboSegmento').find('option:not(:first)').remove();
	$('#cboTamano').find('option:not(:first)').remove();
	$('#cboRango').find('option:not(:first)').remove();		
	$('#cboUnidadMedida').find('option:not(:first)').remove();		
	$("#codigoBarras").html("");	
	$("#fechaCreacion").html("");	
}
//
function ActualizarPromedio(id){
	//ActProm()
	debugger; 
	//
	swal(
		{
		  title: "Desea Actualizar el Precio",
		  text: ".. con el Promedio ?",
		  type: "warning",
		  showCancelButton: true,
		  confirmButtonClass: "btn-danger",
		  confirmButtonText: "Si",
		  cancelButtonText: "No",
		  closeOnConfirm: false,
		  showLoaderOnConfirm: false,
		},
		function () {			
			//
			var myUrl			= 	"g_pPendUpdateProductosPendientesxTodos.asp";	
			let ajax = {
				idConDetalle	:	id,
				Promedio		: 	parseFloat(preformatFloat($("#Promedio").val())),
				idMoneda		:	parseInt($("#idMon_"+id).val()),
				tasaCambio		:	parseFloat(preformatFloat($("#tasa_"+id).val())),				
				cantidad 		:	parseFloat(preformatFloat($("#cant_"+id).val())),
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
				debugger;
				console.log(data);						
				$("#tabla-DetalleProductosPendientes").css("display", "block");
				buscarDetallesxProductosPendientes();				
				swal("..Aviso..", "Producto Actualizado..!", "success");
				//				
			})
			/*Si la consulta Fallo*/
			.fail(function() {
				swal("Alerta Fallo","Fallo - ActProm()","error");
				//alert("Fallo - ActProm()");
			},'HTML');
			//			
		}
	  );
}	
//
function ActualizarModa(id){
	//ActModa()
	debugger; 
	//	
	swal(
		{
		  title: "Desea Actualizar el Precio",
		  text: ".. con la Moda ?",
		  type: "warning",
		  showCancelButton: true,
		  confirmButtonClass: "btn-danger",
		  confirmButtonText: "Si",
		  cancelButtonText: "No",
		  closeOnConfirm: false,
		  showLoaderOnConfirm: false,
		},
		function () {			
			//
			var myUrl			= 	"g_pPendUpdateProductosPendientesxTodos.asp";	
			let ajax = {
				idConDetalle	:	id,
				Promedio		: 	parseFloat(preformatFloat($("#Moda").val())),
				idMoneda		:	parseInt($("#idMon_"+id).val()),
				tasaCambio		:	parseFloat(preformatFloat($("#tasa_"+id).val())),				
				cantidad 		:	parseFloat(preformatFloat($("#cant_"+id).val())),
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
				debugger;
				console.log(data);						
				$("#tabla-DetalleProductosPendientes").css("display", "block");
				buscarDetallesxProductosPendientes();				
				swal("..Aviso..", "Producto Actualizado..!", "success");
				//				
			})
			/*Si la consulta Fallo*/
			.fail(function() {
				swal("Alerta Fallo","Fallo - ActModa()","error");
				//alert("Fallo - ActProm()");
			},'HTML');
			//			
		}
	  );
}	
//
function ActualizarCantidad(id){
	//CantMan()
	debugger; 
	//
	var valor 	=	parseFloat(preformatFloat($("#cantmod_"+id).val()));
	//
	if (valor == null || valor == 0 || valor.length==0 || isNaN(valor)) {
		swal("Aviso..!", "Introduzca una Cantidad Valida", "error");
		$("#cantmod_"+id).val("");
		$("#cantmod_"+id).focus();
		return false;
	}
	swal(
		{
		  title: "Seguro desea Actualizar",
		  text: "... la Cantidad ?",
		  type: "warning",
		  showCancelButton: true,
		  confirmButtonClass: "btn-danger",
		  confirmButtonText: "Si",
		  cancelButtonText: "No",
		  closeOnConfirm: false,
		  showLoaderOnConfirm: false,
		},
		function () {			
			//
			var myUrl			= 	"g_pPendUpdateCantidadProductosPendientes.asp";	
			let ajax = {
				idConDetalle	:	id,				
				cantidad 		:	valor,
				idMoneda		:	parseInt($("#idMon_"+id).val()),
				tasaCambio		:	parseFloat(preformatFloat($("#tasa_"+id).val())),							
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
				debugger;
				console.log(data);						
				$("#tabla-DetalleProductosPendientes").css("display", "block");
				buscarDetallesxProductosPendientes();				
				swal("..Aviso..", "Cantidad Actualizada..!", "success");
				//				
			})
			/*Si la consulta Fallo*/
			.fail(function() {
				swal("Alerta Fallo","Fallo - CantMan()","error");
			},'HTML');
			//			
		}
	  );
}	
//
function ActualizarManual(id){
	//ActMan()
	debugger; 
	//
	var valor 	=	parseFloat(preformatFloat($("#valor_"+id).val()));
	//
	if (valor == null || valor == 0 || valor.length==0 || isNaN(valor)) {
		swal("Aviso..!", "Introduzca un Valor Manual Valido", "error");
		$("#valor_"+id).val("");
		$("#valor_"+id).focus();
		return false;
	}
	swal(
		{
		  title: "Desea Actualizar el Precio",
		  text: ".. con Valor Manual ?",
		  type: "warning",
		  showCancelButton: true,
		  confirmButtonClass: "btn-danger",
		  confirmButtonText: "Si",
		  cancelButtonText: "No",
		  closeOnConfirm: false,
		  showLoaderOnConfirm: false,
		},
		function () {			
			//
			var myUrl			= 	"g_pPendUpdateProductosPendientesxTodos.asp";	
			let ajax = {
				idConDetalle	:	id,
				Promedio		: 	valor,
				idMoneda		:	parseInt($("#idMon_"+id).val()),
				tasaCambio		:	parseFloat(preformatFloat($("#tasa_"+id).val())),
				cantidad 		:	parseFloat(preformatFloat($("#cant_"+id).val())),
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
				debugger;
				console.log(data);						
				$("#tabla-DetalleProductosPendientes").css("display", "block");
				buscarDetallesxProductosPendientes();				
				swal("..Aviso..", "Producto Actualizado..!", "success");
				//				
			})
			/*Si la consulta Fallo*/
			.fail(function() {
				swal("Alerta Fallo","Fallo - ActMan()","error");
			},'HTML');
			//			
		}
	  );
}
//
function CrearProductos() {
	// Crear Producto
	if(validarCreacionProductos()){		
	  	//
	  swal(
		{
		  title: "Estan Correctos todos",
		  text: ".. los Datos ?",
		  type: "warning",
		  showCancelButton: true,
		  confirmButtonClass: "btn-primary",
		  confirmButtonText: "Si",
		  cancelButtonText: "No",
		  closeOnConfirm: false,
		  showLoaderOnConfirm: false,
		},
		function () {
			//	
			//debugger;
			//
			let ajax = {
				barcode 	: $("#codigoBarras").val().trim(),
				categoria	: $("#cboCategoria").val(),
				fabricante	: $("#cboFabricante").val(),
				marca		: $("#cboMarcas").val(),
				segmento	: $("#cboSegmento").val(),
				tamano		: $("#cboTamano").val(),
				rango		: $("#cboRango").val(),
				unidad		: $("#cboUnidadMedida").val(),
				fecha		: $("#fechaCreacion").val(),
				descProducto: $("#descripcionProducto").val(),
				fragmento   : $("#fragmentacion").val(),				
			};				
			//
			$.ajax({		
				url: "g_pPendCrearProductoPendiente.asp",
				type: 'GET',
				cache: false,
				async: false,				
				data: ajax,
				beforeSend: function(){
					$("#loader").css("display", "block");
				}
			})
			/*Si la consulta se realizo con exito*/
			.done(function(data) {
				debugger;
				console.log(data);													
				$("#loader").css("display", "none");
				$("#CrearProductos").modal("hide");
				$("#categoria").html("Categoria: "+ $("#cboCategoria option:selected" ).text().trim());						
				$("#descripcion").html("Descripci&oacute;n: "+$("#descripcionProducto").val());
				Reset_FormCrearProductos();				
				swal("Aviso..!", "Producto Agregado...!", "success");				
				//				
			})
			/*Si la consulta Fallo*/
			.fail(function() {
				swal("Alerta Fallo","Fallo - CrearProd()","error");				
			},'HTML');
			
			//
		}
	  );
	}		
}
//
function validarCreacionProductos() {		
	//
	// debugger;
	//
	$("#btn-crearProd").attr("disabled", true);
	//
	var Error = 0;
	//
	let barcode = $("#codigoBarras").val().trim();
	if (barcode == null || barcode == "" || barcode.length == 0 || barcode == undefined ) {
		$("#codigobarErr").html("<span style='color:red;'>Introduzca una Codigo de barras!</span>");		
		Error++;
	}else {
		$("#codigobarErr").html("");		
	}	
	//Categoria
	let comboValor = document.getElementById("cboCategoria").selectedIndex;
	if (comboValor == null || comboValor == 0 || comboValor < 0) {
	  $("#categoriaErr").html("<span style='color:red;'>Seleccione Tipo de Categoria..!</span>");
	  Error++;
	} else {
	  $("#categoriaErr").html("");
	}
	//Fabricante
	comboValor = document.getElementById("cboFabricante").selectedIndex;
	if (comboValor == null || comboValor == 0 || comboValor < 0) {
	  $("#fabricanteErr").html("<span style='color:red;'>Seleccione un Fabricante..!</span>");
	  Error++;
	} else {
	  $("#fabricanteErr").html("");
	}
	//Marcas
	comboValor = document.getElementById("cboMarcas").selectedIndex;
	if (comboValor == null || comboValor == 0 || comboValor < 0) {
	  $("#marcaErr").html("<span style='color:red;'>Seleccione una Marca..!</span>");
	  Error++;
	} else {
	  $("#marcaErr").html("");
	}
	//Segmento
	comboValor = document.getElementById("cboSegmento").selectedIndex;
	if (comboValor == null || comboValor == 0 || comboValor < 0) {
	  $("#segmentoErr").html("<span style='color:red;'>Seleccione un Segmento..!</span>");
	  Error++;
	} else {
	  $("#segmentoErr").html("");
	}
	//Tamaño
	comboValor = document.getElementById("cboTamano").selectedIndex;
	if (comboValor == null || comboValor == 0 || comboValor < 0) {
	  $("#tamanoErr").html("<span style='color:red;'>Seleccione un Tamano..!</span>");
	  Error++;
	} else {
	  $("#tamanoErr").html("");
	}
	//Rango
	comboValor = document.getElementById("cboRango").selectedIndex;
	if (comboValor == null || comboValor == 0 || comboValor < 0) {
	  $("#rangoErr").html("<span style='color:red;'>Seleccione un Rango..!</span>");
	  Error++;
	} else {
	  $("#rangoErr").html("");
	}
	//Unidad Medida
	comboValor = document.getElementById("cboUnidadMedida").selectedIndex;
	if (comboValor == null || comboValor == 0 || comboValor < 0) {
	  $("#unidadErr").html("<span style='color:red;'>Seleccione una Unidad de Medida..!</span>");
	  Error++;
	} else {
	  $("#unidadErr").html("");
	}
	// Descripcion Producto
	let descripcionProducto = $("#descripcionProducto").val().trim();
	if (descripcionProducto == null || descripcionProducto == "" || descripcionProducto.length == 0 || descripcionProducto == undefined ) {
		$("#productoErr").html("<span style='color:red;'>Describa las Caracteristicas del producto..!</span>");
		Error++;
	} else {
		$("#productoErr").html("");		
	}
	// Fragmentacion
	let fragmentacion = $("#fragmentacion").val().trim();
	if (fragmentacion == null || fragmentacion == "" || fragmentacion.length == 0 || fragmentacion == undefined ) {
		$("#fragmentoErr").html("<span style='color:red;'>Describa las Caracteristicas del producto..!</span>");
		Error++;
	} else {
		$("#fragmentoErr").html("");		
	}	
	//	
	if (Error == 0) {
		$("#btn-crearProd").attr("disabled", false);
		return true;
	} else {
		$("#btn-crearProd").attr("disabled", false);
		return false;
	}
}
//
function ValidarProductosPendientes() {
	//
	let cmbProductosPendientes = document.getElementById("cboProductosPendientes").selectedIndex;
	if (cmbProductosPendientes == null || cmbProductosPendientes == 0 || cmbProductosPendientes < 0) {
		swal("Aviso..!", "Seleccione un Codigo de Barras..!", "error"); 
		return false;	  
	} 	
	//
	debugger;
	//
	let ajax = {
		barcode 	: $("#cboProductosPendientes").val(),				
	};				
	//
	$.ajax({		
		url: "g_pPendValidarProductoPendienteExista.asp",
		type: 'GET',
		cache: false,
		async: false,				
		data: ajax,
		beforeSend: function(){
			$("#loader").css("display", "block");
		}
	})
	/*Si la consulta se realizo con exito*/
	.done(function(data) {
		debugger;
		console.log(data);																	
		if(data==="True"){
			CerrarProductosPendientes();	
		}else{
			swal("Codigo Barras: "+$("#cboProductosPendientes").val(), "No existe, debe crearlo primero..!", "error");
			return false;			
		}	
		//				
	})
	/*Si la consulta Fallo*/
	.fail(function() {
		swal("Alerta Fallo","Fallo - ValProPend()","error");				
	},'HTML');	
	//
}
//
function CerrarProductosPendientes() {
	//	
	swal(
		{
		  title: "Seguro Estan Verificados",
		  text: ".. todos los Productos ?",
		  type: "warning",
		  showCancelButton: true,
		  confirmButtonClass: "btn-danger",
		  confirmButtonText: "Si",
		  cancelButtonText: "No",
		  closeOnConfirm: false,
		  showLoaderOnConfirm: false,
		},
		function () {
			//			
			let ajax = {
				barcode 	: $("#cboProductosPendientes").val(),				
			};				
			//
			$.ajax({		
				url: "g_pPendCerrarProductoPendiente.asp",
				type: 'GET',
				cache: false,
				async: false,				
				data: ajax,
				beforeSend: function(){
					$("#loader").css("display", "block");
				}
			})
			/*Si la consulta se realizo con exito*/
			.done(function(data) {
				debugger;
				console.log(data);													
				if(data==="True"){
					Reset();
					swal("..Codigo de Barras..",$("#cboProductosPendientes").val() + " VALIDADO..!","success");
					llenarComboCategoria();	
				}else{
					swal("Alerta","Fallo - CerrarProd()","error");
					return false;			
				}	
				//				
			})
			/*Si la consulta Fallo*/
			.fail(function() {
				swal("Alerta","Fallo - CerrarProd()","error");				
			},'HTML');
			
			//
		}
	  );
	//
}
//
function actualizarPrecioMasivo(){
	// ActPreMas()
	//debugger; 
	//
	let cmbProductosPendientes = document.getElementById("cboProductosPendientes").selectedIndex;
	if (cmbProductosPendientes == null || cmbProductosPendientes == 0 || cmbProductosPendientes < 0) {
		swal("Aviso..!", "Seleccione un Codigo de Barras..!", "error"); 
		return false;	  
	} 				
	//	
	var selected = [];
	var checked = 0;
	$('.data input:checked').each(function() {
		selected.push($(this).val());
		checked ++;
	});		
	if (checked == 0) {
		swal("Aviso..!", "Debe marcar al menos un Producto...!", "error"); 
		return false;
	}else{
		GetSelectedValues();
	}
	var valor 	=	parseFloat(preformatFloat($("#txtPrecioMasivo").val()));
	//
	if (valor == null || valor == 0 || valor.length==0 || isNaN(valor)) {
		swal("Aviso..!", "Introduzca un Precio Masivo Valido", "error");
		$("#txtPrecioMasivo").val("");
		$("#txtPrecioMasivo").focus();
		return false;
	}	
	var barCode =	$("#txtBarcode").val();		
	valor = Number(parseFloat(valor)).toLocaleString("es-ES", {minimumFractionDigits: 2});	
	$("#txtPrecioMasivo").html(valor);
	//
	swal(
		{
		  title: valor + " es el Precio",
		  text: "Correcto para cambio masivo ?",
		  type: "warning",
		  showCancelButton: true,
		  confirmButtonClass: "btn-danger",
		  confirmButtonText: "Si",
		  cancelButtonText: "No",
		  closeOnConfirm: false,
		  showLoaderOnConfirm: false,
		},
		function () {			
			//			
			var myUrl	= 	"g_pPendUpdatePrecioMasivoProductosPendientes.asp";	
			//
			let ajax = {
				precio		: parseFloat(preformatFloat($("#txtPrecioMasivo").val())),
				barcode 	: $("#cboProductosPendientes").val(),
				checkboxes  : $("#Hiddenfield2").val(),
			};
			//
			$.ajax({		
				url: myUrl,
				type: 'POST',
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
				debugger;
				console.log(data);
				$("#MasivoPrecios").modal("hide");
				$("#tabla-DetalleProductosPendientes").css("display", "block");
				buscarDetallesxProductosPendientes();				
				swal("..Aviso..", "Masivo Precios Actualizado..!", "success");
				//				
			})
			/*Si la consulta Fallo*/
			.fail(function() {
				swal("Alerta Fallo","Fallo - ActPreMas()","error");				
			},'HTML');
			//			
		}
	  );
}	
//
function actualizarCambioMasivoPrecio(){
	//
	//debugger;	
	//
	// Validar Pendientes Masivo al menos uno activo este Marcado
	//
	var selected = [];
	var checked = 0;
	$('.data input:checked').each(function() {
		selected.push($(this).val());
		checked ++;
	});		
	if (checked == 0) {
		swal("Aviso..!", "Debe marcar al menos un Producto...!", "error"); 
		return false;
	}else{
		GetSelectedValues();
	}
	var valor 	=	parseFloat(preformatFloat($("#txtCambioMasivoPrecio").val()));
	//
	if (valor == null || valor == 0 || valor.length==0 || isNaN(valor)) {
		swal("Aviso..!", "Introduzca un Precio Masivo Valido", "error");
		$("#txtCambioMasivoPrecio").val("");
		$("#txtCambioMasivoPrecio").focus();
		return false;
	}	
	var barCode =	$("#txtCambBarcode").val();	
	//
	swal({
	 	title: "Esta Seguro de Enviar",
		text: "seleccionados a Pendientes ?",
	  	type: "warning",
	  	showCancelButton: true,
	  	confirmButtonClass: "btn-danger",
	  	confirmButtonText: "Si",
	  	cancelButtonText: "No",
	  	closeOnConfirm: true,
	  	showLoaderOnConfirm: true,
	},
	function () {
		//
		let ajax = {							
			checkboxes  : $("#Hiddenfield2").val(),
			barcode		: barCode,			
			precio		: valor,
		};				
		//	
		$.ajax({
			url: "g_ValUpdatePendientesCambioMasivoPrecios.asp",
			type: 'POST',
			cache: false,
			async: false,
			data: ajax,
			beforeSend: function(objeto){
				$("#cargando").css("display", "block");
			},
		})
		/*Si la consulta se realizo con exito*/
		.done(function(data) {
			//debugger;
			console.log(data);				
			$("#Hiddenfield2").val("");
			$("#cargando").css("display", "none");
			swal("Aviso..!", "Cambio precios Masivo...!", "success");			
		})
		/*Si la consulta Fallo*/
		.fail(function() {
			$("#cargando").css("display", "none");
			$("#Hiddenfield2").val("");
			swal("Fallo.!","actualizarCambioMasivoPrecio()", "error");
		},'html');
		//		
	}
  );		
}
//
function GetSelectedValues() {
	//Get the checkbox values and assigned it as a comma separated string to hiddenfield
	$("#Hiddenfield2").val($("input[name=CambioMasivo]:checked").map(function () {return this.value;}).get().join(","));
	//alert($("#Hiddenfield1").val());
}
//