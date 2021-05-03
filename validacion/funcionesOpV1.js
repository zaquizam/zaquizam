//
// funcionesOpV1.js - 22mar21 - 
//
$("#otrosProductos").click(function() { 
	//
	if($("#otrosProductos").is(':checked')) {  
		//alert("Está activado");  
		$("#MasterProductos").css("display", "none");
		$("#cboCategoria").prop("selectedIndex", 0);
		$("#cboCategoria").attr("disabled", "disabled");
		$("#txtBuscarDescripcion").attr("disabled", "disabled");
		$("#txtBuscarDescripcion").val("");
		$("#cboProducto").prop("selectedIndex", 0);
		$("#cboProducto").attr("disabled", "disabled");
		$("#txtCodigoBarras").val("0");		
		//	
		$("#cboMonedaPago").focus();
		$("#showOtrosProductos").css("display", "block");		
		$("#cboCategoriaOtros").prop("selectedIndex", 0);
		$("#cboMarcaCubitos").prop("selectedIndex", 0);
		//
			
	} else {  
		//alert("No está activado");  
		$("#MasterProductos").css("display", "block");
		$("#cboCategoria").removeAttr("disabled");
		$("#cboCategoria").prop("selectedIndex", 0);
		$("#txtBuscarDescripcion").removeAttr("disabled");
		$("#txtBuscarDescripcion").val("");
		$("#cboProducto").prop("selectedIndex", 0);
		$("#cboProducto").removeAttr("disabled");
		$("#txtCodigoBarras").val("");
		//		
		$("#cboCategoria").focus();
		$("#showOtrosProductos").css("display", "none");
		$("#showCubitos").css("display", "none");
		$("#cboCategoriaOtros").prop("selectedIndex", 0);
		$("#cboMarcaCubitos").prop("selectedIndex", 0);
		//
		$("#nombreCantidad").html("Cantidad:");
		$("#nombrePrecio").html("Precio Unitario:"); 
	}  
}); 
//
$("#cboCategoriaOtros").on("change", function() {
	//
	event.preventDefault();
	var value = $(this).val();
	//		
	if(parseInt(value) == 5){
		// Huevos
		$("#showUnidades").css("display", "block");
		$("#showCubitos").css("display", "none");
		$("#unidades").val("");	
		$("#cantidad").val(""); 
		$("#nombreCantidad").html("Cantidad Empaque:");
		$("#nombrePrecio").html("Precio Empaque:");
		$("#cboMarcaCubitos").prop("selectedIndex", 0);
		//
	} else if(parseInt(value) == 6){
		// Botellones
		$("#showUnidades").css("display", "none");
		$("#showCubitos").css("display", "none");		
		$("#unidades").val("0");		
		$("#nombreCantidad").html("Cantidad:");
		$("#nombrePrecio").html("Precio Unitario:"); 
		$("#cboMarcaCubitos").prop("selectedIndex", 0);
		//	
	}else if(parseInt(value) == 8) {
		// Cubitos
		$("#showUnidades").css("display", "none");
		$("#showCubitos").css("display", "block");
		$("#unidades").val("0");
		$("#nombreCantidad").html("Cantidad:");
		$("#nombrePrecio").html("Precio Unitario:"); 
		$("#cboMarcaCubitos").prop("selectedIndex", 0);
		//		
	}else if(parseInt(value) == 9) {
		// Quesos
		$("#showUnidades").css("display", "none");
		$("#showCubitos").css("display", "none");
		$("#unidades").val("0");
		$("#nombreCantidad").html("Cantidad en Gramos:");
		$("#nombrePrecio").html("Precio Pagado:"); 
		$("#cboMarcaCubitos").prop("selectedIndex", 0);
		//		
	}else{
		$("#showCubitos").css("display", "none");
	}
  
});