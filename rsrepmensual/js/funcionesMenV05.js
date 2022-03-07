//
// funcionesMenV05.js // 12jul21 - 27ene22
//
function Reset(){
	//
	$("#DivRetailScanningSem").html("");		
	$("#procesando").css("display", "none");		
	$("#prcargando").css("display", "none");										
	$("#DivRetailScanningSem").css("display", "none");
	$("#cboCategoria").multiselect("deselectAll", false);		
	$("#cboCategoria").multiselect("refresh"); 
	$("#cboArea").multiselect("deselectAll", false);
	$("#cboArea").multiselect("refresh"); 	
	$("#cboZona").multiselect("deselectAll", false);
	$("#cboZona").multiselect("refresh"); 
	$("#cboCanal").multiselect("deselectAll", false);		
	$("#cboCanal").multiselect("refresh"); 			
	$("#cboFabricante").multiselect("deselectAll", false);
	$("#cboFabricante").multiselect("refresh"); 			
	$("#cboMarca").multiselect("deselectAll", false);
	$("#cboMarca").multiselect("refresh"); 			
	$("#cboSegmento").multiselect("deselectAll", false);
	$("#cboSegmento").multiselect("refresh"); 			
	$("#cboTamano").multiselect("deselectAll", false);
	$("#cboTamano").multiselect("refresh"); 			
	$("#cboProducto").multiselect("deselectAll", false);
	$("#cboProducto").multiselect("refresh"); 
	$("#cboIndicadores").multiselect("deselectAll", false);
	$("#cboIndicadores").multiselect("refresh"); 			
	$("#cboSemanas").multiselect("deselectAll", false);
	$("#cboSemanas").multiselect("refresh"); 
	$("#cboMeses").multiselect("deselectAll", false);
	$("#cboMeses").multiselect("refresh"); 
	sessionStorage.setItem("eXcel", 0);
	sessionStorage.setItem("repCompleto", 0);	
	//	
}
//
function tipoProducto() {
	//
	//debugger;
	let idCat = $('#cboCategoria').val();
	let ajax = { opcion: 14, idCat: idCat };
	$.ajax({
		url: "RetMen_llenar_cmb1.asp",
		type: "GET",
		dataType: 'html',
		data:  ajax,		
	})
	.done (function(response, textStatus, jqXHR) {						
		//debugger;
		if( response == 'False' ){			
			$("#tipoFabricante").html("<i class='fas fa-industry'></i>&nbsp;Fabricante:");            
			console.log('TipFab?', 'Fab');
			return true;						
		}else{
			$("#tipoFabricante").html("<i class='fas fa-clinic-medical'></i>&nbsp;Laboratorio:");            
			console.log('TipFab?', 'Lab');
			return true;						
		}				
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		console.log('Fallo en: tipFab ' + ' Err? ' + errorThrown);		
		swal("Algo salio mal.!","Tipo Producto()", "error");
	});
	//	
}	
//
function ValidarCliente(){
	//	
	//debugger;
	let idCliente = sessionStorage.getItem("idCliente");
	let ajax = { opcion: 12, idCli: sessionStorage.getItem("idCliente"), };
	$.ajax({
		url: "RetMen_llenar_cmb1.asp",
		type: "GET",
		dataType: 'html',
		data:  ajax,		
	})
	.done (function(response, textStatus, jqXHR) {
		console.log('Clt? ' , response );		
		//debugger;
		if( parseInt(response) == 0 ){
			$("#cboCategoria").empty();			
			swal("Atenas Grupo Consultor","Servicio No Contratado","info");
			return false;						
		}else{
			 LlenarCategoria();
			 return true;		
		}				
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		console.log('Fallo en: valClt ' + ' Err? ' + errorThrown);		
		swal("Algo salio mal.!","LlenarCategoria()", "error");
	});
	//
}
//
// <!-- CATEGORIA -->
function LlenarCategoria() {
	//
	//debugger;
	bLoquear();
	$("#cargando").show();
	let ajax = { opcion: 1,	idCli: sessionStorage.getItem("idCliente") };	
	$.ajax({
		url: "RetMen_llenar_cmb1.asp",
		type: "GET",
		dataType: 'json',
		data:  ajax,		
	})
	.done (function(response, textStatus, jqXHR) {
		console.log(response);
		debugger;
		var len = response.data.length;
		$("#cboCategoria").multiselect('destroy');
		$("#cboCategoria").empty();
		$("#cboCategoria").append("<option selected disabled value='0'>-- Seleccione --</option>");
		for( var i = 0; i < len; i++){
			var id = response.data[i]['id'];
			var nombre = response.data[i]['nombre'];
			$("#cboCategoria").append("<option value='" + id + "'>" + nombre.trim() + "</option>");
		}		
		//
		$("#cboCategoria").multiselect({ 
			nonSelectedText: '-- Seleccione --',
			disableIfEmpty: true,
			buttonWidth: '285px',
			buttonHeight: '30px',
			includeSelectAllOption: true,
			enableFiltering: true,
			filterPlaceholder: 'Buscar...',
			includeFilterClearBtn: true,
			enableCaseInsensitiveFiltering: true,
			maxHeight: 300,
			onDropdownHide: function(event) {        								
				let optCat   = $("#cboCategoria option:selected").val();
				let idCli    = sessionStorage.getItem("idCliente");	
				$("#Cat").val(optCat);
				Reset();
				showMe('disable');
				$("#cargando").show();				
				//	
				Promise.all([
				    tipoProducto(),
					fillAllCombos(2,  optCat, "#cboArea",  0, idCli, 0),
					fillAllCombos(3,  optCat, "#cboZona",  1, idCli, 3),
					fillAllCombos(4,  optCat, "#cboCanal", 1, idCli, 4),
					fillAllCombos(5,  optCat, "#cboFabricante", 1, idCli, 5),
					fillAllCombos(6,  optCat, "#cboMarca", 1, idCli, 6),
					fillAllCombos(7,  optCat, "#cboSegmento", 1, idCli, 7),
					fillAllCombos(8,  optCat, "#cboTamano", 1, idCli, 8),
					fillAllCombos2(9, optCat, "#cboProducto", 1, idCli, 0),
					fillAllCombos(10, optCat, "#cboIndicadores", 0, idCli, 0),
					fillAllCombos(11, optCat, "#cboSemanas", 0, idCli, 0),	
					fillAllCombos(13, optCat, "#cboMeses", 1, idCli, 0),	
					]).then(() => { // try removing ajax 1 or replacing with ajax2
						//
						setTimeout(function () {
							console.log('All Ajax done with success! ' + $("#cboCategoria option:selected").text().trim() + " - " + $("#cboCategoria option:selected").val());
						}, 3000);							
						//							
					}).catch((response) => {						
						console.log('Algunos Ajax fallaron!');			
						$("#cargando").hide();
						showMe('enable');
				});	
			}
		});
		$('#cboCategoria').multiselect('rebuild');			
		$('#cboCategoria').multiselect('refresh');
		//	
		$("#cargando").hide();
		aCtivar();
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		$("#cargando").hide();		
		aCtivar();
		console.log('Fallo en: cboCat ' + ' Err? ' + errorThrown);		
		swal("Algo salio mal.!","LlenarCategoria()", "error");
	});	
}
//
// <!-- FUNCIONES -->
//
function fillAllCombos(opc, idcat, cbo, mtp, idCli, cmb) {			
	//debugger;		
	let ajax = { opcion: opc, idCat: idcat, idCli: idCli };
	//
	$.ajax({
		url: "RetMen_llenar_cmb1.asp",
		type: "GET",
		dataType: 'json',
		data:  ajax,		
	})
	.done (function(response, textStatus, jqXHR) {
		console.log(cbo);
		console.log(response);
		debugger;
		if (mtp == 0) {			
			//			
			$(cbo).empty();
			let conta=0;			
			var len = response.data.length;
			if(cbo=="#cboFabricante"){
				$(cbo).append("<option value='0'>TOTAL CATEGORIA</option>");
			}
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];				
				if(cbo=="#cboSemanas"){
					if(conta<=4){
						$(cbo).append("<option value='"+id+"' selected>"+nombre.trim()+"</option>");
						conta++;
					}else{
						$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");	
					}					
				}else{
					$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");					
				}
			}
			$(cbo).multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, maxHeight: 200, });
			$(cbo).multiselect('rebuild');	
			$(cbo).multiselect('refresh');				
			if(cbo=="#cboSemanas"){
				$("#cargando").hide();	
				showMe('enable');			
			}
			
		}else{	
			//
			$(cbo).multiselect('destroy');
			var len = response.data.length;
			$(cbo).empty();
			if(cbo=="#cboFabricante"){
				$(cbo).append("<option value='0'>TOTAL CATEGORIA</option>");
			}
			let conta=0;			
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];												
				if(cbo=="#cboSemanas"){
					if(conta<=4){
						$(cbo).append("<option value='" + id + "' selected>" + nombre.trim() + "</option>");
						conta++;
					}else{
						$(cbo).append("<option value='" + id + "'>" + nombre.trim() + "</option>");	
					}					
				}else{
					$(cbo).append("<option value='" + id + "'>" + nombre.trim() + "</option>");					
				}
			}			
			$(cbo).multiselect({ 
				nonSelectedText: '-- Seleccione --',
				disableIfEmpty: true,
  				buttonWidth: '285px',
				buttonHeight: '30px',
				includeSelectAllOption: true,
				enableFiltering: true,
				filterPlaceholder: 'Buscar...',
				includeFilterClearBtn: true,
				enableCaseInsensitiveFiltering: true,
				maxHeight: 200,
				onDropdownHide: function(event) {        			
					GetCambioCombo(cmb);					
				}
			});
			$(cbo).multiselect('rebuild');	
			$(cbo).multiselect('refresh');		
		}
		//		
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		console.log('Fallo en: ' + cbo + ' Err? ' + errorThrown);
		swal("Algo salio mal.!", cbo , "error");
	});	
}
//
function fillAllCombos2(opc, idcat, cbo, mtp, idCli, cmb) {
	//		
	let ajax = { opcion: opc, idCat: idcat, idCli: idCli };
	//
	$.ajax({
		url: "RetMen_llenar_cmb1.asp",
		type: "GET",
		dataType: 'json',
		data:  ajax,		
	})
	.done (function(response, textStatus, jqXHR) {
		console.log(cbo);
		console.log(response);
		//debugger;
		if (mtp == 0) {			
									
			$(cbo).multiselect('destroy');			
			$(cbo).empty();									
			var len = response.data.length;
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];				  
				$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
			}
			$(cbo).multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '285px', includeSelectAllOption: true, });
			$(cbo).multiselect('rebuild');			
			$(cbo).multiselect('refresh');
			
		}else{	
		
			$(cbo).multiselect('destroy');			
			$(cbo).empty();	
			let x=opc;			
			var len = response.data.length;
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				if(opc!=9){
					$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
				}else{
					$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
				}
			}
			
			$(cbo).multiselect({ 
				nonSelectedText: '-- Seleccione --',
				disableIfEmpty: true,
  				buttonWidth: '285px',
				buttonHeight: '30px',
				includeSelectAllOption: true,
				enableFiltering: true,
				filterPlaceholder: 'Buscar...',
				includeFilterClearBtn: true,
				enableCaseInsensitiveFiltering: true,
				maxHeight: 200,
				onDropdownHide: function(event) {        			
					//GetCambioCombo(cmb,jQuery(cbo).val());					
					GetCambioCombo(cmb);					
				}
			});
			$(cbo).multiselect('rebuild');			
			$(cbo).multiselect('refresh');
		}
		//
	})
	.fail (function(jqXHR, textStatus, errorThrown) {		
		console.log('Fallo: ' + cbo + ' Err? ' + errorThrown);		
		swal("Algo salio mal.!", cbo , "error");
	});	
}
//
function showMe(value){	
	//debugger;
	$("#cboCategoria").multiselect(value);
	$("#cboArea").multiselect(value);
	$("#cboZona").multiselect(value);
	$("#cboCanal").multiselect(value);		
	$("#cboFabricante").multiselect(value);
	$("#cboMarca").multiselect(value);
	$("#cboSegmento").multiselect(value);
	$("#cboTamano").multiselect(value);
	$("#cboProducto").multiselect(value);
	$("#cboIndicadores").multiselect(value);
	$("#cboSemanas").multiselect(value);
	$("#cboMeses").multiselect(value);
}
//
function bLoquear(){
	$("#cboCategoria").multiselect('disable');
	//debugger;
	$("#cboArea").multiselect('disable');
	$("#cboZona").multiselect('disable');
	$("#cboCanal").multiselect('disable');		
	$("#cboFabricante").multiselect('disable');
	$("#cboMarca").multiselect('disable');
	$("#cboSegmento").multiselect('disable');
	$("#cboTamano").multiselect('disable');
	$("#cboProducto").multiselect('disable');
	$("#cboIndicadores").multiselect('disable');
	$("#cboSemanas").multiselect('disable');
	$("#cboMeses").multiselect('disable');
	$("#BtnAplicarFiltro").prop('disabled', true);
	$("#BtnHistorico").prop('disabled', true);
	$("#BtnExcel").prop('disabled', true);
	$("#BtnBorrar").prop('disabled', true);
	
}
//
function aCtivar(){
	$("#cboCategoria").multiselect('enable');
	//debugger;
	$("#cboArea").multiselect('enable');
	$("#cboZona").multiselect('enable');
	$("#cboCanal").multiselect('enable');		
	$("#cboFabricante").multiselect('enable');
	$("#cboMarca").multiselect('enable');
	$("#cboSegmento").multiselect('enable');
	$("#cboTamano").multiselect('enable');
	$("#cboProducto").multiselect('enable');
	$("#cboIndicadores").multiselect('enable');
	$("#cboSemanas").multiselect('enable');
	$("#cboMeses").multiselect('enable');
	//debugger;
	$("#BtnAplicarFiltro").prop('disabled', false);
	$("#BtnHistorico").prop('disabled', false);
	$("#BtnExcel").prop('disabled', false);
	$("#BtnBorrar").prop('disabled', false);
}
//
//
function bLankSelect(cmb){	
	debugger;
	bLoquear();
	$("#cargando").show();
	let optCat   = $("#cboCategoria option:selected").val();
	let idCli    = sessionStorage.getItem("idCliente");
	if(cmb==4){		
		//Canal				
		Promise.all([			
			fillAllCombos(5,  optCat, "#cboFabricante", 1, idCli, 5),
			fillAllCombos(6,  optCat, "#cboMarca", 1, idCli, 6),
			fillAllCombos(7,  optCat, "#cboSegmento", 1, idCli, 7),
			fillAllCombos(8,  optCat, "#cboTamano", 1, idCli, 8),
			fillAllCombos2(9, optCat, "#cboProducto", 1, idCli, 0),
			fillAllCombos(10, optCat, "#cboIndicadores", 0, idCli, 0),
			fillAllCombos(11, optCat, "#cboSemanas", 0, idCli, 0),
			fillAllCombos(13, optCat, "#cboMeses", 1, idCli, 0),				
			]).then(() => { // try removing ajax 1 or replacing with ajax2
			//
			setTimeout(function () {
				console.log('All bLankSelect Ajax done with success! ' + $("#cboCategoria option:selected").text().trim() + " - " + $("#cboCategoria option:selected").val());				
				$("#cargando").hide();
				aCtivar();
			}, 3000);
			//
			}).catch((response) => {
			console.log('All Ajax some failed!');
			$("#cargando").hide();
			aCtivar();
		});
		} else if(cmb==5){	
		//Fab
		Promise.all([
			fillAllCombos(6,  optCat, "#cboMarca", 1, idCli, 6),
			fillAllCombos(7,  optCat, "#cboSegmento", 1, idCli, 7),
			fillAllCombos(8,  optCat, "#cboTamano", 1, idCli, 8),
			fillAllCombos2(9, optCat, "#cboProducto", 1, idCli, 0),
			fillAllCombos(10, optCat, "#cboIndicadores", 0, idCli, 0),
			fillAllCombos(11, optCat, "#cboSemanas", 0, idCli, 0),
			fillAllCombos(13, optCat, "#cboMeses", 1, idCli, 0),	
			]).then(() => { // try removing ajax 1 or replacing with ajax2
			//
			setTimeout(function () {
				console.log('All bLankSelect Ajax done with success! ' + $("#cboCategoria option:selected").text().trim() + " - " + $("#cboCategoria option:selected").val());
				$("#cargando").hide();
				aCtivar();
			}, 3000);
			//
			}).catch((response) => {
			console.log('All Ajax some failed!');
			$("#cargando").hide();
			aCtivar();
		});
		} else if(cmb==6){	
		//Marca		
		Promise.all([			
			fillAllCombos(7,  optCat, "#cboSegmento", 1, idCli, 7),
			fillAllCombos(8,  optCat, "#cboTamano", 1, idCli, 8),
			fillAllCombos2(9, optCat, "#cboProducto", 1, idCli, 0),
			fillAllCombos(10, optCat, "#cboIndicadores", 0, idCli, 0),
			fillAllCombos(11, optCat, "#cboSemanas", 0, idCli, 0),
			fillAllCombos(13, optCat, "#cboMeses", 1, idCli, 0),	
			]).then(() => { // try removing ajax 1 or replacing with ajax2
			//
			setTimeout(function () {
				console.log('All bLankSelect Ajax done with success! ' + $("#cboCategoria option:selected").text().trim() + " - " + $("#cboCategoria option:selected").val());				
				$("#cargando").hide();
				aCtivar();
			}, 3000);
			//
			}).catch((response) => {
			console.log('All Ajax some failed!');
			$("#cargando").hide();
			aCtivar();
		});
		} else if(cmb==7){	
		//Segmento				
		Promise.all([						
			fillAllCombos(8,  optCat, "#cboTamano", 1, idCli, 8),
			fillAllCombos2(9, optCat, "#cboProducto", 1, idCli, 0),
			fillAllCombos(10, optCat, "#cboIndicadores", 0, idCli, 0),
			fillAllCombos(11, optCat, "#cboSemanas", 0, idCli, 0),
			fillAllCombos(13, optCat, "#cboMeses", 1, idCli, 0),	
			]).then(() => { // try removing ajax 1 or replacing with ajax2
			//
			setTimeout(function () {
				console.log('All bLankSelect Ajax done with success! ' + $("#cboCategoria option:selected").text().trim() + " - " + $("#cboCategoria option:selected").val());				
				$("#cargando").hide();
				aCtivar();
			}, 3000);
			//
			}).catch((response) => {
			console.log('All Ajax some failed!');
			$("#cargando").hide();
			aCtivar();
		});
		} else if(cmb==8){	
		//Tamaño			
		Promise.all([									
			fillAllCombos2(9, optCat, "#cboProducto", 1, idCli, 0),
			fillAllCombos(10, optCat, "#cboIndicadores", 0, idCli, 0),
			fillAllCombos(11, optCat, "#cboSemanas", 0, idCli, 0),
			fillAllCombos(13, optCat, "#cboMeses", 1, idCli, 0),	
			]).then(() => { // try removing ajax 1 or replacing with ajax2
			//
			setTimeout(function () {
				console.log('All bLankSelect Ajax done with success! ' + $("#cboCategoria option:selected").text().trim() + " - " + $("#cboCategoria option:selected").val());				
				$("#cargando").hide();
				aCtivar();
			}, 3000);
			//
			}).catch((response) => {
			console.log('All Ajax some failed!');
			$("#cargando").hide();
			aCtivar();
		});
		} else if(cmb==9){	
		//Producto		
		Promise.all([												
			fillAllCombos(10, optCat, "#cboIndicadores", 0, idCli, 0),
			fillAllCombos(11, optCat, "#cboSemanas", 0, idCli, 0),
			fillAllCombos(13, optCat, "#cboMeses", 1, idCli, 0),	
			]).then(() => { // try removing ajax 1 or replacing with ajax2
			//
			setTimeout(function () {
				console.log('All bLankSelect Ajax done with success! ' + $("#cboCategoria option:selected").text().trim() + " - " + $("#cboCategoria option:selected").val());				
				$("#cargando").hide();
				aCtivar();
			}, 3000);
			//
			}).catch((response) => {
			console.log('All Ajax some failed!');
			$("#cargando").hide();
			aCtivar();
		});
		
	}
	sessionStorage.setItem("eXcel", 0);
	sessionStorage.setItem("repCompleto", 0);	
}
//