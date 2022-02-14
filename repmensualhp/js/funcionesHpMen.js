//
// funcionesHpSem.js // 09feb22 - 09feb22
//
function Reset(){
	//
	$("#DivHomePantryMen").html("");
	$("#procesando").css("display", "none");
	$("#prcargando").css("display", "none");
	$("#DivHomePantryMen").css("display", "none");
	$("#cboCategoria").multiselect("deselectAll", false);
	$("#cboCategoria").multiselect("refresh");	
	$("#cboFabricante").multiselect("deselectAll", false);
	$("#cboFabricante").multiselect("refresh");
	$("#cboMarca").multiselect("deselectAll", false);
	$("#cboMarca").multiselect("refresh");
	$("#cboSegmento").multiselect("deselectAll", false);
	$("#cboSegmento").multiselect("refresh");	
	$("#cboIndicadores").multiselect("deselectAll", false);
	$("#cboIndicadores").multiselect("refresh");
	$("#cboSemanas").multiselect("deselectAll", false);
	$("#cboSemanas").multiselect("refresh");
	$("#cboMeses").multiselect("deselectAll", false);
	$("#cboMeses").multiselect("refresh");
	sessionStorage.setItem("eXcel", 0);
	sessionStorage.setItem("repCompleto", 0);
}
//
function ValidarCliente(){
	// debugger;
	$("#cargando").show();
	let idCliente = sessionStorage.getItem("idCliente");
	let ajax = { opcion: 12, idCli: sessionStorage.getItem("idCliente"), };
	$.ajax({
		url: "PH_Cte_HomePantryRpMen_Fill_cmb1.asp",
		type: "GET",
		dataType: 'html',
		data:  ajax,
	})
	.done (function(response, textStatus, jqXHR) {		
		if( parseInt(response) == 0 ){
			$("#cboCategoria").empty();
			$("#cargando").hide();
			swal("Atenas Grupo Consultor","Servicio No Contratado","info");
			return false;
		}else{
			LlenarCategoria();
			return true;
		}
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		swal("Algo salio mal.!","LlenarCategoria()", "error");
	});
	//
}
//
// <!-- CATEGORIA -->
function LlenarCategoria() {
	//
	// debugger;
	bLoquear();
	let ajax = { opcion: 1,	idCli: sessionStorage.getItem("idCliente") };
	//
	$.ajax({
		url: "PH_Cte_HomePantryRpMen_Fill_cmb1.asp",
		type: "GET",
		dataType: 'json',
		data:  ajax,
	})
	.done (function(response, textStatus, jqXHR) {
		console.log(response);
		//debugger;
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
			buttonWidth: '275px',
			buttonHeight: '30px',
			includeSelectAllOption: true,
			enableFiltering: true,
			filterPlaceholder: 'Buscar...',
			includeFilterClearBtn: true,
			enableCaseInsensitiveFiltering: true,
			maxHeight: 300,
			onDropdownHide: function(event) {
				debugger;
				let categ = $("#cboCategoria").val();
				if(categ == null || categ == undefined){
					swal("Alerta","Debe seleccionar una Categoria..!","error");
					$("#cargando").css("display", "none");
					aCtivar();
					//return false;
				} else {
					let optCat   = $("#cboCategoria option:selected").val();
					let idCli    = sessionStorage.getItem("idCliente");
					$("#Cat").val(optCat);
					Reset();
					showMe('disable');
					$("#cargando").show();
					//
					Promise.all([
						//tipoProducto(),												
						fillAllCombos(2,  optCat, "#cboFabricante", 1, idCli, 2),						
						fillAllCombos(3,  optCat, "#cboMarca", 1, idCli, 3),
						fillAllCombos(4,  optCat, "#cboSegmento", 1, idCli, 4),						
						fillAllCombos(5,  optCat, "#cboIndicadores", 0, idCli, 0),
						fillAllCombos(6,  optCat, "#cboSemanas", 0, idCli, 0),
						fillAllCombos(7,  optCat, "#cboMeses", 1, idCli, 0),	
						
						]).then(() => { // try removing ajax 1 or replacing with ajax2
						//
						setTimeout(function () {
							console.log('All Ajax done with success! ' + $("#cboCategoria option:selected").text().trim() + " - " + $("#cboCategoria option:selected").val());
							showMe('enable');
						}, 3000);
						//
						}).catch((response) => {
						console.log('All Ajax some failed!');
						$("#cargando").hide();
						showMe('enable');
					});
				}
			}
		});
		$('#cboCategoria').multiselect('rebuild');
		$('#cboCategoria').multiselect('refresh');
		//
		$("#cargando").hide();
		removeLoading();
		aCtivar();
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		//alert("Error " + errorThrown);
		$("#cargando").hide();
		aCtivar();
		swal("Algo salio mal.!","LlenarCategoria()", "error");
	});
}
//
function fillAllCombos(opc, idcat, cbo, mtp, idCli, cmb) {
	//debugger;
	let ajax = { opcion: opc, idCat: idcat, idCli: idCli };	
	$.ajax({
		url: "PH_Cte_HomePantryRpMen_Fill_cmb1.asp",
		type: "GET",
		dataType: 'json',
		data:  ajax,
	})
	.done (function(response, textStatus, jqXHR) {
		console.log(cbo);
		console.log(response);
		//debugger;
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
					if(i <= 4){
						$(cbo).append("<option value='"+id+"' selected>"+nombre.trim()+"</option>");
						//conta++;
					}else{
						$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
					}
				}else{
					$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
				}
			}
			$(cbo).multiselect('destroy');
			$(cbo).multiselect({ nonSelectedText: '-- Seleccione --', disableIfEmpty: true, buttonWidth: '275px', includeSelectAllOption: true, maxHeight: 200 });
			$(cbo).multiselect('rebuild');
			$(cbo).multiselect('refresh');
			if(cbo=="#cboSemanas"){
				$("#cargando").hide();
				showMe('enable');
			}			
		}else {
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
  				buttonWidth: '275px',
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
		$("#cargando").hide();
		//
	})
	.fail (function(jqXHR, textStatus, errorThrown) {
		console.log('Fallo:  ' + cbo + ' Error: '+ errorThrown);
		swal("Algo salio mal.!", cbo , "error");
	});
}
//
function showMe(value){
	$("#cboCategoria").multiselect(value);
	$("#cboFabricante").multiselect(value);
	$("#cboMarca").multiselect(value);
	$("#cboSegmento").multiselect(value);
	$("#cboIndicadores").multiselect(value);
	$("#cboSemanas").multiselect(value);
	$("#cboMeses").multiselect(value);
}
//
function bLoquear(){
	//debugger;
	$("#cboCategoria").multiselect('disable');
	$("#cboFabricante").multiselect('disable');
	$("#cboMarca").multiselect('disable');
	$("#cboSegmento").multiselect('disable');
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
	//debugger;
	$("#cboCategoria").multiselect('enable');
	$("#cboFabricante").multiselect('enable');
	$("#cboMarca").multiselect('enable');
	$("#cboSegmento").multiselect('enable');
	$("#cboIndicadores").multiselect('enable');
	$("#cboSemanas").multiselect('enable');
	$("#cboMeses").multiselect('enable');
	$("#BtnAplicarFiltro").prop('disabled', false);
	$("#BtnHistorico").prop('disabled', false);
	$("#BtnExcel").prop('disabled', false);
	$("#BtnBorrar").prop('disabled', false);
}
//
function bLankSelect(cmb){	
	//debugger;
	bLoquear();
	$("#cargando").show();
	let optCat   = $("#cboCategoria option:selected").val();
	let idCli    = sessionStorage.getItem("idCliente");
	if(cmb==2){	
		//Fab
		Promise.all([			
			fillAllCombos(3, optCat, "#cboMarca", 1, idCli, 3),
			fillAllCombos(4, optCat, "#cboSegmento", 1, idCli, 4),			
			fillAllCombos(5, optCat, "#cboIndicadores", 0, idCli, 0),					
			fillAllCombos(6, optCat, "#cboSemanas", 0, idCli, 0),
			fillAllCombos(7, optCat, "#cboMeses", 1, idCli, 0),	
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
	} else if(cmb==3){	
		//Marca		
		Promise.all([			
			fillAllCombos(4, optCat, "#cboSegmento", 1, idCli, 4),						
			fillAllCombos(5, optCat, "#cboIndicadores", 0, idCli, 0),								
			fillAllCombos(6, optCat, "#cboSemanas", 0, idCli, 0),
			fillAllCombos(7, optCat, "#cboMeses", 1, idCli, 0),	
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
	} else if(cmb==4){	
		//Segmento						
		$("#cargando").hide();
		aCtivar();
	} 
	sessionStorage.setItem("eXcel", 0);
	sessionStorage.setItem("repCompleto", 0);	
}
//