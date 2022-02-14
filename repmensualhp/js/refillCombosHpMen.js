//
// RefillCombosMen.js - 09feb22 - 09feb22
//
function GetCambioCombo(opc) {
	//GCC-Fabricante	
	debugger;
	console.log('opcion ' + opc);	
	$("#cargando").css("display", "block");
	bLoquear();
	$("#DivHomePantryMen").html("");
	$("#DivHomePantryMen").css("display", "none");
	//
	if (opc == 2) {
		debugger;
		// Cambio en cboFabricante
		let idCli = sessionStorage.getItem("idCliente");
		let idCatg = $("#cboCategoria option:selected").val();
		let idFabr = $("#cboFabricante :selected").map((_, e) => e.value).get();
		idFabr = idFabr.join();
		if (idFabr.length == 0 || idFabr == undefined) {
			aCtivar();
			bLankSelect(opc);
			return false;
		}
		// Llenar Marca		
		let ajax = { opcion: opc, idCat: idCatg, idFab: idFabr };
		$.ajax({
			url: "PH_Cte_HomePantryRpMen_Fill_cmb2.asp",
			type: "GET",
			dataType: 'json',
			data: ajax,
		})
			.done(function (response, textStatus, jqXHR) {
				console.log(response);
				debugger;
				let cbo = "#cboMarca";
				$(cbo).multiselect('destroy');
				var len = response.data.length;
				$(cbo).empty();
				for (var i = 0; i < len; i++) {
					var id = response.data[i]['id'];
					var nombre = response.data[i]['nombre'];
					$(cbo).append("<option value='" + id + "'>" + nombre.trim() + "</option>");
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
					onDropdownHide: function (event) {
						GetCambioCombo(3, jQuery(cbo).val());
					}
				});
				// Llenar Segmento				
				opc++;
				let ajax = { opcion: opc, idCat: idCatg, idFab: idFabr };
				$.ajax({
					url: "PH_Cte_HomePantryRpMen_Fill_cmb2.asp",
					type: "GET",
					dataType: 'json',
					data: ajax,
				})
					.done(function (response, textStatus, jqXHR) {
						console.log(response);
						//debugger;
						let cbo = "#cboSegmento";
						$(cbo).multiselect('destroy');
						var len = response.data.length;
						$(cbo).empty();
						for (var i = 0; i < len; i++) {
							var id = response.data[i]['id'];
							var nombre = response.data[i]['nombre'];
							$(cbo).append("<option value='" + id + "'>" + nombre.trim() + "</option>");
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
							onDropdownHide: function (event) {
								GetCambioCombo(4, jQuery(cbo).val());
							}
						});

					})
					.fail(function (jqXHR, textStatus, errorThrown) {
						console.log('Error GCC-Segmento opc-2:  ' + cbo + ' Error: '+ errorThrown);						
						swal("Algo salio mal.!", "GCC-Segmento opc-2", "error");
						$("#cargando").css("display", "none");
						aCtivar();						
					});
				$("#cargando").css("display", "none");
				aCtivar();
			})
			.fail(function (jqXHR, textStatus, errorThrown) {
				console.log('Error GCC-Marca opc-2:  ' + cbo + ' Error: '+ errorThrown);						
				swal("Algo salio mal.!", "GCC-Marca opc-2", "error");
				$("#cargando").css("display", "none");
				aCtivar();										
			});

	} else if (opc == 3) {
		// Cambio en cboMarca
		debugger;		
		let idCli = sessionStorage.getItem("idCliente");
		let idCatg = $("#cboCategoria option:selected").val();
		let idFabr = $("#cboFabricante :selected").map((_, e) => e.value).get();
		let idMar = $("#cboMarca :selected").map((_, e) => e.value).get();		
		idFabr = idFabr.join();
		idMar = idMar.join();		
		if (idMar.length == 0 || idMar == undefined) {
			$("#cargando").css("display", "none");
			aCtivar();
			bLankSelect(opc);
			return false;
		}
		// Llenar Segmento		
		let ajax = { opcion: opc, idCat: idCatg, idFab: idFabr, idMar: idMar };
		$.ajax({
			url: "PH_Cte_HomePantryRpMen_Fill_cmb3.asp",
			type: "GET",
			dataType: 'json',
			data: ajax,
		})
			.done(function (response, textStatus, jqXHR) {
				console.log(response);
				debugger;				
				let cbo = "#cboSegmento";
				$(cbo).multiselect('destroy');
				var len = response.data.length;
				$(cbo).empty();
				for (var i = 0; i < len; i++) {
					var id = response.data[i]['id'];
					var nombre = response.data[i]['nombre'];
					var separa = nombre.split(".");
					nombre = separa[0];
					$(cbo).append("<option value='" + id + "'>" + nombre.trim() + "</option>");
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
					onDropdownHide: function (event) {
						//Enviar procedimiento nulo opcion=0
						GetCambioCombo(0, jQuery(cbo).val());
					}
				});
				$("#cargando").css("display", "none");
				aCtivar();
				//				
			})
			.fail(function (jqXHR, textStatus, errorThrown) {				
				console.log('Error GCC-Segmento opc-3:  ' + cbo + ' Error: '+ errorThrown);						
				swal("Algo salio mal.!", "Error GCC-Segmento opc-3", "error");
				$("#cargando").css("display", "none");
				aCtivar();										
				//				
			});
			//$("#cargando").css("display", "none");

	} else if (opc == 4) {
		//Segmento
		$("#cargando").css("display", "none");
		aCtivar();
		bLankSelect(opc);

	} else if (opc == 0) {
		$("#cargando").css("display", "none");
		aCtivar();
	}

}
//
