//
// RefillAllCombos.js -- 15jul21
//
function GetCambioCombo( opc ){
	//GCC-Marca
	bloquearCombos(opc);
	console.log('opcion ' + opc);
	//
	//debugger;
	if(opc==2) {
		// Cambio en cboArea
		let idCli   = sessionStorage.getItem("idCliente");
		let idCatg  = $("#cboCategoria option:selected").val();			
		let idArea  = $("#cboArea :selected").map((_,e) => e.value).get();				
		if(idArea.length==0 || idArea==undefined ){
			//idArea=0;			
			return false;
		}else{
			idArea  = idArea.join(); 
		}
		// Llenar Zona
		let ajax = { opcion: opc, idCat: idCatg, idArea: idArea };
		$.ajax({
			url: "RetSem_llenar_cmb2.asp",
			type: "GET",
			dataType: 'json',
			data:  ajax,
			beforeSend: function(){
				$("#cargando").css("display", "block");
			}
		})
		.done (function(response, textStatus, jqXHR) {
			console.log(response);
			////debugger;
			$("#cargando").css("display", "none");
			let cbo="#cboZona";
			$(cbo).multiselect('destroy');
			var len = response.data.length;
			$(cbo).empty();
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
				maxHeight: 300,
				onDropdownHide: function(event) {
					GetCambioCombo(3,jQuery(cbo).val());
				}
			});
			// Llenar Canal
			//debugger;
			opc++;
			let ajax = { opcion: opc, idCat: idCatg, idArea: idArea };
			$.ajax({
				url: "RetSem_llenar_cmb2.asp",
				type: "GET",
				dataType: 'json',
				data:  ajax,
				beforeSend: function(){
					$("#cargando").css("display", "block");
				}
			})
			.done (function(response, textStatus, jqXHR) {
				console.log(response);
				////debugger;				
				let cbo="#cboCanal";
				$(cbo).multiselect('destroy');
				var len = response.data.length;
				$(cbo).empty();
				for( var i = 0; i < len; i++){
					var id = response.data[i]['id'];
					var nombre = response.data[i]['nombre'];
					$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
					maxHeight: 300,
					onDropdownHide: function(event) {
						GetCambioCombo(4,jQuery(cbo).val());
					}
				});
				// Llenar Fabricante
				opc++;
				////debugger;
				let ajax = { opcion: opc, idCat: idCatg, idArea: idArea };
				$.ajax({
					url: "RetSem_llenar_cmb2.asp",
					type: "GET",
					dataType: 'json',
					data:  ajax,
					beforeSend: function(){
						$("#cargando").css("display", "block");
					}
				})
				.done (function(response, textStatus, jqXHR) {
					console.log(response);
					$("#cargando").css("display", "none");
					//debugger;
					let cbo="#cboFabricante";
					$(cbo).multiselect('destroy');
					var len = response.data.length;
					$(cbo).empty();
					for( var i = 0; i < len; i++){
						var id = response.data[i]['id'];
						var nombre = response.data[i]['nombre'];
						$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
						maxHeight: 300,
						onDropdownHide: function(event) {
							GetCambioCombo(5,jQuery(cbo).val());
						}
					});
					// Llenar Marca
					opc++;
					//debugger;
					let ajax = { opcion: opc, idCat: idCatg, idArea: idArea };
					$.ajax({
						url: "RetSem_llenar_cmb2.asp",
						type: "GET",
						dataType: 'json',
						data:  ajax,
						beforeSend: function(){
							$("#cargando").css("display", "block");
						}
					})
					.done (function(response, textStatus, jqXHR) {
							console.log(response);
							//debugger;
							$("#cargando").css("display", "none");
							let cbo="#cboMarca";
							$(cbo).multiselect('destroy');
							var len = response.data.length;
							$(cbo).empty();
							for( var i = 0; i < len; i++){
								var id = response.data[i]['id'];
								var nombre = response.data[i]['nombre'];
								$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
								maxHeight: 300,
								onDropdownHide: function(event) {
									GetCambioCombo(6,jQuery(cbo).val());
								}
							});
							// Llenar Segmento
							opc++;
							//debugger;
							let ajax = { opcion: opc, idCat: idCatg, idArea: idArea };
							$.ajax({
								url: "RetSem_llenar_cmb2.asp",
								type: "GET",
								dataType: 'json',
								data:  ajax,
								beforeSend: function(){
									$("#cargando").css("display", "block");
								}
							})
							.done (function(response, textStatus, jqXHR) {
								console.log(response);
								//debugger;
								$("#cargando").css("display", "none");
								let cbo="#cboSegmento";
								$(cbo).multiselect('destroy');
								var len = response.data.length;
								$(cbo).empty();
								for( var i = 0; i < len; i++){
									var id = response.data[i]['id'];
									var nombre = response.data[i]['nombre'];
									$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
									maxHeight: 300,
									onDropdownHide: function(event) {
										GetCambioCombo(7,jQuery(cbo).val());
									}
								});								
								// Llenar Tamaño
								opc++;
								//debugger;
								let ajax = { opcion: opc, idCat: idCatg, idArea: idArea };
								$.ajax({
									url: "RetSem_llenar_cmb2.asp",
									type: "GET",
									dataType: 'json',
									data:  ajax,
									beforeSend: function(){
										$("#cargando").css("display", "block");
									}
								})
								.done (function(response, textStatus, jqXHR) {
									console.log(response);
									//debugger;
									$("#cargando").css("display", "none");
									let cbo="#cboTamano";
									$(cbo).multiselect('destroy');
									var len = response.data.length;
									$(cbo).empty();
									for( var i = 0; i < len; i++){
										var id = response.data[i]['id'];
										var nombre = response.data[i]['nombre'];
										$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
										maxHeight: 300,
										onDropdownHide: function(event) {
											GetCambioCombo(8,jQuery(cbo).val());
										}
									});									
									//
									// Llenar Producto
									opc++;
									//debugger;
									let ajax = { opcion: opc, idCat: idCatg, idArea: idArea };
									$.ajax({
										url: "RetSem_llenar_cmb2.asp",
										type: "GET",
										dataType: 'json',
										data:  ajax,
										beforeSend: function(){
											$("#cargando").css("display", "block");
										}
									})
									.done (function(response, textStatus, jqXHR) {
										console.log(response);
										//debugger;
										$("#cargando").css("display", "none");
										let cbo="#cboProducto";
										$(cbo).multiselect('destroy');
										var len = response.data.length;
										$(cbo).empty();
										for( var i = 0; i < len; i++){
											var id = response.data[i]['id'];
											var nombre = response.data[i]['nombre'];
											$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
											maxHeight: 300,
											onDropdownHide: function(event) {
												GetCambioCombo(9,jQuery(cbo).val());
											}
										});										
										
									})
									.fail (function(jqXHR, textStatus, errorThrown) {
										console.log('Error GCC-Producto opc-2');
										swal("Algo salio mal.!","GCC-Producto opc-2", "error");
									});
								})
								.fail (function(jqXHR, textStatus, errorThrown) {
									console.log('Error GCC-Tamaño opc-2');
									swal("Algo salio mal.!","GCC-Tamaño opc-2", "error");
								});
							})
							.fail (function(jqXHR, textStatus, errorThrown) {
								console.log('Error GCC-Marca opc-2');
								swal("Algo salio mal.!","GCC-Marca opc-2", "error");
							});
						})
						.fail (function(jqXHR, textStatus, errorThrown) {
							console.log('Error GCC-Fabricante opc-2');
							swal("Algo salio mal.!","GCC-Fabricante opc-2", "error");
						});
					})
					.fail (function(jqXHR, textStatus, errorThrown) {
						console.log('Error GCC-Canal opc-2');
						swal("Algo salio mal.!","GCC-Canal opc-2", "error");
					});
			})
			.fail (function(jqXHR, textStatus, errorThrown) {
				console.log('Error GCC-Zona opc-2');
				swal("Algo salio mal.!","GCC-Zona opc-2", "error");
			});
		})
		.fail (function(jqXHR, textStatus, errorThrown) {
			console.log('Error GCC-Area opc-2');
			swal("Algo salio mal.!","GCC-Area opc-2", "error");
		});
		
	}else if(opc==3) {
		// Cambio en cboZona
		let idCli   = sessionStorage.getItem("idCliente");
		let idCatg  = $("#cboCategoria option:selected").val();
		let idArea  = $("#cboArea :selected").map((_,e) => e.value).get();
		let idZona  = $("#cboZona :selected").map((_,e) => e.value).get(); 
		idArea  = idArea.join(); 								
		idZona  = idZona.join();
		if(idZona.length==0 || idZona==undefined){
			return false;
		}		
		// Llenar Canal
		let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona };
		$.ajax({
			url: "RetSem_llenar_cmb3.asp",
			type: "GET",
			dataType: 'json',
			data:  ajax,
			beforeSend: function(){
				$("#cargando").css("display", "block");
			}
		})
		.done (function(response, textStatus, jqXHR) {
			console.log(response);
			////debugger;
			$("#cargando").css("display", "none");
			let cbo="#cboCanal";
			$(cbo).multiselect('destroy');
			var len = response.data.length;
			$(cbo).empty();
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
				maxHeight: 300,
				onDropdownHide: function(event) {
					GetCambioCombo(4,jQuery(cbo).val());
				}
			});
			// Llenar Fabricante
			//debugger;
			opc++;
			let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona };
			$.ajax({
				url: "RetSem_llenar_cmb3.asp",
				type: "GET",
				dataType: 'json',
				data:  ajax,
				beforeSend: function(){
					$("#cargando").css("display", "block");
				}
			})
			.done (function(response, textStatus, jqXHR) {
				console.log(response);
				////debugger;
				$("#cargando").css("display", "none");
				let cbo="#cboFabricante";
				$(cbo).multiselect('destroy');
				var len = response.data.length;
				$(cbo).empty();
				for( var i = 0; i < len; i++){
					var id = response.data[i]['id'];
					var nombre = response.data[i]['nombre'];
					$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
					maxHeight: 300,
					onDropdownHide: function(event) {
						GetCambioCombo(5,jQuery(cbo).val());
					}
				});
				// Llenar Marca
				opc++;
				////debugger;
				let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona };
				$.ajax({
					url: "RetSem_llenar_cmb3.asp",
					type: "GET",
					dataType: 'json',
					data:  ajax,
					beforeSend: function(){
						$("#cargando").css("display", "block");
					}
				})
				.done (function(response, textStatus, jqXHR) {
					console.log(response);
					$("#cargando").css("display", "none");
					//debugger;
					let cbo="#cboMarca";
					$(cbo).multiselect('destroy');
					var len = response.data.length;
					$(cbo).empty();
					for( var i = 0; i < len; i++){
						var id = response.data[i]['id'];
						var nombre = response.data[i]['nombre'];
						$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
						maxHeight: 300,
						onDropdownHide: function(event) {
							GetCambioCombo(6,jQuery(cbo).val());
						}
					});
					// Llenar Segmento
					opc++;
					//debugger;
					let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona };
					$.ajax({
						url: "RetSem_llenar_cmb3.asp",
						type: "GET",
						dataType: 'json',
						data:  ajax,
						beforeSend: function(){
							$("#cargando").css("display", "block");
						}
					})
					.done (function(response, textStatus, jqXHR) {
							console.log(response);
							//debugger;
							$("#cargando").css("display", "none");
							let cbo="#cboSegmento";
							$(cbo).multiselect('destroy');
							var len = response.data.length;
							$(cbo).empty();
							for( var i = 0; i < len; i++){
								var id = response.data[i]['id'];
								var nombre = response.data[i]['nombre'];
								$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
								maxHeight: 300,
								onDropdownHide: function(event) {
									GetCambioCombo(7,jQuery(cbo).val());
								}
							});
							// Llenar Tamaño
							opc++;
							//debugger;
							let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona };
							$.ajax({
								url: "RetSem_llenar_cmb3.asp",
								type: "GET",
								dataType: 'json',
								data:  ajax,
								beforeSend: function(){
									$("#cargando").css("display", "block");
								}
							})
							.done (function(response, textStatus, jqXHR) {
								console.log(response);
								//debugger;
								$("#cargando").css("display", "none");
								let cbo="#cboTamano";
								$(cbo).multiselect('destroy');
								var len = response.data.length;
								$(cbo).empty();
								for( var i = 0; i < len; i++){
									var id = response.data[i]['id'];
									var nombre = response.data[i]['nombre'];
									$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
									maxHeight: 300,
									onDropdownHide: function(event) {
										GetCambioCombo(8,jQuery(cbo).val());
									}
								});								
								// Llenar Producto
								opc++;
								//debugger;
								let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona };
								$.ajax({
									url: "RetSem_llenar_cmb3.asp",
									type: "GET",
									dataType: 'json',
									data:  ajax,
									beforeSend: function(){
										$("#cargando").css("display", "block");
									}
								})
								.done (function(response, textStatus, jqXHR) {
									console.log(response);
									//debugger;
									$("#cargando").css("display", "none");
									let cbo="#cboProducto";
									$(cbo).multiselect('destroy');
									var len = response.data.length;
									$(cbo).empty();
									for( var i = 0; i < len; i++){
										var id = response.data[i]['id'];
										var nombre = response.data[i]['nombre'];
										$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
										maxHeight: 300,
										onDropdownHide: function(event) {
											GetCambioCombo(9,jQuery(cbo).val());
										}
									});									
									
								})
								.fail (function(jqXHR, textStatus, errorThrown) {
									console.log('Error GCC-Producto opc-3');
									swal("Algo salio mal.!","GCC-Producto opc-3", "error");
								});
							})
							.fail (function(jqXHR, textStatus, errorThrown) {
								console.log('Error GCC-Tamaño opc-3');
								swal("Algo salio mal.!","GCC-Tamaño opc-3", "error");
							});
						})
						.fail (function(jqXHR, textStatus, errorThrown) {
							console.log('Error GCC-Segmento opc-2');
							swal("Algo salio mal.!","GCC-Segmento opc-3", "error");
						});
					})
					.fail (function(jqXHR, textStatus, errorThrown) {
						console.log('Error GCC-Marca opc-3');
						swal("Algo salio mal.!","GCC-Marca opc-3", "error");
					});
			})
			.fail (function(jqXHR, textStatus, errorThrown) {
				console.log('Error GCC-Fabricante opc-3');
				swal("Algo salio mal.!","GCC-Fabricante opc-3", "error");
			});
		})
		.fail (function(jqXHR, textStatus, errorThrown) {
			console.log('Error GCC-Canal opc-3');
			swal("Algo salio mal.!","GCC-Canal opc-3", "error");
		});
	
	}else if(opc==4) {
		// Cambio en cboCanal
		let idCli   = sessionStorage.getItem("idCliente");
		let idCatg  = $("#cboCategoria option:selected").val();
		let idArea  = $("#cboArea  :selected").map((_,e) => e.value).get();
		let idZona  = $("#cboZona  :selected").map((_,e) => e.value).get(); 
		let idCanal = $("#cboCanal :selected").map((_,e) => e.value).get(); 
		idArea  = idArea.join(); 								
		idZona  = idZona.join();
		idCanal = idCanal.join();		
		if(idCanal.length==0 || idCanal==undefined){
			return false;
		}
		// Llenar Fabricante
		let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal };
		$.ajax({
			url: "RetSem_llenar_cmb4.asp",
			type: "GET",
			dataType: 'json',
			data:  ajax,
			beforeSend: function(){
				$("#cargando").css("display", "block");
			}
		})
		.done (function(response, textStatus, jqXHR) {
			console.log(response);
			////debugger;
			$("#cargando").css("display", "none");
			let cbo="#cboFabricante";
			$(cbo).multiselect('destroy');
			var len = response.data.length;
			$(cbo).empty();
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
				maxHeight: 300,
				onDropdownHide: function(event) {
					GetCambioCombo(5,jQuery(cbo).val());
				}
			});
			// Llenar Marca
			//debugger;
			opc++;
			let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal };
			$.ajax({
				url: "RetSem_llenar_cmb4.asp",
				type: "GET",
				dataType: 'json',
				data:  ajax,
				beforeSend: function(){
					$("#cargando").css("display", "block");
				}
			})
			.done (function(response, textStatus, jqXHR) {
				console.log(response);
				////debugger;
				$("#cargando").css("display", "none");
				let cbo="#cboMarca";
				$(cbo).multiselect('destroy');
				var len = response.data.length;
				$(cbo).empty();
				for( var i = 0; i < len; i++){
					var id = response.data[i]['id'];
					var nombre = response.data[i]['nombre'];
					$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
					maxHeight: 300,
					onDropdownHide: function(event) {
						GetCambioCombo(6,jQuery(cbo).val());
					}
				});
				// Llenar Segmento
				opc++;
				////debugger;
				let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal };
				$.ajax({
					url: "RetSem_llenar_cmb4.asp",
					type: "GET",
					dataType: 'json',
					data:  ajax,
					beforeSend: function(){
						$("#cargando").css("display", "block");
					}
				})
				.done (function(response, textStatus, jqXHR) {
					console.log(response);
					$("#cargando").css("display", "none");
					//debugger;
					let cbo="#cboSegmento";
					$(cbo).multiselect('destroy');
					var len = response.data.length;
					$(cbo).empty();
					for( var i = 0; i < len; i++){
						var id = response.data[i]['id'];
						var nombre = response.data[i]['nombre'];
						$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
						maxHeight: 300,
						onDropdownHide: function(event) {
							GetCambioCombo(7,jQuery(cbo).val());
						}
					});
					// Llenar Tamaño
					opc++;
					//debugger;
					let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal };
					$.ajax({
						url: "RetSem_llenar_cmb4.asp",
						type: "GET",
						dataType: 'json',
						data:  ajax,
						beforeSend: function(){
							$("#cargando").css("display", "block");
						}
					})
					.done (function(response, textStatus, jqXHR) {
							console.log(response);
							//debugger;
							$("#cargando").css("display", "none");
							let cbo="#cboTamano";
							$(cbo).multiselect('destroy');
							var len = response.data.length;
							$(cbo).empty();
							for( var i = 0; i < len; i++){
								var id = response.data[i]['id'];
								var nombre = response.data[i]['nombre'];
								$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
								maxHeight: 300,
								onDropdownHide: function(event) {
									GetCambioCombo(8,jQuery(cbo).val());
								}
							});
							// Llenar Producto
							opc++;
							//debugger;
							let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal };
							$.ajax({
								url: "RetSem_llenar_cmb4.asp",
								type: "GET",
								dataType: 'json',
								data:  ajax,
								beforeSend: function(){
									$("#cargando").css("display", "block");
								}
							})
							.done (function(response, textStatus, jqXHR) {
								console.log(response);
								//debugger;
								$("#cargando").css("display", "none");
								let cbo="#cboProducto";
								$(cbo).multiselect('destroy');
								var len = response.data.length;
								$(cbo).empty();
								for( var i = 0; i < len; i++){
									var id = response.data[i]['id'];
									var nombre = response.data[i]['nombre'];
									$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
									maxHeight: 300,
									onDropdownHide: function(event) {
										GetCambioCombo(9,jQuery(cbo).val());
									}
								});								
								
							})
							.fail (function(jqXHR, textStatus, errorThrown) {
								console.log('Error GCC-Producto opc-4');
								swal("Algo salio mal.!","GCC-Producto opc-4", "error");
							});
						})
						.fail (function(jqXHR, textStatus, errorThrown) {
							console.log('Error GCC-Tamaño opc-4');
							swal("Algo salio mal.!","GCC-Tamaño opc-4", "error");
						});
					})
					.fail (function(jqXHR, textStatus, errorThrown) {
						console.log('Error GCC-Segmento opc-4');
						swal("Algo salio mal.!","GCC-Segmento opc-4", "error");
					});
			})
			.fail (function(jqXHR, textStatus, errorThrown) {
				console.log('Error GCC-Marca opc-3');
				swal("Algo salio mal.!","GCC-Marca opc-3", "error");	
			});
		})
		.fail (function(jqXHR, textStatus, errorThrown) {
			console.log('Error GCC-Fabricante opc-3');
			swal("Algo salio mal.!","GCC-Fabricante opc-3", "error");			
		});

	}else if(opc==5) {
		// Cambio en cboFabricante
		let idCli   = sessionStorage.getItem("idCliente");
		let idCatg  = $("#cboCategoria option:selected").val();
		let idArea  = $("#cboArea  :selected").map((_,e) => e.value).get();
		let idZona  = $("#cboZona  :selected").map((_,e) => e.value).get(); 
		let idCanal = $("#cboCanal :selected").map((_,e) => e.value).get(); 
		let idFabr  = $("#cboFabricante :selected").map((_,e) => e.value).get(); 
		idArea  = idArea.join(); 								
		idZona  = idZona.join();
		idCanal = idCanal.join();		
		idFabr  = idFabr.join();				
		if(idFabr.length==0 || idFabr==undefined){
			return false;
		}
		// Llenar Marca		
		let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal, idFab: idFabr };
		$.ajax({
			url: "RetSem_llenar_cmb5.asp",
			type: "GET",
			dataType: 'json',
			data:  ajax,
			beforeSend: function(){
				$("#cargando").css("display", "block");
			}
		})
		.done (function(response, textStatus, jqXHR) {
			console.log(response);
			////debugger;
			$("#cargando").css("display", "none");
			let cbo="#cboMarca";
			$(cbo).multiselect('destroy');
			var len = response.data.length;
			$(cbo).empty();
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
				maxHeight: 300,
				onDropdownHide: function(event) {
					GetCambioCombo(6,jQuery(cbo).val());
				}
			});
			// Llenar Segmento
			////debugger;
			opc++;
			let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal, idFab: idFabr };
			$.ajax({
				url: "RetSem_llenar_cmb5.asp",
				type: "GET",
				dataType: 'json',
				data:  ajax,
				beforeSend: function(){
					$("#cargando").css("display", "block");
				}
			})
			.done (function(response, textStatus, jqXHR) {
				console.log(response);
				////debugger;
				$("#cargando").css("display", "none");
				let cbo="#cboSegmento";
				$(cbo).multiselect('destroy');
				var len = response.data.length;
				$(cbo).empty();
				for( var i = 0; i < len; i++){
					var id = response.data[i]['id'];
					var nombre = response.data[i]['nombre'];
					$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
					maxHeight: 300,
					onDropdownHide: function(event) {
						GetCambioCombo(7,jQuery(cbo).val());
					}
				});
				// Llenar Tamaño
				opc++;
				////debugger;
				let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal, idFab: idFabr };
				$.ajax({
					url: "RetSem_llenar_cmb5.asp",
					type: "GET",
					dataType: 'json',
					data:  ajax,
					beforeSend: function(){
						$("#cargando").css("display", "block");
					}
				})
				.done (function(response, textStatus, jqXHR) {
					console.log(response);
					$("#cargando").css("display", "none");
					//debugger;
					let cbo="#cboTamano";
					$(cbo).multiselect('destroy');
					var len = response.data.length;
					$(cbo).empty();
					for( var i = 0; i < len; i++){
						var id = response.data[i]['id'];
						var nombre = response.data[i]['nombre'];
						$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
						maxHeight: 300,
						onDropdownHide: function(event) {
							GetCambioCombo(8,jQuery(cbo).val());
						}
					});
					// Llenar Producto - codigo barras
					opc++;
					//debugger;
					let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal, idFab: idFabr };
					$.ajax({
						url: "RetSem_llenar_cmb5.asp",
						type: "GET",
						dataType: 'json',
						data:  ajax,
						beforeSend: function(){
							$("#cargando").css("display", "block");
						}
					})
					.done (function(response, textStatus, jqXHR) {
						console.log(response);
						//debugger;
						$("#cargando").css("display", "none");
						let cbo="#cboProducto";
						$(cbo).multiselect('destroy');
						var len = response.data.length;
						$(cbo).empty();
						for( var i = 0; i < len; i++){
							var id = response.data[i]['id'];
							var nombre = response.data[i]['nombre'];
							$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
							maxHeight: 300,
							onDropdownHide: function(event) {
								GetCambioCombo(9,jQuery(cbo).val());
							}
						});

					})
					.fail (function(jqXHR, textStatus, errorThrown) {
						console.log('Error GCC-Producto opc-5');
						swal("Algo salio mal.!","GCC-Producto opc-5", "error");
					});
				})
				.fail (function(jqXHR, textStatus, errorThrown) {
					console.log('Error GCC-Tamaño opc-5');
					swal("Algo salio mal.!","GCC-Tamaño opc-5", "error");
				});
			})
			.fail (function(jqXHR, textStatus, errorThrown) {
				console.log('Error GCC-Segmento opc-5');
				swal("Algo salio mal.!","GCC-Segmento opc-5", "error");
			});
		})
		.fail (function(jqXHR, textStatus, errorThrown) {
			console.log('Error GCC-Marca opc-5');
			swal("Algo salio mal.!","GCC-Marca opc-5", "error");
		});

	}else if(opc==6){
		// Cambio en cboMarca
		let idCli   = sessionStorage.getItem("idCliente");
		let idCatg  = $("#cboCategoria option:selected").val();
		let idArea  = $("#cboArea  :selected").map((_,e) => e.value).get();
		let idZona  = $("#cboZona  :selected").map((_,e) => e.value).get(); 
		let idCanal = $("#cboCanal :selected").map((_,e) => e.value).get(); 
		let idFabr  = $("#cboFabricante :selected").map((_,e) => e.value).get(); 
		let idMar  = $("#cboMarca :selected").map((_,e) => e.value).get(); 
		idArea  = idArea.join(); 								
		idZona  = idZona.join();
		idCanal = idCanal.join();		
		idFabr  = idFabr.join();				
		idMar  = idMar.join();				
		
		if(idMar.length==0 || idMar==undefined){
			return false;
		}
		// Llenar Segmento		
		let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal, idFab: idFabr, idMar: idMar };
		$.ajax({
			url: "RetSem_llenar_cmb6.asp",
			type: "GET",
			dataType: 'json',
			data:  ajax,
			beforeSend: function(){
				$("#cargando").css("display", "block");
			}
		})
		.done (function(response, textStatus, jqXHR) {
			console.log(response);
			////debugger;
			$("#cargando").css("display", "none");
			let cbo="#cboSegmento";
			$(cbo).multiselect('destroy');
			var len = response.data.length;
			$(cbo).empty();
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
				maxHeight: 300,
				onDropdownHide: function(event) {
					GetCambioCombo(7,jQuery(cbo).val());
				}
			});
			// Llenar Tamaño
			////debugger;
			opc++;
			let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal, idFab: idFabr, idMar: idMar };
			$.ajax({
				url: "RetSem_llenar_cmb6.asp",
				type: "GET",
				dataType: 'json',
				data:  ajax,
				beforeSend: function(){
					$("#cargando").css("display", "block");
				}
			})
			.done (function(response, textStatus, jqXHR) {
				console.log(response);
				////debugger;
				$("#cargando").css("display", "none");
				let cbo="#cboTamano";
				$(cbo).multiselect('destroy');
				var len = response.data.length;
				$(cbo).empty();
				for( var i = 0; i < len; i++){
					var id = response.data[i]['id'];
					var nombre = response.data[i]['nombre'];
					$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
					maxHeight: 300,
					onDropdownHide: function(event) {
						GetCambioCombo(8,jQuery(cbo).val());
					}
				});
				// Llenar Producto - codigo barras
				opc++;
				////debugger;
				let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal, idFab: idFabr, idMar: idMar };
				$.ajax({
					url: "RetSem_llenar_cmb6.asp",
					type: "GET",
					dataType: 'json',
					data:  ajax,
					beforeSend: function(){
						$("#cargando").css("display", "block");
					}
				})
				.done (function(response, textStatus, jqXHR) {
					console.log(response);
					$("#cargando").css("display", "none");
					//debugger;
					let cbo="#cboProducto";
					$(cbo).multiselect('destroy');
					var len = response.data.length;
					$(cbo).empty();
					for( var i = 0; i < len; i++){
						var id = response.data[i]['id'];
						var nombre = response.data[i]['nombre'];
						$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
						maxHeight: 300,
						onDropdownHide: function(event) {
							GetCambioCombo(9,jQuery(cbo).val());
						}
					});
				})
				.fail (function(jqXHR, textStatus, errorThrown) {
					console.log('Error GCC-Producto opc-6');
					swal("Algo salio mal.!","GCC-Producto opc-6", "error");
				});
			})
			.fail (function(jqXHR, textStatus, errorThrown) {
				console.log('Error GCC-Tamaño opc-6');
				swal("Algo salio mal.!","GCC-Tamaño opc-6", "error");
			});
		})
		.fail (function(jqXHR, textStatus, errorThrown) {
			console.log('Error GCC-Segmento opc-6');
			swal("Algo salio mal.!","GCC-Segmento opc-6", "error");
		});

	}else if(opc==7){
		// Cambio en cboSegmento
		let idCli   = sessionStorage.getItem("idCliente");
		let idCatg  = $("#cboCategoria option:selected").val();
		let idArea  = $("#cboArea  :selected").map((_,e) => e.value).get();
		let idZona  = $("#cboZona  :selected").map((_,e) => e.value).get(); 
		let idCanal = $("#cboCanal :selected").map((_,e) => e.value).get(); 
		let idFabr  = $("#cboFabricante :selected").map((_,e) => e.value).get(); 
		let idMar  = $("#cboMarca :selected").map((_,e) => e.value).get(); 
		let idSegm  = $("#cboSegmento :selected").map((_,e) => e.value).get(); 
		idArea  = idArea.join(); 								
		idZona  = idZona.join();
		idCanal = idCanal.join();		
		idFabr  = idFabr.join();				
		idMar   = idMar.join();				
		idSegm  = idSegm.join();		
		if(idSegm.length==0 || idSegm==undefined){
			return false;
		}
		// Llenar Tamaño		
		let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal, idFab: idFabr, idMar: idMar, idSeg: idSegm };
		$.ajax({
			url: "RetSem_llenar_cmb7.asp",
			type: "GET",
			dataType: 'json',
			data:  ajax,
			beforeSend: function(){
				$("#cargando").css("display", "block");
			}
		})
		.done (function(response, textStatus, jqXHR) {
			console.log(response);
			////debugger;
			$("#cargando").css("display", "none");
			let cbo="#cboTamano";
			$(cbo).multiselect('destroy');
			var len = response.data.length;
			$(cbo).empty();
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
				maxHeight: 300,
				onDropdownHide: function(event) {
					GetCambioCombo(8,jQuery(cbo).val());
				}
			});
			// Llenar Producto - codigo barras
			////debugger;
			opc++;
			let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal, idFab: idFabr, idMar: idMar, idSeg: idSegm };
			$.ajax({
				url: "RetSem_llenar_cmb7.asp",
				type: "GET",
				dataType: 'json',
				data:  ajax,
				beforeSend: function(){
					$("#cargando").css("display", "block");
				}
			})
			.done (function(response, textStatus, jqXHR) {
				console.log(response);
				////debugger;
				$("#cargando").css("display", "none");
				let cbo="#cboProducto";
				$(cbo).multiselect('destroy');
				var len = response.data.length;
				$(cbo).empty();
				for( var i = 0; i < len; i++){
					var id = response.data[i]['id'];
					var nombre = response.data[i]['nombre'];
					$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
					maxHeight: 300,
					onDropdownHide: function(event) {
						GetCambioCombo(9,jQuery(cbo).val());
					}
				});

			})
			.fail (function(jqXHR, textStatus, errorThrown) {
				console.log('Error GCC-Producto opc-7');
				swal("Algo salio mal.!","GCC-Producto opc-7", "error");
			});
		})
		.fail (function(jqXHR, textStatus, errorThrown) {
			console.log('Error GCC-Tamaño opc-7');
			swal("Algo salio mal.!","GCC-Tamaño opc-7", "error");
		});
	}else if(opc==8){
		// Cambio en Tamaño
		let idCli   = sessionStorage.getItem("idCliente");
		let idCatg  = $("#cboCategoria option:selected").val();
		let idArea  = $("#cboArea  :selected").map((_,e) => e.value).get();
		let idZona  = $("#cboZona  :selected").map((_,e) => e.value).get(); 
		let idCanal = $("#cboCanal :selected").map((_,e) => e.value).get(); 
		let idFabr  = $("#cboFabricante :selected").map((_,e) => e.value).get(); 
		let idMar  = $("#cboMarca :selected").map((_,e) => e.value).get(); 
		let idSegm  = $("#cboSegmento :selected").map((_,e) => e.value).get(); 
		let idTama  = $("#cboTamano :selected").map((_,e) => e.value).get(); 
		idArea  = idArea.join(); 								
		idZona  = idZona.join();
		idCanal = idCanal.join();		
		idFabr  = idFabr.join();				
		idMar   = idMar.join();				
		idSegm  = idSegm.join();		
		idTama  = idTama.join();				
		if(idTama.length==0 || idTama==undefined){
			return false;
		}
		// Llenar Producto		
		let ajax = { opcion: opc, idCat: idCatg, idArea: idArea, idZona: idZona, idCanal: idCanal, idFab: idFabr, idMar: idMar, idSeg: idSegm, idTam: idTama };
		$.ajax({
			url: "RetSem_llenar_cmb8.asp",
			type: "GET",
			dataType: 'json',
			data:  ajax,
			beforeSend: function(){
				$("#cargando").css("display", "block");
			}
		})
		.done (function(response, textStatus, jqXHR) {
			console.log(response);
			////debugger;
			$("#cargando").css("display", "none");
			let cbo="#cboProducto";
			$(cbo).multiselect('destroy');
			var len = response.data.length;
			$(cbo).empty();
			for( var i = 0; i < len; i++){
				var id = response.data[i]['id'];
				var nombre = response.data[i]['nombre'];
				$(cbo).append("<option value='"+id+"'>"+nombre.trim()+"</option>");
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
				maxHeight: 300,
				onDropdownHide: function(event) {
					GetCambioCombo(9,jQuery(cbo).val());
				}
			});

		})
		.fail (function(jqXHR, textStatus, errorThrown) {
			console.log('Error GCC-Producto opc-8');
			swal("Algo salio mal.!","GCC-Producto opc-8", "error");
		});
	}

}
//
function bloquearCombos(opc){
	//
	$("#DivRetailScanningSem").css("display", "none");		
	$("#DivRetailScanningSem").html("");		
	//
	if(opc==2){
		$("#cboZona").find('option:not(:first)').remove();
		$("#cboCanal").find('option:not(:first)').remove();
		$("#cboFabricante").find('option:not(:first)').remove();
		$("#cboMarca").find('option:not(:first)').remove();
		$("#cboSegmento").find('option:not(:first)').remove();
		$("#cboTamano").find('option:not(:first)').remove();
		$("#cboProducto").find('option:not(:first)').remove();
		//
		$('#cboCanal').prop('disabled', true);		
		$('#cboZona').prop('disabled', true);				
		$('#cboFabricante').prop('disabled', true);		
		$('#cboMarca').prop('disabled', true);		
		$('#cboSegmento').prop('disabled', true);		
		$('#cboTamano').prop('disabled', true);		
		$('#cboProducto').prop('disabled', true);
	}else if(opc==3){		
		$("#cboCanal").find('option:not(:first)').remove();
		$("#cboCanal").prop("selectedIndex", 0);
		$("#cboFabricante").find('option:not(:first)').remove();
		$("#cboFabricante").prop("selectedIndex", 0);
		$("#cboMarca").find('option:not(:first)').remove();
		$("#cboMarca").prop("selectedIndex", 0);
		$("#cboSegmento").find('option:not(:first)').remove();
		$("#cboSegmento").prop("selectedIndex", 0);
		$("#cboTamano").find('option:not(:first)').remove();
		$("#cboTamano").prop("selectedIndex", 0);
		$("#cboProducto").find('option:not(:first)').remove();
		$("#cboProducto").prop("selectedIndex", 0);	
	}else if(opc==4){				
		$("#cboFabricante").find('option:not(:first)').remove();
		$("#cboFabricante").prop("selectedIndex", 0);
		$("#cboMarca").find('option:not(:first)').remove();
		$("#cboMarca").prop("selectedIndex", 0);
		$("#cboSegmento").find('option:not(:first)').remove();
		$("#cboSegmento").prop("selectedIndex", 0);
		$("#cboTamano").find('option:not(:first)').remove();
		$("#cboTamano").prop("selectedIndex", 0);
		$("#cboProducto").find('option:not(:first)').remove();
		$("#cboProducto").prop("selectedIndex", 0);	
	}else if(opc==5){						
		$("#cboMarca").find('option:not(:first)').remove();
		$("#cboMarca").prop("selectedIndex", 0);
		$("#cboSegmento").find('option:not(:first)').remove();
		$("#cboSegmento").prop("selectedIndex", 0);
		$("#cboTamano").find('option:not(:first)').remove();
		$("#cboTamano").prop("selectedIndex", 0);
		$("#cboProducto").find('option:not(:first)').remove();
		$("#cboProducto").prop("selectedIndex", 0);	
	} else if(opc==6){
		$("#cboSegmento").find('option:not(:first)').remove();
		$("#cboSegmento").prop("selectedIndex", 0);
		$("#cboTamano").find('option:not(:first)').remove();
		$("#cboTamano").prop("selectedIndex", 0);
		$("#cboProducto").find('option:not(:first)').remove();
		$("#cboProducto").prop("selectedIndex", 0);
	} else if(opc==7){
		$("#cboTamano").find('option:not(:first)').remove();
		$("#cboTamano").prop("selectedIndex", 0);
		$("#cboProducto").find('option:not(:first)').remove();
		$("#cboProducto").prop("selectedIndex", 0);
	} else if(opc==8){
		$("#cboProducto").find('option:not(:first)').remove();
		$("#cboProducto").prop("selectedIndex", 0);
	}
}
//