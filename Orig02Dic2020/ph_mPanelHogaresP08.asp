<script>
	//LR

	//**Inicio Registrar Inc
	function registrarinc(){
		//alert("Llego Registrar Inc");
		num = document.getElementById("Hogar").value;
		if (num === "" ) {
			//swal("Composicion del Hogar","Debe Registrar Primero los Datos de Identificacion","error");
			alert("Debe Registrar Primero los Datos de Identificacion");
			return;
		}
		//alert("Llego Registrar Inc");
		var mensajeerr="";
		var err=0;
		nombre1 = document.getElementById("PrimerNombreInc").value;
		nombre2 = document.getElementById("SegundoNombreInc").value;
		apellido1 = document.getElementById("PrimerApellidoInc").value;
		apellido2 = document.getElementById("SegundoApellidoInc").value;
		nacionalidad = document.getElementById("NacionalidadInc").value;
		cedula = document.getElementById("CedulaInc").value;
		celular1 = document.getElementById("CelularInc").value;
		correo1 = document.getElementById("CorreoInc").value;
		parentesco = document.getElementById("ParentescoInc").value;
		fechanacimiento = document.getElementById("FechaNacimientoInc").value;
		sexo = document.getElementById("SexoInc").value;
		educacion = document.getElementById("EducacionInc").value;
		tipoingreso = document.getElementById("TipoIngresoInc").value;
		if(nombre1 == "")
		{
			mensajeerr = mensajeerr + "\nPrimer Nombre"
			err=err+1;
		}
		if(apellido1 == "")
		{
			mensajeerr = mensajeerr + "\nPrimer Apellido"
			err=err+1;
		}
		if(err > 0)
		{
			alert("Falta informacion de:" + mensajeerr);
			//swal("Falta informacion de:",mensajeerr,"error");
			return
		}
		/*nombre1 = document.getElementById("PrimerNombreInc").value;
		nombre2 = document.getElementById("SegundoNombreInc").value;
		apellido1 = document.getElementById("PrimerApellidoInc").value;
		apellido2 = document.getElementById("SegundoApellidoInc").value;
		nacionalidad = document.getElementById("NacionalidadInc").value;
		cedula = document.getElementById("CedulaInc").value;
		celular1 = document.getElementById("CelularInc").value;
		correo1 = document.getElementById("CorreoInc").value;
		parentesco = document.getElementById("ParentescoInc").value;
		fechanacimiento = document.getElementById("FechaNacimientoInc").value;
		sexo = document.getElementById("SexoInc").value;
		educacion = document.getElementById("EducacionInc").value;
		tipoingreso = document.getElementById("TipoIngresoInc").value;*/

		var stodo = "num=" + num;
		stodo = stodo + "&nom1=" + nombre1;
		stodo = stodo + "&nom2=" + nombre2;
		stodo = stodo + "&ape1=" + apellido1;
		stodo = stodo + "&ape2=" + apellido2;
		stodo = stodo + "&naci=" + nacionalidad;
		stodo = stodo + "&cedu=" + cedula;
		stodo = stodo + "&cel1=" + celular1;
		stodo = stodo + "&cor1=" + correo1;
		stodo = stodo + "&pare=" + parentesco;
		stodo = stodo + "&fech=" + fechanacimiento;
		stodo = stodo + "&sexo=" + sexo;
		stodo = stodo + "&educ=" + educacion;
		stodo = stodo + "&tipo=" + tipoingreso;
		//alert("Todo:=" + stodo);
		$.ajax({
			url:'g_GrabarBloque02inc.asp?'+stodo,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data); 
				alert("Registro Actualizado"); 
				//swal("Responsable del Panel","Registro Actualizado","success");
				//buscar_panelistas();
				document.getElementById("PrimerNombreInc").value = "";
				document.getElementById("SegundoNombreInc").value = "";
				document.getElementById("PrimerApellidoInc").value = "";
				document.getElementById("SegundoApellidoInc").value = ""; 
				document.getElementById("NacionalidadInc").value = 0;
				document.getElementById("CedulaInc").value = "";
				document.getElementById("CelularInc").value = "";
				document.getElementById("CorreoInc").value = "";
				document.getElementById("ParentescoInc").value = 0;
				document.getElementById("EstadoCivilInc").value = 0;
				document.getElementById("FechaNacimientoInc").value = "";
				document.getElementById("EdadInc").value = "";
				document.getElementById("SexoInc").value = 0;
				document.getElementById("EducacionInc").value = 0;
				document.getElementById("TipoIngresoInc").value = 0;
				buscar_panelistas();
				CierraModal();
			}
		})
	}	
	//**Fin Registrar Inc
	
	//**Inicio Buscar Edad Mod
	function buscar_edadmod(){
		//LR
		//alert("Llego Edad:=" + document.getElementById("FechaNacimientoMod").value);
		//debugger;
		//let fecha = $("#FechaNacimientoMod").val();
		let fecha = document.getElementById("FechaNacimientoMod").value
		  if (
			fecha == null ||
			fecha == "" ||
			fecha == undefined ||
			fecha.length == 0 ||
			!isNaN(fecha)
		  ) {
			$("#fechaErr").html(
			  '<span style="color:red;">Introduzca fecha valida!</span>'
			);
			//$("#FechaNacimientoMod").focus();
			document.getElementById("FechaNacimientoMod").focus();
			return false;
		  } else {
			if (!validarFormatoFecha(fecha)) {
			  //$("#fechaErr").html(
			  document.getElementById("FechaNacimientoMod").html(
				'<span style="color:red;">Introduzca fecha valida!</span>'
			  );
			  //$("#FechaNacimientoMod").focus();
			  document.getElementById("FechaNacimientoMod").focus();
			  return false;
			} else {
			  //debugger;
			  let hoy = new Date();
			  let fechaFormulario = new Date(fecha);
			  hoy.setHours(0, 0, 0, 0); // Lo iniciamos a 00:00 horas
			  if (hoy <= fechaFormulario) {
				//$("#fechaErr").html(
				document.getElementById("FechaNacimientoMod").html(
				  '<span style="color:red;">Fecha ingreso posterior a Hoy!</span>'
				);
				return false;
			  } else {
				//$("#fechaErr").html("");
				//document.getElementById("FechaNacimientoMod").html("");
			  }
			}
		  }
		  //alert("Llego Edad2");
		  // Si la fecha es correcta, calculamos la edad
		  let values = fecha.split("-");
		  let dia = values[2];
		  let mes = values[1];
		  let ano = values[0];
		  // tomamos los valores actuales
		  let fecha_hoy = new Date();
		  let ahora_ano = fecha_hoy.getYear();
		  let ahora_mes = fecha_hoy.getMonth() + 1;
		  let ahora_dia = fecha_hoy.getDate();
		  // realizamos el calculo
		  let edad = ahora_ano + 1900 - ano;
		  if (ahora_mes < mes) {
			edad--;
		  }
		  if (mes == ahora_mes && ahora_dia < dia) {
			edad--;
		  }
		  if (edad > 1900) {
			edad -= 1900;
		  }
		  // calculamos los meses
		  let meses = 0;
		  if (ahora_mes > mes && dia > ahora_dia) meses = ahora_mes - mes - 1;
		  else if (ahora_mes > mes) meses = ahora_mes - mes;
		  if (ahora_mes < mes && dia < ahora_dia) meses = 12 - (mes - ahora_mes);
		  else if (ahora_mes < mes) meses = 12 - (mes - ahora_mes + 1);
		  if (ahora_mes == mes && dia > ahora_dia) meses = 11;
		  // calculamos los dias
		  let dias = 0;
		  if (ahora_dia > dia) dias = ahora_dia - dia;
		  if (ahora_dia < dia) {
			ultimoDiaMes = new Date(ahora_ano, ahora_mes - 1, 0);
			dias = ultimoDiaMes.getDate() - (dia - ahora_dia);
		  }
		  //let tiempo = edad + " años, " + meses + " meses y " + dias + " días";
		  let tiempo = edad + " años ";
		  //edad + " años";
		  //$("#Edad").val(tiempo);
		  document.getElementById("EdadMod").value = tiempo;

		  return; //edad + " años, " + meses + " meses y " + dias + " días";
	}	
	//**Fin Buscar Edad Inc
	
	//**Inicio Buscar Edad Inc
	function buscar_edadinc(){
		//LR
		//alert("Llego Edad:=" + document.getElementById("FechaNacimientoInc").value);
		//debugger;
		//let fecha = $("#FechaNacimiento").val();
		let fecha = document.getElementById("FechaNacimientoInc").value
		  if (
			fecha == null ||
			fecha == "" ||
			fecha == undefined ||
			fecha.length == 0 ||
			!isNaN(fecha)
		  ) {
			$("#fechaErr").html(
			  '<span style="color:red;">Introduzca fecha valida!</span>'
			);
			//$("#FechaNacimientoInc").focus();
			document.getElementById("FechaNacimientoInc").focus();
			return false;
		  } else {
			if (!validarFormatoFecha(fecha)) {
			  //$("#fechaErr").html(
			  document.getElementById("FechaNacimientoInc").html(
				'<span style="color:red;">Introduzca fecha valida!</span>'
			  );
			  //$("#FechaNacimiento").focus();
			  document.getElementById("FechaNacimientoInc").focus();
			  return false;
			} else {
			  //debugger;
			  let hoy = new Date();
			  let fechaFormulario = new Date(fecha);
			  hoy.setHours(0, 0, 0, 0); // Lo iniciamos a 00:00 horas
			  if (hoy <= fechaFormulario) {
				//$("#fechaErr").html(
				document.getElementById("FechaNacimientoInc").html(
				  '<span style="color:red;">Fecha ingreso posterior a Hoy!</span>'
				);
				return false;
			  } else {
				//$("#fechaErr").html("");
				//document.getElementById("FechaNacimientoInc").html("");
			  }
			}
		  }
		  //alert("Llego Edad2");
		  // Si la fecha es correcta, calculamos la edad
		  let values = fecha.split("-");
		  let dia = values[2];
		  let mes = values[1];
		  let ano = values[0];
		  // tomamos los valores actuales
		  let fecha_hoy = new Date();
		  let ahora_ano = fecha_hoy.getYear();
		  let ahora_mes = fecha_hoy.getMonth() + 1;
		  let ahora_dia = fecha_hoy.getDate();
		  // realizamos el calculo
		  let edad = ahora_ano + 1900 - ano;
		  if (ahora_mes < mes) {
			edad--;
		  }
		  if (mes == ahora_mes && ahora_dia < dia) {
			edad--;
		  }
		  if (edad > 1900) {
			edad -= 1900;
		  }
		  // calculamos los meses
		  let meses = 0;
		  if (ahora_mes > mes && dia > ahora_dia) meses = ahora_mes - mes - 1;
		  else if (ahora_mes > mes) meses = ahora_mes - mes;
		  if (ahora_mes < mes && dia < ahora_dia) meses = 12 - (mes - ahora_mes);
		  else if (ahora_mes < mes) meses = 12 - (mes - ahora_mes + 1);
		  if (ahora_mes == mes && dia > ahora_dia) meses = 11;
		  // calculamos los dias
		  let dias = 0;
		  if (ahora_dia > dia) dias = ahora_dia - dia;
		  if (ahora_dia < dia) {
			ultimoDiaMes = new Date(ahora_ano, ahora_mes - 1, 0);
			dias = ultimoDiaMes.getDate() - (dia - ahora_dia);
		  }
		  //let tiempo = edad + " años, " + meses + " meses y " + dias + " días";
		  let tiempo = edad + " años ";
		  //edad + " años";
		  //$("#Edad").val(tiempo);
		  document.getElementById("EdadInc").value = tiempo;

		  return; //edad + " años, " + meses + " meses y " + dias + " días";
	}	
	//**Fin Buscar Edad Inc
	
	function buscar_panelistas(){
		//debugger;
		num = document.getElementById("Hogar").value;
		if (num === "" ) {
			return;
		}
		//alert("Llego a Buscar Panelista");
		var stodo = "num=" + num;
		//alert("Todo:=" + stodo);
		$.ajax({
			url:'g_BuscarPanelistas.asp?'+stodo,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivBuscarPanelistas').html(data);
			}
		})
	
	}


	
	
	//**Inicio Buscar Nacionalidad INC
	function buscar_nacionalidadinc(){
		//alert("Llego Nacionalidad INC");
		$.ajax({
			url:'g_BuscarNacionalidadInc.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivNacionalidadInc').html(data);
			}
		})
	}	
	//**Fin Buscar Nacionalidad Inc

	//**Inicio Buscar Nacionalidad Mod
	function buscar_nacionalidadmod(){
		//alert("Llego Nacionalidad INC");
		$.ajax({
			url:'g_BuscarNacionalidadMod.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivNacionalidadMod').html(data);
			}
		})
	}	
	//**Fin Buscar Nacionalidad Mod
	
	//**Inicio Buscar Parentesco Inc
	function buscar_parentescoinc(){
		//alert("Llego Parentesco Inc");
		$.ajax({
			url:'g_BuscarParentescoInc.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivParentescoInc').html(data);
			}
		})
	}	
	//**Fin Buscar Parentesco Inc

	//**Inicio Buscar Parentesco Mod
	function buscar_parentescomod(){
		//alert("Llego Parentesco Mod");
		$.ajax({
			url:'g_BuscarParentescoMod.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivParentescoMod').html(data);
			}
		})
	}	
	//**Fin Buscar Parentesco Mod
	
	//**Inicio Buscar EstadoCivil Inc
	function buscar_estadocivilinc(){
		//alert("Llego EstadoCivil Inc");
		$.ajax({
			url:'g_BuscarEstadoCivilInc.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivEstadoCivilInc').html(data);
			}
		})
	}	
	//**Fin Buscar EstadoCivil Inc

	//**Inicio Buscar EstadoCivil Mod
	function buscar_estadocivilmod(){
		//alert("Llego EstadoCivil Inc");
		$.ajax({
			url:'g_BuscarEstadoCivilMod.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivEstadoCivilMod').html(data);
			}
		})
	}	
	//**Fin Buscar EstadoCivil Mod
	
	//**Inicio Buscar Sexo Inc
	function buscar_sexoinc(){
		//alert("Llego Sexo Inc");
		$.ajax({
			url:'g_BuscarSexoInc.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivSexoInc').html(data);
			}
		})
	}	
	//**Fin Buscar Sexo Inc

	//**Inicio Buscar Sexo Mod
	function buscar_sexomod(){
		//alert("Llego Sexo Inc");
		$.ajax({
			url:'g_BuscarSexoMod.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivSexoMod').html(data);
			}
		})
	}	
	//**Fin Buscar Sexo Mod
	
	//**Inicio Buscar Educacion Inc
	function buscar_educacioninc(){
		//alert("Llego Educacion Inc");
		$.ajax({
			url:'g_BuscarEducacionInc.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivEducacionInc').html(data);
			}
		})
	}	
	//**Fin Buscar Educacion Inc

	//**Inicio Buscar TipoIngreso Inc
	function buscar_tipoingresoinc(){
		//alert("Llego TipoIngreso Inc");
		$.ajax({
			url:'g_BuscarTipoIngresoInc.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivTipoIngresoInc').html(data);
			}
		})
	}	
	//**Fin Buscar TipoIngreso Inc

</script>

<style>
#customers {
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    width: 100%;
}

#customers td, #customers th {
    border: 1px solid #ddd;
    text-align: left;
    padding: 8px;
}

#customers tr:nth-child(even){background-color: #f2f2f2}

#customers tr:hover {background-color: #ddd;}

#customers th {
    padding-top: 12px; 
    padding-bottom: 12px;
    background-color: #4CAF50;
    color: white;
	
}</style>
  

	<div id="DetallePauta">
		<div id="2" class="modalmask">
			<div class="modalbox movedown">
				<a href="#close" title="Cerrar" class="close">x</a>
				<fieldset class="det">
					<legend class="det">Incluir Panelista</legend>
					<!--Bloque 1a-->
					<table id="customers"> 
						<tr>
							<th>Primer Nombre</th>
							<th>Segundo Nombre</th>
							<th>Primer Apellido</th>
							<th>Segundo Apellido</th>
						</tr>
						<tr>
							<td>
								<div id="DivPrimerNombreInc" style="width:35%; float:left;">
									<input type="text" name="PrimerNombreInc" id="PrimerNombreInc" align="right" maxlength=20 size=20>
								</div>
							</td>
							<td>
								<div id="DivSegundoNombreInc" style="width:35%; float:left;">
									<input type="text" name="SegundoNombreInc" id="SegundoNombreInc" align="right" maxlength=20 size=20>
								</div>
							</td>
							<td>
								<div id="DivPrimerApellidoInc" style="width:35%; float:left;">
									<input type="text" name="PrimerApellidoInc" id="PrimerApellidoInc" align="right" maxlength=20 size=20>
								</div>
							</td>
							<td>
								<div id="DivSegundoApellidoInc" style="width:35%; float:left;">
									<input type="text" name="SegundoApellidoInc" id="SegundoApellidoInc" align="right" maxlength=20 size=20>
								</div>
							</td>
						</tr>
					</table>
					<!--Bloque 1b-->
					<table id="customers"> 
						<tr>
							<th>Nacionalidad</th>
							<th>Cedula</th>
							<th>Celular</th>
							<th>EstadoCivil</th>
						</tr>
						<tr>
							<td>
								<div id="DivNacionalidadInc" style="width:35%; float:left;">
									<script>buscar_nacionalidadinc()</script>
								</div>
							</td>
							<td>
								<div id="DivCedulaInc" style="width:35%; float:left;">
									<input type="text" name="CedulaInc" id="CedulaInc" align="right" maxlength=20 size=20>
								</div>
							</td>
							<td>
								<div id="DivCelularInc">
									<input type="text" name="CelularInc" id="CelularInc" align="right" maxlength=20 size=20>
								</div>
							</td>
							<td>
								<div id="DivEstadoCivilInc">
									<script>buscar_estadocivilinc()</script>
								</div>
							</td>
						</tr>
					</table>
					<!--Bloque 1c-->
					<table id="customers"> 
						<tr>
							<th>Correo</th>
							<th>Parentesco</th>
						</tr>
						<tr>
							<td>
								<div id="DivCorreoInc">
									<input type="text" name="CorreoInc" id="CorreoInc" align="right" maxlength=50 size=50>
								</div>
							</td>
							<td>
								<div id="DivParentescoInc">
									<script>buscar_parentescoinc()</script>
								</div>
							</td>
						</tr>
					</table>
					<!--Bloque 1d-->
					<table id="customers"> 
						<tr>
							<th>Fecha Nacimiento</th>
							<th>Edad</th>
							<th>Sexo</th>
						</tr>
						<tr>
							<td>
								<div id="DivFechaNacimientoInc" style="width:35%; float:left;">
									<input type="date" name="FechaNacimientoInc" id="FechaNacimientoInc" align="right" size=20 onblur="buscar_edadinc()">
								</div>
							</td>
							<td>
								<div id="DivEdadInc" style="width:35%; float:left;">
									<input type="text" name="EdadInc" id="EdadInc" disabled align="right" size=10>
								</div>
							</td>
							<td>
								<div id="DivSexoInc">
									<script>buscar_sexoinc()</script>
								</div>
							</td>
						</tr>
					</table>
					<!--Bloque 1e-->
					<table id="customers"> 
						<tr>
							<th>Educacion</th>
							<th>Tipo de Ingreso</th>
						</tr>
						<tr>
							<td>
								<div id="DivEducacionInc" style="width:35%; float:left;">
									<script>buscar_educacioninc()</script>
								</div>
							</td>
							<td>
								<div id="DivTipoIngresoInc">
									<script>buscar_tipoingresoinc()</script>
								</div>
							</td>
						</tr>
					</table>
					<button type="button" onclick="registrarinc()">Registrar</button>
				</fieldset>
			</div> <!--modalbox movedown-->
		</div> <!--modal2-->
		<div id="3" class="modalmask">
			<div class="modalbox movedown">
				<a href="#close" title="Cerrar" class="close">x</a>
				<fieldset class="det">
					<legend class="det">Modificar Panelista</legend>
					<!--Bloque 1a-->
					<table id="customers"> 
						<tr>
							<th>Primer Nombre</th>
							<th>Segundo Nombre</th>
							<th>Primer Apellido</th>
							<th>Segundo Apellido</th>
						</tr>
						<tr>
							<td>
								<div id="DivPrimerNombreMod" style="width:35%; float:left;">
									<input type="text" name="PrimerNombreMod" id="PrimerNombreMod" align="right" maxlength=20 size=20>
								</div>
							</td>
							<td>
								<div id="DivSegundoNombreMod" style="width:35%; float:left;">
									<input type="text" name="SegundoNombreMod" id="SegundoNombreMod" align="right" maxlength=20 size=20>
								</div>
							</td>
							<td>
								<div id="DivPrimerApellidoMod" style="width:35%; float:left;">
									<input type="text" name="PrimerApellidoMod" id="PrimerApellidoMod" align="right" maxlength=20 size=20>
								</div>
							</td>
							<td>
								<div id="DivSegundoApellidoMod" style="width:35%; float:left;">
									<input type="text" name="SegundoApellidoMod" id="SegundoApellidoMod" align="right" maxlength=20 size=20>
								</div>
							</td>
						</tr>
					</table>
					<!--Bloque 1b-->
					<table id="customers"> 
						<tr>
							<th>Nacionalidad</th>
							<th>Cedula</th>
							<th>Celular</th>
							<th>EstadoCivil</th>
						</tr>
						<tr>
							<td>
								<div id="DivNacionalidadMod" style="width:35%; float:left;">
									<script>buscar_nacionalidadmod()</script>
								</div>
							</td>
							<td>
								<div id="DivCedulaInc" style="width:35%; float:left;">
									<input type="text" name="CedulaMod" id="CedulaMod" align="right" maxlength=20 size=20>
								</div>
							</td>
							<td>
								<div id="DivCelularMod">
									<input type="text" name="CelularMod" id="CelularMod" align="right" maxlength=20 size=20>
								</div>
							</td>
							<td>
								<div id="DivEstadoCivilMod">
									<script>buscar_estadocivilmod()</script>
								</div>
							</td>
						</tr>
					</table>
					<!--Bloque 1c-->
					<table id="customers"> 
						<tr>
							<th>Correo</th>
							<th>Parentesco</th>
						</tr>
						<tr>
							<td>
								<div id="DivCorreoMod">
									<input type="text" name="CorreoMod" id="CorreoMod" align="right" maxlength=50 size=50>
								</div>
							</td>
							<td>
								<div id="DivParentescoMod">
									<script>buscar_parentescomod()</script>
								</div>
							</td>
						</tr>
					</table>
					<!--Bloque 1d-->
					<table id="customers"> 
						<tr>
							<th>Fecha Nacimiento</th>
							<th>Edad</th>
							<th>Sexo</th>
						</tr>
						<tr>
							<td>
								<div id="DivFechaNacimientoMod" style="width:35%; float:left;">
									<input type="date" name="FechaNacimientoMod" id="FechaNacimientoMod" align="right" size=20 onblur="buscar_edadmod()">
								</div>
							</td>
							<td>
								<div id="DivEdadMod" style="width:35%; float:left;">
									<input type="text" name="EdadMod" id="EdadMod" disabled align="right" size=10>
								</div>
							</td>
							<td>
								<div id="DivSexoMod">
									<script>buscar_sexomod()</script>
								</div>
							</td>
						</tr>
					</table>
					<button type="button" onclick="registrarmod()">Registrar</button>
				</fieldset>
			</div> <!--modalbox movedown-->
		</div> <!--modal4-->

	</div> <!--DetallePauta1-->

	<button type="button" onclick="buscar_panelistas()">Buscar Integrantes</button>
	<div id="DivBuscarPanelistas" style="width:35%; float:left;">
	</div>
	
	<div style="width:98%">
		<div class="container-fluid">        
			<div class="row">
				<!--Contenido General-->			
				<div class="container">
					<div class="col-md-8 col-sm-8 col-xs-12">
						<div class="pull-right">
							<a href="#2" title="Nuevo Panelista">
								<img src="images/NuevoUsuario.jpg"  style="margin-left:0px;" alt="BuscarLista" width="55px" onclick=""/>
							</a>
							<!--<a href="#3" title="Modificar Panelista">
								<img src="images/NuevoUsuario.jpg"  style="margin-left:0px; " alt="BuscarLista" width="55px" onclick="" />
							</a>
							<a href="#3" title="Modificar Panelista" id="send_click">
								<img src="images/NuevoUsuario.jpg"  style="margin-left:0px; " alt="BuscarLista" width="55px" onclick="" />
							</a>-->
							
						</div>
					</div>
				</div>
			</div>
		</div>
	</div>
	
<script>
	//LR
	//**Inicio Buscar Panelista
	function SelPanelista(id){
		return; 
		//alert("Buscar Panelista:="+id);
		debugger;		
		//BLOQUE 1
		var value= id;
		let ajax = {
			id: id						
		  };
		  $.ajax({
			url: "g_BuscarBloque2.asp",
			type: "POST",
			cache: false,
			data: ajax,			
			dataType:"JSON",			
			success: function (data1, textStatus, jqXHR) {				
				console.log(data1);
				debugger;				
				//					
				$("#Hogar").val(id);
				$("#PrimerNombreMod").val(data1[0].Nombre1Mod);
				$("#SegundoNombreMod").val(data1[0].Nombre2Mod);
				$("#PrimerApellidoMod").val(data1[0].Apellido1Mod);
				$("#SegundoApellidoMod").val(data1[0].Apellido2Mod);
				$("#NacionalidadMod").val(data1[0].Id_NacionalidadMod);
				$("#CedulaMod").val(data1[0].CedulaMod);
				$("#CelularMod").val(data1[0].CelularMod);
				$("#CorreoMod").val(data1[0].CorreoMod);
				$("#ParentescoMod").val(data1[0].Id_ParentescoMod);
				$("#EstadoCivilMod").val(data1[0].Id_EstadoCivilMod);
				$("#FechaNacimientoMod").val(data1[0].Fec_NacimientoMod);
				$("#SexoMod").val(data1[0].Id_SexoMod);
				//buscar_edadmod();
				var rutina = "#send_click";
    var $a = $(rutina);
    var a_url = $a.attr('href');
				
			},
			error: function (request, textStatus, errorThrown) {
			  alert("Error "+ request.responseJSON.message + "error");
			},
		  });
		  //document.getElementById('#3').onclick = true;
		  //alert("LLEGO a Modificar");
		
		  
	}	
	//**Fin Buscar Panelista
	
</script>
