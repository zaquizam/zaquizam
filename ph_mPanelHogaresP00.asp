<script>
	
	
	//**Inicio Registrar
	function registrar0(){
		//alert("Llego Registrar");
		totalobligatorias();
		num = document.getElementById("Hogar").value;
		if (num === "" ) {
			//alert("Debe Registrar Primero los Datos de Identificacion ");
			swal("Responsable del Panel","Debe Registrar Primero los Datos de Identificacion","error");
			return;
		}
		//alert("Llego Registrar");
		var mensajeerr="";
		var err=0;
		nombre1 = document.getElementById("PrimerNombre").value;
		nombre2 = document.getElementById("SegundoNombre").value;
		apellido1 = document.getElementById("PrimerApellido").value;
		apellido2 = document.getElementById("SegundoApellido").value;
		nacionalidad = document.getElementById("Nacionalidad").value;
		cedula = document.getElementById("Cedula").value;
		celular1 = document.getElementById("Celular").value;
		celular2 = document.getElementById("CelularAdicional").value;
		numero = document.getElementById("NumeroCortesia").value;
		correo1 = document.getElementById("Correo").value;
		correo2 = document.getElementById("CorreoAlterno").value;
		titular = document.getElementById("Titular").value;
		cedulatitular = document.getElementById("CedulaTitular").value;
		banco = document.getElementById("Banco").value;
		cuenta = document.getElementById("Cuenta").value;
		pago = document.getElementById("PagoRapido").value;
		parentesco = document.getElementById("Parentesco").value;
		estadocivil = document.getElementById("EstadoCivil").value;
		fechanacimiento = document.getElementById("FechaNacimiento").value;
		sexo = document.getElementById("Sexo").value;
		educacion = document.getElementById("Educacion").value;
		frecuenciacompra = document.getElementById("FrecuenciaCompra").value;
		tipoingreso = document.getElementById("TipoIngreso").value;
		numeropersonas = document.getElementById("NumeroPersonas").value;
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
		if(cedula == "")
		{
			mensajeerr = mensajeerr + "\nCedula"
			err=err+1;
		}
		if(celular1 == "")
		{
			mensajeerr = mensajeerr + "\nCelular"
			err=err+1;
		}
		if(numero == "")
		{
			mensajeerr = mensajeerr + "\nNumero Cortesia"
			err=err+1;
		}
		if(correo1 == "")
		{
			mensajeerr = mensajeerr + "\nCorreo"
			err=err+1;
		}
		if(titular == "")
		{
			mensajeerr = mensajeerr + "\nTitular"
			err=err+1;
		}
		if(banco == 0)
		{
			mensajeerr = mensajeerr + "\nBanco"
			err=err+1;
		}
		if(cuenta == "")
		{
			mensajeerr = mensajeerr + "\nCuenta"
			err=err+1;
		}
		if(numeropersonas == 0)
		{
			mensajeerr = mensajeerr + "\nPersonas del Hogar"
			err=err+1;
		}
		if(frecuenciacompra == 0)
		{
			mensajeerr = mensajeerr + "\nFrecuencia de Compra"
			err=err+1;
		}
		var Max_Length = 20;
		var length = $("#Cuenta").val().length;
		if (length < Max_Length) 
		{
		  //alert("Advertencia: El Numero de Cuenta No Tiene 20 Digitos");
		  swal("Responsable del Panel","Advertencia: El Numero de Cuenta No Tiene 20 Digitos","warning");
		}		
		
		if(err > 0)
		{
			//alert("Falta informacion de:" + mensajeerr);
			swal("Falta informacion de:",mensajeerr,"error");
			return
		}
		var stodo = "num=" + num;
		stodo = stodo + "&nom1=" + nombre1;
		stodo = stodo + "&nom2=" + nombre2;
		stodo = stodo + "&ape1=" + apellido1;
		stodo = stodo + "&ape2=" + apellido2;
		stodo = stodo + "&naci=" + nacionalidad;
		stodo = stodo + "&cedu=" + cedula;
		stodo = stodo + "&cel1=" + celular1;
		stodo = stodo + "&cel2=" + celular2;
		stodo = stodo + "&nume=" + numero;
		stodo = stodo + "&cor1=" + correo1;
		stodo = stodo + "&cor2=" + correo2;
		stodo = stodo + "&titu=" + titular;
		stodo = stodo + "&cedt=" + cedulatitular;
		stodo = stodo + "&banc=" + banco;
		stodo = stodo + "&cuen=" + cuenta;
		stodo = stodo + "&pago=" + pago;
		stodo = stodo + "&pare=" + parentesco;
		stodo = stodo + "&esta=" + estadocivil;
		stodo = stodo + "&fech=" + fechanacimiento;
		stodo = stodo + "&sexo=" + sexo;
		stodo = stodo + "&educ=" + educacion;
		stodo = stodo + "&frec=" + frecuenciacompra;
		stodo = stodo + "&tipo=" + tipoingreso;
		stodo = stodo + "&nump=" + numeropersonas;
		document.getElementById("Grabar0").value = stodo;
		//alert("Todo:=" + stodo);
		$.ajax({
			url:'g_GrabarBloque00.asp?'+stodo,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data); 
				//$('#DivHogar').html(data);
				//alert("Llego Ciudad2");
				//alert("Registro Actualizado"); 
				
				swal("Responsable del Panel","Registro Actualizado","success");
				//swal("Responsable del Panel","Registro Actualizado","warning");
				//swal("Responsable del Panel","Registro Actualizado","error");
				/*
				/swal({
				  title: "/nel",
				  text: "Registro Actualizado",
				  icon: "success",
				  button: "Ok",
				});
				*/
			}
		})
	}	
	//**Fin Registrar


	//**Inicio Buscar Banco
	function buscar_banco(){
		//alert("Llego Banco");
		$.ajax({
			url:'g_BuscarBanco.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivBanco').html(data);
			}
		})
	}	
	//**Fin Buscar Banco

	//**Inicio Frecuencia Compra
	function buscar_frecuenciacompra(){
		//alert("Frecuencia Compra");
		$.ajax({
			url:'g_BuscarFrecuenciaCompra.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivFrecuenciaCompra').html(data);
			}
		})
	}	
	//**Fin Buscar Frecuencia Compra
	
	//**Inicio Buscar Parentesco1
	function buscar_parentesco(){
		//alert("Llego Parentesco1");
		$.ajax({
			url:'g_BuscarParentesco.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivParentesco0').html(data);
			}
		})
	}	
	//**Fin Buscar Parentesco1

	//**Inicio Buscar EstadoCivil
	function buscar_estadocivil(){
		//alert("Llego EstadoCivil");
		$.ajax({
			url:'g_BuscarEstadoCivil.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivEstadoCivil0').html(data);
			}
		})
	}	
	//**Fin Buscar EstadoCivil

	//**Inicio Buscar Sexo
	function buscar_sexo(){
		//alert("Llego Sexo");
		$.ajax({
			url:'g_BuscarSexo.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivSexo0').html(data);
			}
		})
	}	
	//**Fin Buscar Sexo
	
	//**Inicio Buscar Educacion
	function buscar_educacion(){
		//alert("Llego Educacion");
		$.ajax({
			url:'g_BuscarEducacion.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivEducacion').html(data);
			}
		})
	}	
	//**Fin Buscar Educacion

	//**Inicio Buscar CondicionLaboral
	function buscar_condicionlaboral(){
		//alert("Llego CondicionLaboral");
		$.ajax({
			url:'g_BuscarCondicionLaboral.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivCondicionLaboral0').html(data);
			}
		})
	}	
	//**Fin Buscar CondicionLaboral

	//**Inicio Buscar Ocupacion
	function buscar_ocupacion(){
		//alert("Llego Ocupacion");
		$.ajax({
			url:'g_BuscarOcupacion.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivOcupacion0').html(data);
			}
		})
	}	
	//**Fin Buscar Ocupacion
 
	//**Inicio Buscar TipoIngreso
	function buscar_tipoingreso(){
		//alert("Llego TipoIngreso");
		$.ajax({
			url:'g_BuscarTipoIngreso.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivTipoIngreso0').html(data);
			}
		})
	}	
	//**Fin Buscar TipoIngreso
	
	//**Inicio Buscar Nacionalidad
	function buscar_nacionalidad(){
		//alert("Llego Nacionalidad");
		$.ajax({
			url:'g_BuscarNacionalidad.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivNacionalidad0').html(data);
			}
		})
	}	
	//**Fin Buscar TipoIngreso

	function validarFormatoFecha(campo) {
	  let temp = campo.split("-");
	  let fecha = temp[2] + "/" + temp[1] + "/" + temp[0];
	  let RegExPattern = /^\d{1,2}\/\d{1,2}\/\d{2,4}$/;
	  if (fecha.match(RegExPattern)) {
		return true;
	  } else {
		return false;
	  }
	}

	//**Inicio Buscar Edad
	function buscar_edad(){
		//LR
		//alert("Llego Edad:=" + document.getElementById("FechaNacimiento").value);
		//debugger;
		//let fecha = $("#FechaNacimiento").val();
		let fecha = document.getElementById("FechaNacimiento").value
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
			//$("#FechaNacimiento").focus();
			document.getElementById("FechaNacimiento").focus();
			return false;
		  } else {
			if (!validarFormatoFecha(fecha)) {
			  //$("#fechaErr").html(
			  document.getElementById("FechaNacimiento").html(
				'<span style="color:red;">Introduzca fecha valida!</span>'
			  );
			  //$("#FechaNacimiento").focus();
			  document.getElementById("FechaNacimiento").focus();
			  return false;
			} else {
			  //debugger;
			  let hoy = new Date();
			  let fechaFormulario = new Date(fecha);
			  hoy.setHours(0, 0, 0, 0); // Lo iniciamos a 00:00 horas
			  if (hoy <= fechaFormulario) {
				//$("#fechaErr").html(
				document.getElementById("FechaNacimiento").html(
				  '<span style="color:red;">Fecha ingreso posterior a Hoy!</span>'
				);
				return false;
			  } else {
				//$("#fechaErr").html("");
				//document.getElementById("FechaNacimiento").html("");
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
		  document.getElementById("Edad").value = tiempo;

		  return; //edad + " años, " + meses + " meses y " + dias + " días";
	}	
	//**Fin Buscar Edad
	
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
<H3>Responsable del Panel</H3>
<!--hidden-->
<input type="hidden" name="Grabar0" id="Grabar0" align="right" size=250>
<button type="button" onclick="registrar0()">Registrar</button>
<!--Bloque 1a-->
<table id="customers"> 
	<tr>
		<th>Primer Nombre</th>
		<th>Segundo Nombre</th>
		<th>Primer Apellido</th>
		<th>Segundo Apellido</th>
		<th>Nacionalidad</th>
		<th>Cedula</th>
	</tr>
	<tr>
		<td>
			<div id="DivPrimerNombre" style="width:35%; float:left;">
				<input type="text" name="PrimerNombre" id="PrimerNombre" align="right" maxlength=20 size=20>
			</div>
		</td>
		<td>
			<div id="DivSegundoNombre" style="width:35%; float:left;">
				<input type="text" name="SegundoNombre" id="SegundoNombre" align="right" maxlength=20 size=20>
			</div>
		</td>
		<td>
			<div id="DivPrimerApellido" style="width:35%; float:left;">
				<input type="text" name="PrimerApellido" id="PrimerApellido" align="right" maxlength=20 size=20>
			</div>
		</td>
		<td>
			<div id="DivSegundoApellido" style="width:35%; float:left;">
				<input type="text" name="SegundoApellido" id="SegundoApellido" align="right" maxlength=20 size=20>
			</div>
		</td>
		<td>
			<div id="DivNacionalidad0" style="width:35%; float:left;">
				<script>buscar_nacionalidad()</script>
			</div>
		</td>
		<td>
			<div id="DivCedula" style="width:35%; float:left;">
				<input type="text" name="Cedula" id="Cedula" align="right" maxlength=20 size=20>
			</div>
		</td>
	</tr>
</table>
<!--Bloque 1b-->
<table id="customers"> 
	<tr>
		<th>Celular</th>
		<th>Celular Adicional</th>
		<th>Número Cortesia</th>
		<th>Correo</th>
		<th>Correo Alterno</th>
	</tr>
	<tr>
		<td>
			<div id="DivCelular">
				<input type="text" name="Celular" id="Celular" align="right" maxlength=20 size=20>
			</div>
		</td>
		<td>
			<div id="DivCelularAdicional">
				<input type="text" name="CelularAdicional" id="CelularAdicional" align="right" maxlength=20 size=20>
			</div>
		</td>
		<td>
			<div id="DivNumeroCortesia">
				<input type="text" name="NumeroCortesia" id="NumeroCortesia" align="right" maxlength=20 size=20>
			</div>
		</td>
		<td>
			<div id="DivCorreo">
				<input type="text" name="Correo" id="Correo" align="right" maxlength=50 size=20>
			</div>
		</td>
		<td>
			<div id="DivCorreoAlterno">
				<input type="text" name="CorreoAlterno" id="CorreoAlterno" align="right" maxlength=50 size=20>
			</div>
		</td>
	</tr>
</table>
<!--Bloque 1d-->
<table id="customers"> 
	<tr>
		<th>Parentesco</th>
		<th>EstadoCivil</th>
		<th>Fecha Nacimiento</th>
		<th>Edad</th>
		<th>Sexo</th>
	</tr>
	<tr>
		<td>
			<div id="DivParentesco0">
				<script>buscar_parentesco()</script>
			</div>
		</td>
		<td>
			<div id="DivEstadoCivil0">
				<script>buscar_estadocivil()</script>
			</div>
		</td>
		<td>
			<div id="DivFechaNacimiento" style="width:35%; float:left;">
				<input type="date" name="FechaNacimiento" id="FechaNacimiento" align="right" size=20 onblur="buscar_edad()">
			</div>
		</td>
		<td>
			<div id="DivEdad" style="width:35%; float:left;">
				<input type="text" name="Edad" id="Edad" disabled align="right" size=10>
			</div>
		</td>
		<td>
			<div id="DivSexo0">
				<script>buscar_sexo()</script>
			</div>
		</td>
	</tr>
</table>
<!--Bloque 1e-->
<table id="customers"> 
	<tr>
		<!--<th>Condicion Laboral</th>-->
		<th>Educacion</th>
		<!--<th>Ocupacion</th>-->
		<th>Tipo Ingreso</th>
		<th># Personas del Hogar</th>
		<th># Frecuencia de Compra</th>
	</tr>
	<tr>
		<!--<td>
			<div id="DivCondicionLaboral0">
				<script>buscar_condicionlaboral()</script>
			</div>
		</td>-->
		<td>
			<div id="DivEducacion">
				<script>buscar_educacion()</script>
			</div>
		</td>
		<!--<td>
			<div id="DivOcupacion0">
				<script>buscar_ocupacion()</script>
			</div>
		</td>-->
		<td>
			<div id="DivTipoIngreso0">
				<script>buscar_tipoingreso()</script>
			</div>
		</td>
		<td>
			<div id="DivNumeroPersonas">
				<select name="NumeroPersonas" id="NumeroPersonas" onchange="" >
					<option value="0">Seleccione</option> 
					<option value="1">1 Persona</option> 
					<option value="2">2 Personas</option> 
					<option value="3">3 Personas</option> 
					<option value="4">4 Personas</option> 
					<option value="5">5 Personas</option> 
					<option value="6">6 Persona</option> 
					<option value="7">7 Personas</option> 
					<option value="8">8 Personas</option> 
					<option value="9">9 Personas</option> 
					<option value="10">10 Personas</option> 
				</select>
				
			</div>
		</td>
		<td>
			<div id="DivFrecuenciaCompra">
				<script>buscar_frecuenciacompra()</script>
			</div>
		</td>
	</tr>
</table>
<H3>Datos para la transferencia del Incentivo</H3>
<!--Bloque 1c-->
<table id="customers"> 
	<tr>
		<th>Titular</th>
		<th>Cedula Titular</th>
		<th>Banco</th>
		<th>Número Cuenta</th>
		<th>Pago Rápido</th>
	</tr>
		<td>
			<div id="DivTitular">
				<input type="text" name="Titular" id="Titular" align="right" maxlength=30 size=30>
			</div>
		</td>
		<td>
			<div id="DivCedulaTitular">
				<input type="text" name="CedulaTitular" id="CedulaTitular" align="right" maxlength=30 size=30>
			</div>
		</td>
		<td>
			<div id="DivBanco">
				<script>buscar_banco()</script>
			</div>
		</td>
		<td>
			<div id="DivCuenta">
				<input type="text" name="Cuenta" id="Cuenta" align="right" maxlength=20 size=20>
			</div>
		</td>
		<td>
			<div id="DivPagoRapido">
				<select name="PagoRapido" id="PagoRapido" onchange="" >
					<option value="0">Seleccionar</option> 
					<option value="1">Si</option> 
					<option value="2">No</option> 
				</select>
			</div>
		</td>
		
	</tr>
</table>

<!--<button type="button" onclick="alert('Registrado')">Registrar</button>-->

