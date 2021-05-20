<script>

	//**Inicio Obligatorias
	function totalobligatorias(){
		debugger;
		//alert("Paso");
		tot = 0;
		estado = document.getElementById("Estado").value;
		ciudad = document.getElementById("Ciudad").value;
		municipio = document.getElementById("Municipio").value;
		parroquia = document.getElementById("Parroquia").value;
		barrio = document.getElementById("Barrio").value;
		telefono = document.getElementById("TelefonoLocal").value;
		if(estado == "0")
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(ciudad == "0")
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(municipio == "0")
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(parroquia == "0")
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(barrio == "")
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(telefono == "")
		{
		}
		else
		{
			tot = tot + 1;
		}
		nombre1 = document.getElementById("PrimerNombre").value;
		apellido1 = document.getElementById("PrimerApellido").value;
		cedula = document.getElementById("Cedula").value;
		celular1 = document.getElementById("Celular").value;
		numero = document.getElementById("NumeroCortesia").value;
		correo1 = document.getElementById("Correo").value;
		titular = document.getElementById("Titular").value;
		banco = document.getElementById("Banco").value;
		cuenta = document.getElementById("Cuenta").value;
		frecuenciacompra = document.getElementById("FrecuenciaCompra").value;
		numeropersonas = document.getElementById("NumeroPersonas").value;
		if(nombre1 == "")
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(apellido1 == "")
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(cedula == "")
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(celular1 == "")
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(numero == "")
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(correo1 == "")
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(titular == "")
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(banco == 0)
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(cuenta == "")
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(numeropersonas == 0)
		{
		}
		else
		{
			tot = tot + 1;
		}
		if(frecuenciacompra == 0)
		{
		}
		else
		{
			tot = tot + 1;
		}
		document.getElementById("TotalValidas").value = tot;
	}	
	//**Fin Obligatorias

	//**Inicio Registrar
	function registrar1(){
		//alert("Llego Registrar");
		totalobligatorias();
		var mensajeerr="";
		var err=0;
		num = document.getElementById("Hogar").value;
		pais = document.getElementById("Pais").value;
		estado = document.getElementById("Estado").value;
		ciudad = document.getElementById("Ciudad").value;
		municipio = document.getElementById("Municipio").value;
		parroquia = document.getElementById("Parroquia").value;
		calle = document.getElementById("Calle").value;
		edificio = document.getElementById("Edificio").value;
		casa = document.getElementById("Casa").value;
		escalera = document.getElementById("Escalera").value;
		piso = document.getElementById("Piso").value;
		apto = document.getElementById("Apartamento").value;
		barrio = document.getElementById("Barrio").value;
		referencia = document.getElementById("Referencia").value;
		telefono = document.getElementById("TelefonoLocal").value;
		usuario = document.getElementById("idUsuario").value;
		if(estado == "0")
		{
			mensajeerr = mensajeerr + "\nEstado"
			err=err+1;
		}
		if(ciudad == "0")
		{
			mensajeerr = mensajeerr + "\nCiudad"
			err=err+1;
		}
		if(municipio == "0")
		{
			mensajeerr = mensajeerr + "\nMunicipio"
			err=err+1;
		}
		if(parroquia == "0")
		{
			mensajeerr = mensajeerr + "\nParroquia"
			err=err+1;
		}
		//if(calle == "")
		//{
		//	mensajeerr = mensajeerr + "\nCalle / Callejón / Av. / Trs. / Carrera"	
		//	err=err+1;
		//}
		//if(edificio == "")
		//{
		//	mensajeerr = mensajeerr + "\nEdificio"	
		//	err=err+1;
		//}
		//if(casa == "")
		//{
		//	mensajeerr = mensajeerr + "\nCasa Nro. / Qta."	
		//	err=err+1;
		//}
		//if(escalera == "")
		//{
		//	mensajeerr = mensajeerr + "\nEscalera"	
		//	err=err+1;
		//}
		//if(piso == "")
		//{
		//	mensajeerr = mensajeerr + "\nPiso"	
		//	err=err+1;
		//}
		//if(apto == "")
		//{
		//	mensajeerr = mensajeerr + "\nApto."	
		//	err=err+1;
		//}
		if(barrio == "")
		{
			mensajeerr = mensajeerr + "\nBarrio / Urb. / Zona."	
			err=err+1;
		}
		//if(referencia == "")
		//{
		//	mensajeerr = mensajeerr + "\nReferencia"	
		//	err=err+1;
		//}
		if(telefono == "")
		{
			mensajeerr = mensajeerr + "\nTelefono Local"	
			err=err+1;
		}
		if(err > 0)
		{
			//alert("Falta informacion de:" + mensajeerr);
			swal("Falta informacion de:",mensajeerr,"error");
			return
		}
		var stodo = "num=" + num;
		stodo = stodo + "&pai=" + pais;
		stodo = stodo + "&est=" + estado;
		stodo = stodo + "&ciu=" + ciudad;
		stodo = stodo + "&mun=" + municipio;
		stodo = stodo + "&par=" + parroquia;
		stodo = stodo + "&cal=" + calle;
		stodo = stodo + "&edi=" + edificio;
		stodo = stodo + "&cas=" + casa;
		stodo = stodo + "&esc=" + escalera;
		stodo = stodo + "&pis=" + piso;
		stodo = stodo + "&apt=" + apto;
		stodo = stodo + "&bar=" + barrio;
		stodo = stodo + "&ref=" + referencia;
		stodo = stodo + "&tel=" + telefono;
		stodo = stodo + "&usu=" + usuario;
		document.getElementById("Grabar").value = stodo;
		//alert("Todo:=" + stodo);
		$.ajax({
			url:'g_GrabarBloque01.asp?'+stodo,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivHogar').html(data);
				//alert("Registrado");
				swal("Datos de Identificacion del Hogar","Registrado","success");
			}
		})
	}	
	//**Fin Registrar

	//**Inicio Buscar Parroquia
	function buscar_parroquia(){
		//alert("Llego Parroquia1");
		var mreg = document.getElementById("Hogar").value;
		if(typeof (document.getElementById("Municipio").value!==undefined) || document.getElementById("Municipio").value!==null)
		{
				//alert("Llego Parroquia2");
				num = document.getElementById("Municipio").value;
		} else {
				//alert("Llego Parroquia3");
				num = "0";
		}
		
		//alert("Llego Parroquia4");
		$.ajax({
			url:'g_BuscarParroquia.asp?num='+num+'&reg='+mreg,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivParroquia').html(data);
				
			}
		})
	}	
	//**Fin Buscar Parroquia

	//**Inicio Buscar Municipio
	function buscar_municipio(){
		//alert("Llego Municipio1");
		var num = 0;
		var mreg = document.getElementById("Hogar").value;
		//alert("paso");
		if(typeof(document.getElementById("Estado").value!==undefined) || document.getElementById("Estado").value!==null )
		{
				//alert("Llego Municipio2");
				//alert(document.getElementById("Estado").value);
				num = document.getElementById("Estado").value;
				//alert("Llego Municipio2-1");
		} else {
				//alert("Llego Municipio3");
				num = "0";
		}
		
		//alert("Llego Municipio4");
		$.ajax({
			url:'g_BuscarMunicipio.asp?num='+num+'&reg='+mreg,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivMunicipio').html(data);
				buscar_ciudad();
				buscar_parroquia();
				//alert("Llego Municipio5");
			}
		})
	}	
	//**Fin Buscar Municipio

	//**Inicio Buscar Ciudad
	function buscar_ciudad(){
		//alert("Llego Ciudad0");
		var num = "0";
		var mreg = document.getElementById("Hogar").value;
		if(typeof (document.getElementById("Estado").value!==undefined) || document.getElementById("Estado").value!==null)
		{
			num = document.getElementById("Estado").value;
		} 
		var sx = "g_BuscarCiudad.asp?num="+num+"&reg="+mreg;
		$.ajax({
			url:sx,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
			},
			success:function(data){
				
				$('#loader2').html('');
				console.log(data);
				$('#DivCiudad').html(data);
			}
		})
	}	
	//**Fin Buscar Ciudad
	
	//**Inicio Buscar Estado
	function buscar_estado(){
		//alert("Llego estado");
		num = document.getElementById("Pais").value;
		idusu = document.getElementById("idUsuario").value;
		//alert("Llego Usuario=" + idusu);
		$.ajax({
			url:'g_BuscarEstado.asp?num='+num+'&idusu='+idusu,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivEstado').html(data);
				buscar_municipio();
				buscar_ciudad();
				
			}
		})
	}	
	//**Fin Buscar Estado
	
	//**Inicio Buscar Pais
	function buscar_pais(){
		//alert("Llego pais");
		$.ajax({
			url:'g_BuscarPais.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivPais').html(data);
				buscar_estado();
			}
		})
	}	
	//**Fin Buscar Pais


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
<p>Dirección Exacta de la Vivienda</p>
<!--hidden-->
<input type="hidden" name="Grabar" id="Grabar" align="right" size=250>
<input type="hidden" name="idUsuario" id="idUsuario" value="<%=Session("idUsu")%>" align="right" size=4>
<button type="button" onclick="registrar1()">Registrar</button>
<table id="customers">
	<tr>
		<th>Pais</th>
		<th>Estado</th>
		<th>Ciudad</th>
		<th>Municipio</th>
		<th>Parroquia</th>
	</tr>
	<tr>
		<td>
			<div id="DivPais">
				<script>buscar_pais()</script>
			</div>
		</td>
		<td>
			<div id="DivEstado">
			</div>
		</td>
		<td>
			<div id="DivCiudad">

			</div>
		</td>
		<td>
			<div id="DivMunicipio">
			</div>
		</td>
		<td>
			<div id="DivParroquia">
			</div>
		</td>
	</tr>
</table>

<table id="customers">
	<tr>
		<th>Calle / Callejón / Av. / Trs. / Carrera </th>
		<th>Edificio</th>
		<th>Casa Nro. / Qta. </th>
		<th>Escalera</th>
		<th>Piso</th>
		<th>Apto.</th>
	</tr>
	<tr>
		<td>
			<div id="DivCalle" style="width:35%; float:left;">
				<input type="text" name="Calle" id="Calle" align="right" maxlength=30  size=30>
			</div>
		</td>
		<td>
			<div id="DivEdificio" style="width:35%; float:left;">
				<input type="text" name="Edificio" id="Edificio" align="right" maxlength=30 size=30>
			</div>
		</td>
		<td>
			<div id="DivCasa" style="width:35%; float:left;">
				<input type="text" name="Casa" id="Casa" align="right" maxlength=30 size=30>
			</div>
		</td>
		<td>
			<div id="DivEscalera" style="width:10%; float:left;">
				<input type="text" name="Escalera" id="Escalera" align="right" maxlength=10 size=5>
			</div>
		</td>
		<td>
			<div id="DivPiso" style="width:10%; float:left;">
				<input type="text" name="Piso" id="Piso" align="right" maxlength=10 size=5>
			</div>
		</td>
		<td>
			<div id="DivApartamento" style="10%; float:left;">
				<input type="text" name="Apartamento" id="Apartamento" align="right" maxlength=10 size=5>
			</div>
		</td>
	</tr>
</table>

<table id="customers">
	<tr>
		<th>Barrio / Urb. / Zona</th>
		<th>Referencia</th>
		<th>Telefono Local</th>
	</tr>
	<tr>
		<td  >
			<div id="DivBarrio" style="width:50%; float:left;">
				<input type="text" name="Barrio" id="Barrio" align="right" maxlength=30 size=30>
			</div>
		</td>
		<td>
			<div id="DivReferencia" style="width:50%; float:left;">
				<input type="text" name="Referencia" id="Referencia" align="right" maxlength=50 size=50>
			</div>
		</td>
		<td>
			<div id="DivTelefonoLocal">
				<input type="text" name="TelefonoLocal" id="TelefonoLocal" align="right" maxlength=30 size=30>
			</div>
		</td>
	</tr>
</table>

