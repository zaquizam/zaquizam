<script>

	//**Inicio Registrar
	function registrar3(){
		//alert("Llego Registrar");
		num = document.getElementById("Hogar").value;
		if (num === "" ) {
			//alert("Debe Registrar Primero los Datos de Identificacion ");
			swal("Servicios Publicos","Debe Registrar Primero los Datos de Identificacion","error");
			return;
		}
		//alert("Llego Registrar");
		aguasb = document.getElementById("AguasBlancas").value;
		aguasn = document.getElementById("AguasNegras").value;
		aseo = document.getElementById("AseoUrbano").value;
		electricidad = document.getElementById("Electricidad").value;
		telefonico = document.getElementById("Telefono").value;
		var stodo = "num=" + num;
		stodo = stodo + "&agub=" + aguasb;
		stodo = stodo + "&agun=" + aguasn;
		stodo = stodo + "&ase=" + aseo;
		stodo = stodo + "&ele=" + electricidad;
		stodo = stodo + "&tel=" + telefonico;
		//alert(aguasb);
		//alert(aguasn);
		//alert(aseo);
		//alert(electricidad);
		//alert(telefonico);
		//alert("Todo:=" + stodo);
		$.ajax({
			url:'g_GrabarBloque05.asp?'+stodo,
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
				swal("Servicios Publicos","Registro Actualizado","success");
				
			}
		})
	}	
	//**Fin Registrar

	//**Inicio Buscar Aguas Blancas
	function buscar_aguasblancas(){
		//alert("Llego AguasBlancas");
		$.ajax({
			url:'g_BuscarAguasBlancas.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivAguasBlancas').html(data);
			}
		})
	}	
	//**Fin Buscar Aguas Blancas

	//**Inicio Buscar Aguas Negras
	function buscar_aguasnegras(){
		//alert("Llego AguasNegras");
		$.ajax({
			url:'g_BuscarAguasNegras.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivAguasNegras').html(data);
			}
		})
	}	
	//**Fin Buscar Aguas Negras

	//**Inicio Buscar Aseo Urbano
	function buscar_aseourbano(){
		//alert("Llego Aseo Urbano");
		$.ajax({
			url:'g_BuscarAseoUrbano.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivAseoUrbano').html(data);
			}
		})
	}	
	//**Fin Buscar Aseo Urbano

	//**Inicio Buscar Electricidad
	function buscar_electricidad(){
		$.ajax({
			url:'g_BuscarElectricidad.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivElectricidad').html(data);
			}
		})
	}	
	//**Fin Buscar Electricidad

	//**Inicio Buscar Telefono
	function buscar_telefono(){
		$.ajax({
			url:'g_BuscarTelefono.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivTelefono').html(data);
			}
		})
	}	
	//**Fin Buscar Electricidad
	
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
<button type="button" onclick="registrar3()">Registrar</button>

<table id="customers">
	<tr>
		<th>Aguas blancas</th>
		<th>Aguas negras</th>
		<th>Aseo Urbano</th>
		<th>Servicio de electricidad</th>
		<th>Servicio telefónico</th>
	</tr>
	<tr>
		<td>
			<div id="DivAguasBlancas">
				<script>buscar_aguasblancas()</script>
			</div>
		</td>
		<td>
			<div id="DivAguasNegras">
				<script>buscar_aguasnegras()</script>
			</div>
		</td>
		<td>
			<div id="DivAseoUrbano">
				<script>buscar_aseourbano()</script>
			</div>
		</td>
		<td>
			<div id="DivElectricidad">
				<script>buscar_electricidad()</script>
			</div>
		</td>
		<td>
			<div id="DivTelefono">
				<script>buscar_telefono()</script>
			</div>
		</td>
	</tr>
</table>

