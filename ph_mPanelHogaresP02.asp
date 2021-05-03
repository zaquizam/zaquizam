<script>

	//**Inicio Registrar
	function registrar2(){
		//alert("Llego Registrar");
		num = document.getElementById("Hogar").value;
		if (num === "" ) {
			//alert("Debe Registrar Primero los Datos de Identificacion ");
			swal("Caracteristicas de la Vivienda","Debe Registrar Primero los Datos de Identificacion","error");
			return;
		}
		//alert("Llego Registrar");
		var mensajeerr="";
		var err=0;
		tipo = document.getElementById("TipoVivienda").value;
		explique = document.getElementById("Explique").value;
		metros = document.getElementById("MetrosVivienda").value;
		ambientes = document.getElementById("TotalAmbientes").value;
		banos = document.getElementById("TotalBanos").value;
		luz = document.getElementById("PuntosLuz").value;
		var stodo = "num=" + num;
		stodo = stodo + "&tip=" + tipo;
		stodo = stodo + "&exp=" + explique;
		stodo = stodo + "&met=" + metros;
		stodo = stodo + "&amb=" + ambientes;
		stodo = stodo + "&ban=" + banos;
		stodo = stodo + "&luz=" + luz;
		document.getElementById("Grabar3").value = stodo;
		//alert("Todo:=" + stodo);
		$.ajax({
			url:'g_GrabarBloque03.asp?'+stodo,
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
				swal("Caracteristicas de la Vivienda","Registro Actualizado","success");
			}
		})
	}	
	//**Fin Registrar

	//**Inicio Valida Tipo Vivienda
	function validatipovivienda(){
		//alert("Llego a Validar Tipo Vivienda");
		var sTipo = document.getElementById("TipoVivienda").value;
		if (sTipo == "12")
		{
			document.getElementById('Explique').style.display = 'block';
		}
		else
		{
			document.getElementById('Explique').style.display = 'none';
		}
		
	}
	
	//**Fin Valida Tipo Vivienda

	//**Inicio Buscar Tipo Vivienda
	function buscar_tipovivienda(){
		//alert("Llego TipoVivienda");
		$.ajax({
			url:'g_BuscarTipoVivienda.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivTipoVivienda').html(data);
			}
		})
	}	
	//**Fin Buscar TipoVivienda

	//**Inicio Buscar Metros Vivienda
	function buscar_metrosvivienda(){
		//alert("Llego MetrosVivienda");
		$.ajax({
			url:'g_BuscarMetrosVivienda.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivMetrosVivienda').html(data);
			}
		})
	}	
	//**Fin Buscar MetrosVivienda
		
	//**Inicio Buscar Puntos Luz
	function buscar_puntosluz(){
		//alert("Llego PuntosLuz");
		$.ajax({
			url:'g_BuscarPuntosLuz.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivPuntosLuz').html(data);
			}
		})
	}	
	//**Fin Buscar PuntosLuz
	
	
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
<!--hidden-->
<input type="hidden" name="Grabar3" id="Grabar3" align="right" size=250>
<button type="button" onclick="registrar2()">Registrar</button>
<table id="customers">
	<tr>
		<th>Tipo de Vivienda</th>
		<th>(*)Especifique</th>
	</tr>
	<tr>
		<td>
			<div id="DivTipoVivienda">
			<script>buscar_tipovivienda()</script>
			</div>
		</td>
		<td>
			<div id="DivExplique" style="width:33%; float:left;">
				<input type="text" name="Explique" id="Explique" align="right" maxlength=30 size=50>
			</div>
		</td>
	</tr>
</table>
<table id="customers">
	<tr>
		<th>Metros Cuadrados</th>
		<th>Número total de ambientes</th>
		<th>Número total de baños</th>
		<th>Número total de puntos de luz</th>
	</tr>
	<tr>
		<td>
			<div id="DivMetrosVivienda">
				<script>buscar_metrosvivienda()</script>
			</div>
		</td>
		<td>
			<div id="DivTotalAmbientes" style="width:33%; float:left;">
				<input type="number" name="TotalAmbientes" id="TotalAmbientes" align="right" size=5>
			</div>
		</td>
		<td>
			<div id="DivTotalBanos" style="width:33%; float:left;">
				<input type="number" name="TotalBanos" id="TotalBanos" align="right" size=5>
			</div>
		</td>
		<td>
			<div id="DivPuntosLuz" style="width:33%; float:left;">
				<script>buscar_puntosluz()</script>
			</div>
		</td>
	</tr>
</table>

