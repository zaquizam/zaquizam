<script>

	//**Inicio Registrar
	function registrar6(){
		//alert("Llego Registrar");
		num = document.getElementById("Hogar").value;
		if (num === "" ) {
			//alert("Debe Registrar Primero los Datos de Identificacion ");
			swal("Vehiculos","Debe Registrar Primero los Datos de Identificacion","error");
			return;
		}
		//alert("Llego Registrar");
		autos = document.getElementById("Autos").value;
		moto = document.getElementById("Moto").value;
		segurocasco = document.getElementById("SeguroCasco").value;
		var stodo = "num=" + num;
		stodo = stodo + "&aut=" + autos;
		stodo = stodo + "&mot=" + moto;
		stodo = stodo + "&seg=" + segurocasco;
		//alert(autos);
		//alert(segurocasco);
		//alert(moto);
		//return;
		//alert("Todo:=" + stodo);
		//return;
		$.ajax({
			url:'g_GrabarBloque08.asp?'+stodo,
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
				swal("Vehiculos","Registro Actualizado","success");
				
			}
		})
	}	
	//**Fin Registrar

	//**Inicio Buscar Seguro Casco
	function buscar_segurocasco(){
		$.ajax({
			url:'g_BuscarSeguroCasco.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivSeguroCasco').html(data);
			}
		})
	}	
	//**Fin Buscar Seguro Casco

	//**Inicio Buscar Moto
	function buscar_moto(){
		$.ajax({
			url:'g_BuscarMoto.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivMoto').html(data);
			}
		})
	}	
	//**Fin Buscar Moto
	
	//**Inicio Buscar Autos
	function buscar_autos(){
		//alert("Llego Autos");
		num = document.getElementById("Hogar").value;
		//alert(num);
		$.ajax({
			url:'g_BuscarAutos.asp?num='+num,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				//alert("TV1");
				$('#DivAutos').html(data);
				//alert("TV2");
			}
		})
	}	
	//**Fin Buscar Autos


	//LR

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
<button type="button" onclick="registrar6()">Registrar</button>
<!--Bloque 1-->
<table id="customers"> 
	<tr>
		<th>Total de Autos Propios</th>
		<th>Posee Moto</th>
		<th>Cuenta al menos uno de ellos con Seguro de Casco?</th>
	</tr>
	<tr>
		<td>
			<div id="DivAutos">
				<script>buscar_autos()</script>
			</div>
		</td>
		<td>
			<div id="DivMoto">
				<script>buscar_moto()</script>
			</div>
		</td>
		<td>
			<div id="DivSeguroCasco">
				<script>buscar_segurocasco()</script>
			</div>
		</td>
	</tr>
</table>
