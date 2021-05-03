<script>

	//**Inicio Registrar
	function registrar10(){
		//alert("Llego Registrar");
		num = document.getElementById("Hogar").value;
		if (num === "" ) {
			//alert("Debe Registrar Primero los Datos de Identificacion ");
			swal("Otros","Debe Registrar Primero los Datos de Identificacion","error");
			return;
		}
		//alert("Llego Registrar");
		mascotas = document.getElementById("Mascotas").value;
		if (document.getElementById("Perro").checked)
		{
			perro = 1;
		}
		else
		{
			perro = 0;
		}
		if (document.getElementById("Gato").checked)
		{
			gato = 1;
		}
		else
		{
			gato = 0;
		}
		if (document.getElementById("Pez").checked)
		{
			pez = 1;
		}
		else
		{
			pez = 0;
		}
		if (document.getElementById("Ave").checked)
		{
			ave = 1;
		}
		else
		{
			ave = 0;
		}
		if (document.getElementById("Roedor").checked)
		{
			roedor = 1;
		}
		else
		{
			roedor = 0;
		}
		if (document.getElementById("Otro").checked)
		{
			otro = 1;
		}
		else
		{
			otro = 0;
		}
		var stodo = "num=" + num;
		stodo = stodo + "&mas=" + mascotas;
		stodo = stodo + "&per=" + perro;
		stodo = stodo + "&gat=" + gato;
		stodo = stodo + "&pez=" + pez;
		stodo = stodo + "&ave=" + ave;
		stodo = stodo + "&roe=" + roedor;
		stodo = stodo + "&otr=" + otro;
		//alert("mascotas:=" + mascotas);
		//alert("perro:=" + perro);
		//alert("gato:=" + gato);
		//alert("pez:=" + pez);
		//alert("ave:=" + ave);
		//alert("roedor:=" + roedor);
		//alert("otro:=" + otro);
		//alert("Todo:=" + stodo);
		//return;
		
		$.ajax({
			url:'g_GrabarBloque09.asp?'+stodo,
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
				swal("Otros","Registro Actualizado","success");
			}
		})
	}	
	//**Fin Registrar

	//**Inicio Buscar Mascotas
	function buscar_mascotas(){
		//alert("Llego Mascotas");
		num = document.getElementById("Hogar").value;
		//alert(num);
		$.ajax({
			url:'g_BuscarMascotas.asp?num='+num,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				//alert("TV1");
				$('#DivMascotas').html(data);
				//alert("TV2");
			}
		})
	}	
	//**Fin Buscar Mascotas

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
}
 input.larger {
        width: 20px;
        height: 20px;
      }

</style>	
<button type="button" onclick="registrar10()">Registrar</button>
<!--Bloque 1-->
<table id="customers"> 
	<tr>
		<th>Tiene Mascota</th>
		<th>Tipo Mascota</th>
	</tr>
		<td>
			<div id="DivMascotas">
				<script>buscar_mascotas()</script>
			</div>
		</td>
		<td>
			<div id="DivTipoMascota" style="font-size:150%;" >
				Perro <input type="checkbox" id="Perro" name="Perro" value="" class="larger">
				Gato <input type="checkbox" id="Gato" name="Gato" value="" class="larger">
				Pez <input type="checkbox" id="Pez" name="Pez" value="" class="larger">
				Ave <input type="checkbox" id="Ave" name="Ave" value="" class="larger">
				Roedor <input type="checkbox" id="Roedor" name="Roedor" value="" class="larger">
				Otro <input type="checkbox" id="Otro" name="Otro" value="" class="larger">
			</div>
		</td>
		
	</tr>
</table>
