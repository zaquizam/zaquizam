<script>
	//**Inicio Registrar
	function registrar5(){
		//alert("Llego Registrar");
		num = document.getElementById("Hogar").value;
		if (num === "" ) {
			//alert("Debe Registrar Primero los Datos de Identificacion ");
			swal("Televisores","Debe Registrar Primero los Datos de Identificacion","error");
			return;
		}
		//alert("Llego Registrar");
		televisores = document.getElementById("NumeroTeletisores").value;
		tipotelevisores = document.getElementById("TipoTelevisores").value;
		senal = document.getElementById("Senal").value;
		cablera1 = document.getElementById("Cableras1").value;
		cablera2 = document.getElementById("Cableras2").value;
		tvonline1 = document.getElementById("TvOnline1").value;
		tvonline2 = document.getElementById("TvOnline2").value;
		var stodo = "num=" + num;
		stodo = stodo + "&numtv=" + televisores;
		stodo = stodo + "&tiptv=" + tipotelevisores;
		stodo = stodo + "&senal=" + senal;
		stodo = stodo + "&cabl1=" + cablera1;
		stodo = stodo + "&cabl2=" + cablera2;
		stodo = stodo + "&tvon1=" + tvonline1;
		stodo = stodo + "&tvon2=" + tvonline2;
		//alert(televisores);
		//alert(tipotelevisores);
		//alert(senal);
		//alert(cablera1);
		//alert(cablera2);
		//alert(tvonline1);
		//alert(tvonline2);
		//return;
		//alert("Todo:=" + stodo);
		$.ajax({
			url:'g_GrabarBloque07.asp?'+stodo,
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
				swal("Televisores","Registro Actualizado","success");
			}
		})
	}	
	//**Fin Registrar

	//**Inicio Buscar TvOnline1
	function buscar_tvonline1(){
		//alert("Llego TvOnline1");
		num = document.getElementById("Hogar").value;
		//alert(num);
		var stodo = "num=" + num;
		//alert("Todo:=" + stodo);
		//return;
		$.ajax({
			url:'g_BuscarTvOnline1.asp?'+stodo,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivTvOnline1').html(data);
			}
		})
	}	
	//**Fin Buscar TvOnline1

	//**Inicio Buscar TvOnline2
	function buscar_tvonline2(){
		//alert("Llego TvOnline2");
		num = document.getElementById("Hogar").value;
		//alert(num);
		var stodo = "num=" + num;
		//alert("Todo:=" + stodo);
		//return;
		$.ajax({
			url:'g_BuscarTvOnline2.asp?'+stodo,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivTvOnline2').html(data);
			}
		})
	}	
	//**Fin Buscar TvOnline1
	
	//**Inicio Buscar Cablera2
	function buscar_cablera2(){
		//alert("Llego Cablera1");
		num = document.getElementById("Hogar").value;
		ciu = 10; //document.getElementById("Ciudad").value;
		//alert(num);
		var stodo = "num=" + num;
		stodo = stodo + "&ciu=" + ciu;
		//alert("Todo:=" + stodo);
		//return;
		$.ajax({
			url:'g_BuscarCablera2.asp?'+stodo,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				//alert("TV1");
				$('#DivCablera2').html(data);
				//alert("TV2");
			}
		})
	}	
	//**Fin Buscar Cablera12

	//**Inicio Buscar Cablera1
	function buscar_cablera1(){
		//alert("Llego Cablera1");
		num = document.getElementById("Hogar").value;
		ciu = 10; //document.getElementById("Ciudad").value;
		//ciu = document.getElementById("Ciudad").value;
		//alert(num);
		var stodo = "num=" + num;
		stodo = stodo + "&ciu=" + ciu;
		//alert("Todo:=" + stodo);
		//return;
		$.ajax({
			url:'g_BuscarCablera1.asp?'+stodo,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				//alert("TV1");
				$('#DivCablera1').html(data);
				//alert("TV2");
			}
		})
	}	
	//**Fin Buscar Cablera11

	//**Inicio Buscar Señal
	function buscar_senal(){
		//alert("Llego Señal");
		num = document.getElementById("Hogar").value;
		//alert(num);
		$.ajax({
			url:'g_BuscarSenal.asp?num='+num,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				//alert("TV1");
				$('#DivSenal').html(data);
				//alert("TV2");
			}
		})
	}	
	//**Fin Buscar Señal

	//**Inicio Buscar Tipo Televisores
	function buscar_tipotelevisores(){
		//alert("Llego Tipo Televisores");
		num = document.getElementById("Hogar").value;
		//alert(num);
		$.ajax({
			url:'g_BuscarTipoTelevisores.asp?num='+num,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				//alert("TV1");
				$('#DivTipoTelevisores').html(data);
				//alert("TV2");
			}
		})
	}	
	//**Fin Buscar Tipo Televisores

	//**Inicio Buscar Televisores
	function buscar_televisores(){
		num = document.getElementById("Hogar").value;
		//alert(num);
		$.ajax({
			url:'g_BuscarTelevisores.asp?num='+num,
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				//alert("TV1");
				$('#DivNumeroTeletisores').html(data);
				//alert("TV2");
			}
		})
	}	
	//**Fin Buscar Televisores

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
<button type="button" onclick="registrar5()">Registrar</button>
<!--Bloque 1-->
<table id="customers"> 
	<tr>
		<th>Cuantos Televisores hay en su Hogar?</th>
		<th>Tipo de TV</th>
		<th>Señal</th>
		<th>Cableras #1</th>
		<th>Cableras #2</th>
		<th>TV Online #1</th>
		<th>TV Online #2</th>
	</tr>
	<tr>
		<td>
			<div id="DivNumeroTeletisores">
				<script>buscar_televisores()</script>
			</div>
		</td>
		<td>
			<div id="DivTipoTelevisores">
				<script>buscar_tipotelevisores()</script>
			</div>
		</td>
		<td>
			<div id="DivSenal">
				<script>buscar_senal()</script>
			</div>
		</td>
		<td>
			<div id="DivCablera1">
				<script>buscar_cablera1()</script>
			</div>
		</td>
		<td>
			<div id="DivCablera2">
				<script>buscar_cablera2()</script>
			</div>
		</td>
		<td>
			<div id="DivTvOnline1">
				<script>buscar_tvonline1()</script>
			</div>
		</td>
		<td>
			<div id="DivTvOnline2">
				<script>buscar_tvonline2()</script>
			</div>
		</td>
	</tr>
</table>
