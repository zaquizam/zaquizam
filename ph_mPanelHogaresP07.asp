<script>
	//**Inicio Registrar
	function registrar7(){
		//alert("Llego Registrar");
		num = document.getElementById("Hogar").value;
		if (num === "" ) {
			//alert("Debe Registrar Primero los Datos de Identificacion ");
			swal("Tenencia de la Vivienda","Debe Registrar Primero los Datos de Identificacion","error");
			return;
		}
		//alert("Llego Registrar");
		ocupacion = document.getElementById("OcupacionVivienda").value;
		explique = document.getElementById("ExpliqueOcupacion").value;
		monto = document.getElementById("MontoVivienda").value;
		var stodo = "num=" + num;
		stodo = stodo + "&ocu=" + ocupacion;
		stodo = stodo + "&exp=" + explique;
		stodo = stodo + "&mon=" + monto;
		//alert(ocupacion);
		//alert(explique);
		//alert(monto);
		//return;
		//alert("Todo:=" + stodo);
		$.ajax({
			url:'g_GrabarBloque04.asp?'+stodo,
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
				swal("Tenencia de la Vivienda","Registro Actualizado","success");
			}
		})
	}	
	//**Fin Registrar


	//**Inicio Valida Ocupacion Vivienda
	function validaocupacion(){
		//alert("Llego a Validar Ocupacion Vivienda");
		var sTipo = document.getElementById("OcupacionVivienda").value;
		//alert(sTipo);
		if (sTipo == "6")
		{
			document.getElementById('ExpliqueOcupacion').style.display = 'block';
		}
		else
		{
			document.getElementById('ExpliqueOcupacion').style.display = 'none';
		}
		
	}
	
	//**Fin Valida Ocupacion Vivienda

	//**Inicio Buscar Ocupacion Vivienda
	function buscar_ocupacionvivienda(){
		//alert("Llego OcupacionVivienda");
		$.ajax({
			url:'g_BuscarOcupacionVivienda.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivOcupacionVivienda').html(data);
			}
		})
	}	
	//**Fin Buscar Ocupacion Vivienda

	//**Inicio Buscar Monto Vivienda
	function buscar_montovivienda(){
		//alert("Llego MontoVivienda");
		$.ajax({
			url:'g_BuscarMontoVivienda.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivMontoVivienda').html(data);
			}
		})
	}	
	//**Fin Buscar Monto Vivienda

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
<button type="button" onclick="registrar7()">Registrar</button>
<!--Bloque 1-->
<table id="customers"> 
	<tr>
		<th>Ocupación Actual de la Vivienda</th>
		<th>(*)Especifique:</th>
		<th>Monto de Alquiler o Hipoteca</th>
	</tr>
		<td>
			<div id="DivOcupacionVivienda">
				<script>buscar_ocupacionvivienda()</script>
			</div>
		</td>
		<td>
			<div id="DivExpliqueOcupacion" style="width:33%; float:left;">
				<input type="text" name="ExpliqueOcupacion" id="ExpliqueOcupacion" align="right" maxlength=30  size=50>
			</div>
		</td>
		<td>
			<div id="DivMontoVivienda">
			<script>buscar_montovivienda()</script>
			</div>
		</td>
	</tr>
</table>
