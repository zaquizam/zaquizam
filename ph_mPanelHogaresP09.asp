<script>
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
<button type="button" onclick="alert('Registrado')">Registrar</button>	
<!--Bloque 1-->
<table id="customers"> 
	<tr>
		<th>Titular</th>
		<th>Banco</th>
		<th>Número Cuenta</th>
		<th>Pago Rápido</th>
	</tr>
		<td>
			<div id="DivTitular">
				<input type="text" name="Titular" id="Titular" align="right" size=30>
			</div>
		</td>
		<td>
			<div id="DivBanco">
				<script>buscar_banco()</script>
			</div>
		</td>
		<td>
			<div id="DivCuenta">
				<input type="text" name="Cuenta" id="Cuenta" align="right" size=20>
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
<!--Bloque 2-->
<table id="customers"> 
	<tr>
		<th>Celular</th>
		<th>Telefono Local</th>
		<th>Telefono  Extra o de Cortesia</th>
	</tr>
		<td>
			<div id="DivCelular">
				<input type="text" name="Celular" id="Celular" align="right" size=20>
			</div>
		</td>
		<td>
			<div id="DivTelefonoLocal">
				<input type="text" name="TelefonoLocal" id="TelefonoLocal" align="right" size=30>
			</div>
		</td>
		<td>
			<div id="DivTelefonoCortesia">
				<input type="text" name="TelefonoCortesia" id="TelefonoCortesia" align="right" size=30>
			</div>
		</td>
		
	</tr>
</table>
