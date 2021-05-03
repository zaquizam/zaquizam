<script>
	//**Inicio Registrar
	function registrar4(){
		//alert("Llego Registrar");
		num = document.getElementById("Hogar").value;
		if (num === "" ) {
			//alert("Debe Registrar Primero los Datos de Identificacion ");
			swal("Servicios y Equipamiento del Hogar","Debe Registrar Primero los Datos de Identificacion","error");
			return;
		}
		//alert("Llego Registrar");
		sx01 = document.getElementById("DomesticaFija").value;
		sx02 = document.getElementById("PersonalLabores").value;
		sx03 = document.getElementById("DomesticaDia").value;
		sx04 = document.getElementById("Conexion1").value;
		sx05 = document.getElementById("Conexion2").value;
		sx06 = document.getElementById("Conexion3").value;
		sx07 = document.getElementById("TelefonoCelular").value;
		sx08 = document.getElementById("Seguro1").value;
		sx09 = document.getElementById("Seguro2").value;
		sx10 = document.getElementById("Seguro3").value;
		sx11 = document.getElementById("Aire").value;
		sx12 = document.getElementById("CalentadorElectrico").value;
		sx13 = document.getElementById("CalentadorGas").value;
		sx14 = document.getElementById("ComputadorPc").value;
		sx15 = document.getElementById("ComputadorLaptop").value;
		sx16 = document.getElementById("Dvd").value;
		sx17 = document.getElementById("Home").value;
		sx18 = document.getElementById("Juego").value;
		sx19 = document.getElementById("Horno").value;
		sx20 = document.getElementById("Secadora").value;
		sx21 = document.getElementById("LavadoraAutomatica").value;
		sx22 = document.getElementById("LavadoraSemi").value;
		sx23 = document.getElementById("LavadoraRodillo").value;
		sx24 = document.getElementById("Nevera").value;
		sx25 = document.getElementById("Freezer").value;
		sx26 = document.getElementById("CocinaElectrica").value;
		sx27 = document.getElementById("CocinaBombona").value;
		sx28 = document.getElementById("CocinaGas").value;
		sx29 = document.getElementById("CocinaKerosene").value;
		sx30 = document.getElementById("Lavaplatos").value;
		
		var stodo = "num=" + num;
		stodo = stodo + "&sx01=" + sx01;
		stodo = stodo + "&sx02=" + sx02;
		stodo = stodo + "&sx03=" + sx03;
		stodo = stodo + "&sx04=" + sx04;
		stodo = stodo + "&sx05=" + sx05;
		stodo = stodo + "&sx06=" + sx06;
		stodo = stodo + "&sx07=" + sx07;
		stodo = stodo + "&sx08=" + sx08;
		stodo = stodo + "&sx09=" + sx09;
		stodo = stodo + "&sx10=" + sx10;
		stodo = stodo + "&sx11=" + sx11;
		stodo = stodo + "&sx12=" + sx12;
		stodo = stodo + "&sx13=" + sx13;
		stodo = stodo + "&sx14=" + sx14;
		stodo = stodo + "&sx15=" + sx15;
		stodo = stodo + "&sx16=" + sx16;
		stodo = stodo + "&sx17=" + sx17;
		stodo = stodo + "&sx18=" + sx18;
		stodo = stodo + "&sx19=" + sx19;
		stodo = stodo + "&sx20=" + sx20;
		stodo = stodo + "&sx21=" + sx21;
		stodo = stodo + "&sx22=" + sx22;
		stodo = stodo + "&sx23=" + sx23;
		stodo = stodo + "&sx24=" + sx24;
		stodo = stodo + "&sx25=" + sx25;
		stodo = stodo + "&sx26=" + sx26;
		stodo = stodo + "&sx27=" + sx27;
		stodo = stodo + "&sx28=" + sx28;
		stodo = stodo + "&sx29=" + sx29;
		stodo = stodo + "&sx30=" + sx30;
		//alert(sx01);
		//alert(sx02);
		//alert(sx03);
		//alert(sx04);
		//alert(sx05);
		//alert(sx06);
		//alert(sx07);
		//alert(sx08);
		//alert(sx09);
		//alert(sx10);
		//alert(sx11);
		//alert(sx12);
		//alert(sx13);
		//alert(sx14);
		//alert(sx15);
		//alert(sx16);
		//alert(sx17);
		//alert(sx18);
		//alert(sx19);
		//alert(sx20);
		//alert(sx21);
		//alert(sx22);
		//alert(sx23);
		//alert(sx24);
		//alert(sx25);
		//alert(sx26);
		//alert(sx27);
		//alert(sx28);
		//alert(sx29);
		//alert(sx30);
		//alert("Todo:=" + stodo);
		//return;
		$.ajax({
			url:'g_GrabarBloque06.asp?'+stodo,
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
				swal("Servicios y Equipamiento del Hogar","Registro Actualizado","success");
			}
		})
	}	
	//**Fin Registrar

	//**Inicio Buscar Domestica Fija
	function buscar_domesticafija(){
		$.ajax({
			url:'g_BuscarDomesticaFija.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivDomesticaFija').html(data);
			}
		})
	}	
	//**Fin Buscar Domestica Fija

	//**Inicio Buscar Personal Labores
	function buscar_personallabores(){
		$.ajax({
			url:'g_BuscarPersonalLabores.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivPersonalLabores').html(data);
			}
		})
	}	
	//**Fin Buscar Personal Labores

	//**Inicio Buscar Domestica Dia
	function buscar_domesticadia(){
		$.ajax({
			url:'g_BuscarDomesticaDia.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivDomesticaDia').html(data);
			}
		})
	}	
	//**Fin Buscar Domestica Dia

	//**Inicio Buscar Conexion1
	function buscar_conexion1(){
		$.ajax({
			url:'g_BuscarConexion1.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivConexion1').html(data);
			}
		})
	}	
	//**Fin Buscar Conexion1

	//**Inicio Buscar Conexion2
	function buscar_conexion2(){
		$.ajax({
			url:'g_BuscarConexion2.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivConexion2').html(data);
			}
		})
	}	
	//**Fin Buscar Conexion2

	//**Inicio Buscar Conexion3
	function buscar_conexion3(){
		$.ajax({
			url:'g_BuscarConexion3.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivConexion3').html(data);
			}
		})
	}	
	//**Fin Buscar Conexion3

	//**Inicio Buscar Telefono Celular
	function buscar_telefonocelular(){
		$.ajax({
			url:'g_BuscarTelefonoCelular.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivTelefonoCelular').html(data);
			}
		})
	}	
	//**Fin Buscar Telefono Celular

	//**Inicio Buscar Seguro1
	function buscar_seguro1(){
		$.ajax({
			url:'g_BuscarSeguro1.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivSeguro1').html(data);
			}
		})
	}	
	//**Fin Buscar Seguro1

	//**Inicio Buscar Seguro2
	function buscar_seguro2(){
		$.ajax({
			url:'g_BuscarSeguro2.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivSeguro2').html(data);
			}
		})
	}	
	//**Fin Buscar Seguro2

	//**Inicio Buscar Seguro3
	function buscar_seguro3(){
		$.ajax({
			url:'g_BuscarSeguro3.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivSeguro3').html(data);
			}
		})
	}	
	//**Fin Buscar Seguro3
	
	//**Inicio Buscar Aire
	function buscar_aire(){
		$.ajax({
			url:'g_BuscarAire.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivAire').html(data);
			}
		})
	}	
	//**Fin Buscar Aire

	//**Inicio Buscar Calentador Electrico
	function buscar_calentadorelectrico(){
		$.ajax({
			url:'g_BuscarCalentadorElectrico.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivCalentadorElectrico').html(data);
			}
		})
	}	
	//**Fin Buscar Calentador Electrico

	//**Inicio Buscar Calentador Gas
	function buscar_calentadorgas(){
		$.ajax({
			url:'g_BuscarCalentadorGas.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivCalentadorGas').html(data);
			}
		})
	}	
	//**Fin Buscar Calentador Gas
	
	//**Inicio Buscar Computador Pc
	function buscar_computadorpc(){
		$.ajax({
			url:'g_BuscarComputadorPc.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivComputadorPc').html(data);
			}
		})
	}	
	//**Fin Buscar Computador Pc

	//**Inicio Buscar Computador Laptop
	function buscar_computadorlaptop(){
		$.ajax({
			url:'g_BuscarComputadorLaptop.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivComputadorLaptop').html(data);
			}
		})
	}	
	//**Fin Buscar Computador Laptop

	//**Inicio Buscar DVD
	function buscar_dvd(){
		$.ajax({
			url:'g_BuscarDvd.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivDvd').html(data);
			}
		})
	}	
	//**Fin Buscar DVD

	//**Inicio Buscar Home
	function buscar_home(){
		$.ajax({
			url:'g_BuscarHome.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivHome').html(data);
			}
		})
	}	
	//**Fin Buscar Home

	//**Inicio Buscar Juego
	function buscar_juego(){
		$.ajax({
			url:'g_BuscarJuego.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivJuego').html(data);
			}
		})
	}	
	//**Fin Buscar Juego

	//**Inicio Buscar Horno
	function buscar_horno(){
		$.ajax({
			url:'g_BuscarHorno.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivHorno').html(data);
			}
		})
	}	
	//**Fin Buscar Horno

	//**Inicio Buscar Secadora
	function buscar_secadora(){
		$.ajax({
			url:'g_BuscarSecadora.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivSecadora').html(data);
			}
		})
	}	
	//**Fin Buscar Secadora
	
	//**Inicio Buscar Lavadora Automatica
	function buscar_lavadoraautomatica(){
		$.ajax({
			url:'g_BuscarLavadoraAutomatica.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivLavadoraAutomatica').html(data);
			}
		})
	}	
	//**Fin Buscar Lavadora Automatica

	//**Inicio Buscar Lavadora Semi
	function buscar_lavadorasemi(){
		$.ajax({
			url:'g_BuscarLavadoraSemi.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivLavadoraSemi').html(data);
			}
		})
	}	
	//**Fin Buscar Lavadora Semi
	
	//**Inicio Buscar Lavadora Rodillo
	function buscar_lavadorarodillo(){
		$.ajax({
			url:'g_BuscarLavadoraRodillo.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivLavadoraRodillo').html(data);
			}
		})
	}	
	//**Fin Buscar Lavadora Rodillo

	//**Inicio Buscar Nevera
	function buscar_nevera(){
		$.ajax({
			url:'g_BuscarNevera.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivNevera').html(data);
			}
		})
	}	
	//**Fin Buscar Nevera

	//**Inicio Buscar Hreezer
	function buscar_freezer(){
		$.ajax({
			url:'g_BuscarFreezer.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivFreezer').html(data);
			}
		})
	}	
	//**Fin Buscar Freezer

	//**Inicio Buscar Cocina Electrica
	function buscar_cocinaelectrica(){
		$.ajax({
			url:'g_BuscarCocinaElectrica.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivCocinaElectrica').html(data);
			}
		})
	}	
	//**Fin Buscar Cocina Electrica
	
	//**Inicio Buscar Cocina Bombona
	function buscar_cocinabombona(){
		$.ajax({
			url:'g_BuscarCocinaBombona.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivCocinaBombona').html(data);
			}
		})
	}	
	//**Fin Buscar Cocina Bombona

	//**Inicio Buscar Cocina Gas
	function buscar_cocinagas(){
		$.ajax({
			url:'g_BuscarCocinaGas.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivCocinaGas').html(data);
			}
		})
	}	
	//**Fin Buscar Cocina Gas
	
	//**Inicio Buscar Cocina Kerosene
	function buscar_cocinakerosene(){
		$.ajax({
			url:'g_BuscarCocinaKerosene.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivCocinaKerosene').html(data);
			}
		})
	}	
	//**Fin Buscar Cocina Kerosene
	
	//**Inicio Buscar Lavaplatos
	function buscar_lavaplatos(){
		$.ajax({
			url:'g_BuscarLavaplatos.asp',
			beforeSend: function(objeto){
				$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');
				
			},
			success:function(data){
				//debugger;
				$('#loader2').html('');
				console.log(data);
				$('#DivLavaplatos').html(data);
			}
		})
	}	
	//**Fin Buscar Lavaplatos
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
<button type="button" onclick="registrar4()">Registrar</button>
<!--Bloque 1-->
<table id="customers"> 
	<tr>
		<th>Doméstica fija</th>
		<th>Personal para labores específicas </th>
		<th>Doméstica por días</th>
		<th>Conexión a Internet vía telefonía fija por discado / dial-up</th>
		<th>Conexión a Internet vía telefonía fija banda ancha o vía cable</th>
	</tr>
	<tr>
		<td>
			<div id="DivDomesticaFija">
				<script>buscar_domesticafija()</script>
			</div>
		</td>
		<td>
			<div id="DivPersonalLabores">
				<script>buscar_personallabores()</script>
			</div>
		</td>
		<td>
			<div id="DivDomesticaDia">
				<script>buscar_domesticadia()</script>
			</div>
		</td>
		<td>
			<div id="DivConexion1">
				<script>buscar_conexion1()</script>
			</div>
		</td>
		<td>
			<div id="DivConexion2">
				<script>buscar_conexion2()</script>
			</div>
		</td>
	</tr>
</table>
<!--Bloque 2-->
<table id="customers">
	<tr>
		<th>Conexión a Internet vía telefonía móvil / celular</th>
		<th>Teléfono Celular  del Jefe de Familia y/o Pareja</th>
		<th>Seguro HCM particular del Jefe de Familia y/o Pareja y/o hijos</th>
		<th>Seguro HCM colectivo de la empresa para JdeF y/o Pareja ocupados</th>
		<th>Seguro Social Obligatorio para Jefe de Familia o Pareja</th>
	</tr>
	<tr>
		<td>
			<div id="DivConexion3">
				<script>buscar_conexion3()</script>
			</div>
		</td>
		<td>
			<div id="DivTelefonoCelular">
				<script>buscar_telefonocelular()</script>
			</div>
		</td>
		<td>
			<div id="DivSeguro1">
				<script>buscar_seguro1()</script>
			</div>
		</td>
		<td>
			<div id="DivSeguro2">
				<script>buscar_seguro2()</script>
			</div>
		</td>
		<td>
			<div id="DivSeguro3">
				<script>buscar_seguro3()</script>
			</div>
		</td>
	</tr>
</table>
<!--Bloque 3-->
<table id="customers">
	<tr>
		<th>Aire acondicionado</th>
		<th>Calentador de agua eléctrico  NO tipo ducha corona</th>
		<th>Calentador de agua a gas</th>
		<th>Computador  personal (PC)</th>
		<th>Computador Laptop</th>
		<th>DVD y/o Blu-Ray</th>
		<th>Home Theater/Teatro en casa</th>
		<th>Juegos de video</th>
		<th>Horno microondas</th>
		<th>Secadora de ropa</th>
	</tr>
	<tr>
		<td>
			<div id="DivAire">
				<script>buscar_aire()</script>
			</div>
		</td>
		<td>
			<div id="DivCalentadorElectrico">
				<script>buscar_calentadorelectrico()</script>
			</div>
		</td>
		<td>
			<div id="DivCalentadorGas">
				<script>buscar_calentadorgas()</script>
			</div>
		</td>
		<td>
			<div id="DivComputadorPc">
				<script>buscar_computadorpc()</script>
			</div>
		</td>
		<td>
			<div id="DivComputadorLaptop">
				<script>buscar_computadorlaptop()</script>
			</div>
		</td>
		<td>
			<div id="DivDvd"> 
				<script>buscar_dvd()</script>
			</div>
		</td>
		<td>
			<div id="DivHome"> 
				<script>buscar_home()</script>
			</div>
		</td>
		<td>
			<div id="DivJuego"> 
				<script>buscar_juego()</script>
			</div>
		</td>
		<td>
			<div id="DivHorno"> 
				<script>buscar_horno()</script>
			</div>
		</td>
		<td>
			<div id="DivSecadora">
				<script>buscar_secadora()</script>
			</div>
		</td>

	</tr>
</table>

<!--Bloque 4-->
<table id="customers">
	<tr>
		<th>Lavadora de ropa automática</th>
		<th>Lavadora  semiautomática</th>
		<th>Lavadora de ropa de rodillo</th>
		<th>Nevera</th>
		<th>Freezer/Congelador (*)</th>
		<th>Cocina eléctrica</th>
		<th>Cocina a gas de bombona</th>
		<th>Cocina por gas directo</th>
		<th>Cocina a  kerosene / leña,…</th>
		<th>Lavaplatos eléctrico</th>
	</tr>
	<tr>
		<td>
			<div id="DivLavadoraAutomatica">
				<script>buscar_lavadoraautomatica()</script>
			</div>
		</td>
		<td>
			<div id="DivLavadoraSemi">
				<script>buscar_lavadorasemi()</script>
			</div>
		</td>
		<td>
			<div id="DivLavadoraRodillo">
				<script>buscar_lavadorarodillo()</script>
			</div>
		</td>
		<td>
			<div id="DivNevera">
				<script>buscar_nevera()</script>
			</div>
		</td>
		<td>
			<div id="DivFreezer">
				<script>buscar_freezer()</script>
			</div>
		</td>
		<td>
			<div id="DivCocinaElectrica">
				<script>buscar_cocinaelectrica()</script>
			</div>
		</td>
		<td>
			<div id="DivCocinaBombona">
				<script>buscar_cocinabombona()</script>
			</div>
		</td>
		<td>
			<div id="DivCocinaGas">
				<script>buscar_cocinagas()</script>
			</div>
		</td>
		<td>
			<div id="DivCocinaKerosene">
				<script>buscar_cocinakerosene()</script>
			</div>
		</td>
		<td>
			<div id="DivLavaplatos">
				<script>buscar_lavaplatos()</script>
			</div>
		</td>

	</tr>
</table>
