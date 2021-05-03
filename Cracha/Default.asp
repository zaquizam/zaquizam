<!DOCTYPE HTML>
<html >
<head>
	<title>Proc Adm PPIBR</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
	<!--<meta http-equiv="refresh" content="240" />-->
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<meta charset="utf-8">
	<script src="jquery.min.js"></script>
	<!--<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />-->
	<script>
		//**Inicio Cracha
		function buscar(){
			//alert("llego a buscar");
			//debugger;
			var num = "0";
			if(typeof (document.getElementById("sCracha").value!==undefined) || document.getElementById("sCracha").value!==null)
			{
				num = document.getElementById("sCracha").value;
			} 
			var sx = "g_BuscarCracha.asp?num="+num;
			$.ajax({
				url:sx,
				type: "GET",
				beforeSend: function(objeto){
					$('#loader2').html('<img src="./ajax-loader2.gif"> cargando...!');
				},
				success:function(data){
					
					$('#loader2').html('');
					console.log(data);
					$('#DivInformacion').html(data);
				}
			})
		}	
		//**Fin Cracha

		//**Inicio Eliminar
		function eliminar(num) {
			var r = confirm("Desea Eliminar el Cracha?");
			if (r == true) {
				//alert("id:= " + num);
				//return;
				$.ajax({
					url:'g_EliminarCracha.asp?num='+num,
				    type: "GET",
					beforeSend: function(objeto){
						$('#loader2').html('<img src="ajax-loader2.gif"> cargando...!');
					},
					success:function(data){
						//debugger;
						$('#loader2').html('');
						console.log(data);
						$('#DivEliminar').html(data);
						
					}
				})
			} else {
				//txt = "You pressed Cancel!";
			}
			return;
		}
		//**Fin Eliminar


	</script>
</head>
<body topmargin="0">

<h2>Proceso Administrativo</h2>
	
	<div id="DivCracha"> 
		Crachá
		<input type="text" name="sCracha" id="sCracha" align="right" size=15>
		<button type="button" onclick="buscar();">Buscar</button>
	</div>
	<span id="loader2"></span>
	<div id="DivInformacion">
	
	</div>
	<div id="DivEliminar">
	
	</div>
</body>
</html>