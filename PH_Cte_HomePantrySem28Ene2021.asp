<!DOCTYPE HTML>
<html >
<head>
	<title>Home Pantry Semanal</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<script type="text/javascript" src="js/sweetalert.min.js"></script>
	<link href="css/sweetalert.css" rel="stylesheet" type="text/css" media="screen" />	

</head>
<script type="text/javascript">
	function GenerarExcel()
	{
		//alert("Generar Excel");
		num = document.getElementById("Excel").value;
		//alert("Generar Excel:="+ num);
		window.open("g_CteHomePartySemExcel.asp?" + num,"_blank");
	}
</script>
	
<body topmargin="0">
<!--#include file="estiloscss.asp"-->
<!--#include file="encabezado.asp"-->
<!--#include file="nn_subN.asp"-->
<!--#include file="in_DataEN.asp"-->

<%
'==========================================================================================
' Variables y Constantes
'==========================================================================================
    Apertura

	dim idCategoria
	dim idFabricante
	dim idMarca
	dim idArea
	dim strSemana
	dim gProductos
	
	dim gCategoria
	dim gArea
	dim gFabricante
	dim gMarca
	dim gSegmento
	dim gRango
	dim gIndicadores
			
	dim gDatos1
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	dim gDatos2
	dim rsx2
	set rsx2 = CreateObject("ADODB.Recordset")
	rsx2.CursorType = adOpenKeyset 
	rsx2.LockType = 2 'adLockOptimistic 


Sub DataCombos
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	sql = sql & " FROM "
	sql = sql & " PH_DataCruda "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = 1 " 
	sql = sql & " GROUP BY "
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gCategoria = rsx1.GetRows
		rsx1.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Area, "
	sql = sql & " Area "
	sql = sql & " FROM "
	sql = sql & " PH_DataCruda "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = 1 " 
	sql = sql & " GROUP BY "
	sql = sql & " Id_Area, "
	sql = sql & " Area "
	sql = sql & " ORDER BY "
	sql = sql & " Area "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gArea = rsx1.GetRows
		rsx1.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Fabricante, "
	sql = sql & " Fabricante "
	sql = sql & " FROM "
	sql = sql & " PH_DataCruda "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = 1 " 
	sql = sql & " GROUP BY "
	sql = sql & " Id_Fabricante, "
	sql = sql & " Fabricante "
	sql = sql & " HAVING "
	sql = sql & " Id_Fabricante <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Fabricante "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gFabricante = rsx1.GetRows
		rsx1.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca "
	sql = sql & " FROM "
	sql = sql & " PH_DataCruda "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = 1 " 
	sql = sql & " GROUP BY "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca "
	sql = sql & " HAVING "
	sql = sql & " Id_Marca <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Marca "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gMarca = rsx1.GetRows
		rsx1.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Segmento, "
	sql = sql & " Segmento "
	sql = sql & " FROM "
	sql = sql & " PH_DataCruda "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = 1 " 
	sql = sql & " GROUP BY "
	sql = sql & " Id_Segmento, "
	sql = sql & " Segmento "
	sql = sql & " HAVING "
	sql = sql & " Id_Segmento <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Segmento "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gSegmento = rsx1.GetRows
		rsx1.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_RangoTamano, "
	sql = sql & " RangoTamano "
	sql = sql & " FROM "
	sql = sql & " PH_DataCruda "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = 1 " 
	sql = sql & " GROUP BY "
	sql = sql & " Id_RangoTamano, "
	sql = sql & " RangoTamano "
	sql = sql & " HAVING "
	sql = sql & " Id_RangoTamano <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " RangoTamano "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gRango = rsx1.GetRows
		rsx1.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Indicador, "
	sql = sql & " Abreviatura, "
	sql = sql & " Ind_Activo " 
	sql = sql & " FROM "
	sql = sql & " PH_Indicadores "
	sql = sql & " WHERE "
	sql = sql & " Ind_Activo = 1 " 
	sql = sql & " ORDER BY "
	sql = sql & " Id_Indicador "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexion
	if rsx1.eof then
		rsx1.close
	else
		gIndicadores = rsx1.GetRows
		rsx1.close
	end if
	
End Sub
   
    LeePar
      
    if ed_iPas<>4 then 
        Encabezado
    end if    

	DataCombos
%>
	
	<!--hidden-->
	<input type="hidden" name="Filtro" id="Filtro" align="right" size=250>
	<link rel="stylesheet" href="https://netdna.bootstrapcdn.com/bootstrap/3.3.2/css/bootstrap.min.css">
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js"></script>
	<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.2/js/bootstrap.min.js"></script>
	<link rel="stylesheet" href="css/bootstrap-multiselect.css" type="text/css">
	<!--=============================================================================================-->
	<link rel="stylesheet" href="css/homePantry.css" type="text/css">
	<script type="text/javascript" src="js/bootstrap-multiselect.js"></script>				
	<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="css/perfect-scrollbar.css">
	<link rel="stylesheet" type="text/css" href="css/util.css">
	<link rel="stylesheet" type="text/css" href="css/main.css">	
	
	<div class="container-fluid" id="grad1">  
			
			<div class="col-sm-6">
											
				<div class="form-group">
					<!--Categoria-->	
					 <label for="categoria"><i class="fas fa-shapes"></i>&nbsp;Categor??a:</label>
					 <select id="Categoria" multiple="multiple" class="form-control input-sm">
						<% for iCat = 0 to  ubound(gCategoria,2) %>
							<option value="<%=gCategoria(0,iCat)%>" selected><%=ucase(gCategoria(1,iCat))%></option>
						<% next	%>
					</select>					                              
				</div>
				
				<div class="form-group">
					<!--Fabricante-->	
					 <label for="fabricante"><i class="fas fa-industry"></i>&nbsp;Fabricante:</label>
					 <select id="Fabricante" multiple="multiple">
						<option value="0">TOTAL CATEGORIA</option>
						<% for iFra = 0 to  ubound(gFabricante,2) %>
							<option value="<%=gFabricante(0,iFra)%>"><%=gFabricante(1,iFra)%></option>
						<% next %>
					 </select>							                              					
				</div>
				
				<div class="form-group">
					<!--Marca-->
					 <label for="marca"><i class="fas fa-registered"></i>&nbsp;Marca:</label>
					 <select id="Marca" multiple="multiple">
						<!--<option value="0">TOTAL MARCA</option>-->
						<% for iMar = 0 to  ubound(gMarca,2) %>
							<option value="<%=gMarca(0,iMar)%>"><%=gMarca(1,iMar)%></option>
						<% next %>
					</select>					 
				</div>
												
			</div>  <!-- class="col-sm-6"> -->
			
			<div class="col-sm-6">
			
				<div class="form-group">
					<!--Segmento-->
				 	<label for="segmento"><i class="fas fa-project-diagram"></i>&nbsp;Segmento:</label>
				 	<select id="Segmento" multiple="multiple">
						<%	for iSeg = 0 to  ubound(gSegmento,2) %>
							<option value="<%=gSegmento(0,iSeg)%>"><%=gSegmento(1,iSeg)%></option>
						<% next %>
					</select>			 
				</div>
				
				<div class="form-group">
					<!--Indicadores-->
				 	<label for="indicadores"><i class="fas fa-tachometer-alt"></i>&nbsp;Indicadores:</label>
				 	<select id="Indicadores" multiple="multiple">
						<%	for iInd = 0 to  ubound(gIndicadores,2) : sx = gIndicadores(1,iInd) %>
							<option value="<%=gIndicadores(0,iInd)%>"><%=sx%></option>
						<% next %>
					</select> 
				</div>
				
				<div class="form-group">
					
					<div class="col-sm-4">				
						<!--Borrar Filtros-->
						<button type="button" title="Borrar Pantalla"  class="btn btn-block btn-sm btn-danger" onclick="BorrarFiltros()">BORRAR FILTROS&nbsp;&nbsp;<i class="fas fa-times-circle"></i></button>
					</div>
					
					<div class="col-sm-4">				
						<!--Ejecutar-->
						<button type="button" title="Aplicar Selecci??n" class="btn btn-block btn-sm btn-success" id="submit">APLICAR SELECCI??N&nbsp;&nbsp;<i class="fas fa-check"></i></button>
						</div>
					
					<div class="col-sm-4">				
						<!--Exportar-->
						<button type="button" title="Exportar a Excel" class="btn btn-block btn-sm btn-primary" onclick="GenerarExcel();">EXPORTAR EXCEL&nbsp;&nbsp;<i class="fas fa-download"></i></button>
						<!--hidden-->
						<input type="hidden" name="Excel" id="Excel" align="right" size=0 value='<%=sExcel%>'>
					</div>
					
				</div>
							
			</div>  <!-- class="col-sm-6"> -->
	
	</div> <!-- class="container-fluid" id=grad1 --> 
	
	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;" >
		<span ><img src="images/ajax-loader7.gif"><strong>&nbsp;Procesando...., Espere!</strong></span>
	</div>
	<div id="DivHomePartySem">
	</div>
	
	<% conexion.close %>
	
</body>
</html>
<!--================================================================================-->
<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<!--===============================================================================================-->
<script src="js/perfect-scrollbar.min.js"></script>
<script>
	$('.js-pscroll').each(function(){
		var ps = new PerfectScrollbar(this);
		$(window).on('resize', function(){
			ps.update();
		})
	});	
</script>
<script src="js/main.js"></script>
<!--===============================================================================================-->


<script type="text/javascript">
	$(document).ready(function() {
		//
		$('#Categoria').multiselect();
		$('#Fabricante').multiselect();
		$('#Marca').multiselect();
		$('#Segmento').multiselect();
		$('#Rango').multiselect();
		$('#Indicadores').multiselect();
		
		$('#submit').click(function() {
			$("#cargando").css("display", "block");		
			var categoria = $("#Categoria :selected").map((_,e) => e.value).get();
			var fabricante = $("#Fabricante :selected").map((_,e) => e.value).get();
			var marca = $("#Marca :selected").map((_,e) => e.value).get();
			var segmento = $("#Segmento :selected").map((_,e) => e.value).get();
			var indicadores = $("#Indicadores :selected").map((_,e) => e.value).get();
			//debugger;
			
			//alert(categoria);
			//alert("fabricante:" + fabricante);
			//alert("marca:" + marca);
			//alert("segmento:" + segmento);
			//return;
			//alert(indicadores);
			var stodo = "cat=" + categoria;
			stodo = stodo + "&are=";
			stodo = stodo + "&fab=" + fabricante;
			stodo = stodo + "&mar=" + marca;
			stodo = stodo + "&seg=" + segmento;
			stodo = stodo + "&ran=";
			stodo = stodo + "&ind=" + indicadores;
			document.getElementById("Filtro").value = "g_CteHomePartySem.asp?" + stodo;
			document.getElementById("Excel").value = stodo;
			//return;
			$('#DivHomePartySem').html("");
			$.ajax({
				url:'g_CteHomePartySem.asp?'+stodo,
				beforeSend: function(objeto){
					//$('#loader2').html('<img src="./images/ajax-loader2.gif"> cargando...!');						
				},
				success:function(data){
					//debugger;
					//$('#loader2').html('');
					console.log(data);
					$('#DivHomePartySem').html(data);
					$("#cargando").css("display", "none");		
					//alert("Registrado");
					//swal("Datos de Identificacion del Hogar","Registrado","success");
				}
			})

		});
	});
	
	function BorrarFiltros() {
		swal({
                title: "Desea Borrar los Filtros ?",
                text: "",
                type: "warning",
                showCancelButton: true,
                confirmButtonClass: "btn-primary",
                confirmButtonText: "Si",
                cancelButtonText: "No",
                closeOnConfirm: true,
                showLoaderOnConfirm: true
            },
            function() {
                //
                window.open("?x=1&smenu=Reporte%20Semanal","_parent");				
				/*
				$("#Categoria").prop("selectedIndex", 0);
				$("#Fabricante").prop("selectedIndex", 0);
				$("#Marca").prop("selectedIndex", 0);
				$("#Segmento").prop("selectedIndex", 0);
				$("#Indicadores").prop("selectedIndex", 0);
				$('#DivHomePartySem').html("");				
				*/
            });
		return;
	}

</script>
