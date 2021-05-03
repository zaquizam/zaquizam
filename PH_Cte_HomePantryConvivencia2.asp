<!Doctype html>
<!-- PH_Cte_HomePantryConvivencia - 09abr21 -->
<html >
<head>
	<title>Conveniencia</title>
	<meta charset="UTF-8">
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
    <link href="de.css" rel="stylesheet" type="text/css" media="screen" />
    <link href="w3.css" rel="stylesheet" type="text/css" media="screen" />
	<link href="favicon.ico" rel="icon" type="image/x-icon">
	<link href="css/sweetalert.css"  rel="stylesheet" type="text/css" />
	<link href="css/convivencia.css"  rel="stylesheet" type="text/css" />	
	<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
</head>
<body topmargin="0">
		
	<!--#include file="encabezado.asp"-->
	<!--#include file="nn_subN.asp"-->
	<!--#include file="in_DataEN.asp"-->
	
	<% 
		' 09abr21 -
		Apertura		
		LeePar		
		if ed_iPas<>4 then 
			Encabezado
		end if    	
		'	
		' dim gDatosSol
		' dim rsx1
		' set rsx1 = CreateObject("ADODB.Recordset")
		' rsx1.CursorType = adOpenKeyset 
		' rsx1.LockType = 2 'adLockOptimistic 

		' sql = ""
		' sql = sql & " SELECT "
		' sql = sql & " Producto "
		' sql = sql & " FROM "
		' sql = sql & " PH_DataCrudaMensual "
		' sql = sql & " Where CodigoBarra = '7591126522277'"
		' sql = sql & " and Id_Semana in(16,17,18,19) "
		' sql = sql & " and Id_Categoria = 17 "
		' sql = sql & " and Id_Fabricante <> 0 "
		' 'response.write "<br>220 sql:=" & sql
		' 'response.end
		' rsx1.Open sql ,conexion
		' 'response.write "<br> Linea 223 " &
		' 'response.end
		' iExiste = 0
		' if rsx1.eof then
			' iExiste = 0
		' else
			' gDatosSol = rsx1.GetRows
			' rsx1.close
			' iExiste = 1
		' end if
		' for iReg = 0 to ubound(gDatosSol,2)
			' Response.write "<br> " & gDatosSol(0,iReg)
		' next
		'
		' Dim My_String
		' Dim My_Array
		' My_String="Welcome, to, plus2net, learn, web, programming, and, design"   
			
		' My_Array=split(My_String," ")
		' For Each item In My_Array
			' Response.Write("<br>" & item)
		' Next
		' Response.Write "<br>" & Ubound(My_Array) + 1
		'//
		' Dim rsRecordCount, QrySql, rsDetalleConsumoHogares, iDetalleConsumoHogares, resultados(10000,4) 'my array for results
    ' '
    ' ' Capturar las variables
    ' '
	' idMeses 	= Request.QueryString("id_Mes")	
	' '	
	' ' Calcular Total hogares del Mes 
	' '	
	' QrySql = vbnullstring
	' QrySql = QrySql & " SELECT"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	' QrySql = QrySql & " FROM"
	' QrySql = QrySql & " PH_DataCrudaMensual"
	' QrySql = QrySql & " WHERE"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Fabricante  <> 0"
	' QrySql = QrySql & " AND PH_DataCrudaMensual.Id_Semana IN (16, 17, 18, 19)"
	' QrySql = QrySql & " GROUP BY"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar,"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria"
	' QrySql = QrySql & " HAVING"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Categoria IN (1, 3 ,12 , 22)"
	' QrySql = QrySql & " ORDER BY"
	' QrySql = QrySql & " PH_DataCrudaMensual.Id_Hogar;"
	' '
	' Set rsRecordCount = Server.CreateObject("ADODB.recordset")
	' Set rsRecordCount = conexion.Execute(QrySql)
	' if not rsRecordCount.Eof then
		' rsDetalleConsumoHogares = rsRecordCount.GetRows() 
	  	' iDetalleConsumoHogares = UBound(rsDetalleConsumoHogares, 2) + 1 	  
	' else
	  ' iTotalHogares = 0
	' end if
	' '
	' rsRecordCount.Close
	' Set rsRecordCount = nothing
	' Set rsDetalleConsumoHogares = nothing
	' '
	' 'Llenar la matriz Resultante
	' '	
 	' 'ReDim Resultados(iDetalleConsumoHogares, 4)
	' '
 	' FOR  i = 0 to UBound(rsDetalleConsumoHogares,2) 
		' '
   		' hogar = CInt(Resultados(i,0))
		' Compra = CInt(Resultados(i,1))
		' '
		' if Compra = 1 then
			' Resultados(hogar,0)=1
			' Resultados(hogar,1)=0
			' Resultados(hogar,2)=0
			' Resultados(hogar,3)=0
    	' end if
		' '
		' if Compra = 3 then
			' Resultados(hogar,0)=0
			' Resultados(hogar,1)=1
			' Resultados(hogar,2)=0
			' Resultados(hogar,3)=0
    	' end if
		' '
		' if Compra = 12 then
			' Resultados(hogar,0)=0
			' Resultados(hogar,1)=0
			' Resultados(hogar,2)=1
			' Resultados(hogar,3)=0
    	' end if
		' '
		' if Compra = 22 then
			' Resultados(hogar,0)=0
			' Resultados(hogar,1)=0
			' Resultados(hogar,2)=0
			' Resultados(hogar,3)=1
    	' end if
		' '				
	' NEXT	
	' '
	' Response.Write(" <TABLE border=0>")
		' Response.Write("<TR><TD>Row</TD> <TD>Car</TD>")
		' Response.Write("<TD>Year</TD><TD>Price</TD></TR>")

		' 'The UBound function will return the 'index' of the highest element in an array.
		' For i = 0 to UBound(Resultados, 2)
			' Response.Write("<TR><TD>#" & i & "</TD>")
			' Response.Write("<TD>" & Resultados(0,i) & "</TD>")
			' Response.Write("<TD>" & Resultados(1,i) & "</TD>")
			' Response.Write("<TD>" & Resultados(2,i) & "</TD>")
			' Response.Write("<TD>" & Resultados(3,i) & "</TD></TR>")
		' Next

	' Response.Write("</TABLE>")		
	'
		
		'//
	%>

	<div class="container-fluid" id="grad1" >  
	
		<div class="form-group" >
	
			<div class="col-sm-3">
				<div class="form-group"  >				
					<label>Seleccione Mes:</label>
					<select class="form-control input-sm" title="Seleccionar Semana" name="cboProcesarFecha" id="cboProcesarFecha" onchange="procesarFecha();"  />
						<option value="0" select>-- Seleccionar --</option> 					
					</select>
				</div>
			</div>
					
		</div>		
		            								
	</div>        
		
	<div class="container-fluid text-center text-primary" id="cargando" style="display:none;">
		<span ><img src="images/ajax-loader7.gif"><strong>&nbsp;Espere, Procesando...!</strong></span>
	</div>
	
	<div class="container-fluid" id="detallesMaestro" >
		
		<div class="form-group" id="detallesPaso3" style="display:none;">
			
				<table class="table table-striped text-center" style=" margin: auto; width: 50% !important; ">
					<thead>
						<tr>
							<th class="text-center text-danger"><i class='fas fa-check-double'></i>&nbsp;PORCENTAJES DE HOGARES QUE COMPRARON AL MENOS UNA BEBIDA REFRESCANTE ENVASADA&nbsp;</th>			  
						</tr>
					</thead>
					<tbody>
						<tr>
							<th class="text-center text-danger"><h4><span id="Paso3"></span></h4></th>
						</tr>
					</tbody>
				</table>
			
		</div>		
		
		<div class="form-group" id="detallesPaso4" style="display:none;">
											
				<!-- Code by w3codegenerator.com -->
				<div class="table-responsive">
				
					 <table class="table table-borderless table-hover" style=" margin: auto; width: 50% !important;">
						<thead>
							<tr>
								<th colspan="3" class="text-center text-danger"><i class='fas fa-check-double'></i>¿ QUE PORCENTAJES DE HOGARES COMPRARON LAS SIGUIENTES COMBINACIONES: ?</th>	  
							</tr>					
					   </thead>
					   <tbody>
						  <tr>
							 <th class="text-center"scope="row">1.-</th>
							 <td>Solo refresco</td>
							 <td class="text-center">14%</td>						 
						  </tr>
						  <tr>
							 <th scope="row">2</th>
							 <td>Refresco y Agua</td>
							 <td>14%</td>						 
						  </tr>
						  <tr>
							 <th scope="row">3</th>
							 <td>Refresco y Té</td>
							 <td>14%</td>						 
						  </tr>
						  <tr>
							 <th scope="row">4</th>
							 <td>Refresco, Jugos y Agua</td>
							 <td>14%</td>						 
						  </tr>
						  <tr>
							 <th scope="row">5</th>
							 <td>Refresco, Jugos y Té</td>
							 <td>14%</td>						 
						  </tr>
						  <tr>
							 <th scope="row">6</th>
							 <td>Refresco, Agua y Té</td>
							 <td>14%</td>						 
						  </tr>
						  <tr>
							 <th scope="row">7</th>
							 <td>Refresco, Agua, Jugos y Té</td>
							 <td>14%</td>						 
						  </tr>
					   </tbody>
					</table>
				</div>

				
				
				
				
				
				
				
			<!---->
		</div>
						
	</div>			
	<hr>	
		
	<%conexion.close%>
	
</body>
</html>
<script src="https://kit.fontawesome.com/9d7cfbccc5.js" crossorigin="anonymous"></script>
<script src="js/jquery-3.1.1.min.js"></script>
<script src="js/sweetalert.min.js"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script>
<script src="valconveniencia/funcionesV1.js"></script>

<script>
	
	$(document).ready(function() {				
		$(function() {
			buscarFechas();
		});					
	});	
	
</script>

