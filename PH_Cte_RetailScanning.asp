<!DOCTYPE HTML>
<html >
<head>
	<title>Retail Scanning</title>
    <meta name="Robots" content="noindex" >
    <meta name="Robots" content="none" >
    <meta name="Robots" content="nofollow" >
	<link rel="icon" href="favicon.ico" type="image/x-icon"> 
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<script type="text/javascript" src="js/sweetalert.min.js"></script>
	<link href="css/sweetalert.css" rel="stylesheet" type="text/css" media="screen" />	

</head>
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
	dim conexionRS
	Set conexionRS = Server.CreateObject("ADODB.Connection")
	
	conexionRS.ConnectionString= "Provider=SQLOLEDB.1;Password='PHaa11..**';Persist Security Info=True;User ID=cacevedo_atenas;Initial Catalog=RetailScannig;Data Source=216.198.73.34"
	conexionRS.mode = 3
	conexionRS.Open
	
	Dim idCliente
	dim idCategoria
	dim idFabricante
	dim idMarca
	dim idArea
	dim idZona
	dim idCanal
	dim idTamano
	dim idProducto
	dim idIndicador
	
	dim strSemana
	dim gCategoria
	dim gArea
	dim gZona
	dim gCanal
	dim gFabricante
	dim gMarca
	dim gSegmento
	dim gTamano
	dim gRango
	dim gIndicadores
	dim gSemanas
	dim gProducto

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

	idCliente = Session("idCliente")

    
%>
	<script>
		function Mensaje(){
			swal("Atenas Grupo Consultor","Servicio No Contratado","info");
			return;
		}	

	</script>   
<%

Sub ed_vCombo2


    dim rst    
    
	set rst = server.CreateObject("ADODB.Recordset")
	rst.CursorType = 1
	rst.LockType = 1
	'response.write "<br>2423 ed_sCombo(i,4):= " & ed_sCombo(i,4)

%>

<%    	
	for i=1 to ed_iCombo
		
		'response.write "<br>" & ed_iCombo
		
		rst.open ed_sCombo(i,1),conexionRS

	    dim gX
        gX=rst.getrows
        'response.write "<br>2357 i:= "  & i
		ed_sPar(i,1)=gX(0,0)
        if isnull(ed_sPar(i,0)) or ed_sPar(i,0)="" then 
            if ed_sCombo(i,2)<>"" then 
                ed_sPar(i,0)=ed_sCombo(i,2)
             else   
                ed_sPar(i,0)=gX(0,0)
             end if   
           ' response.write "<br>klasjjjjjjjjjjjjjjjjjjj" & ed_sPar(i,0)
        end if    
        rst.close   

        'if ubound(gX,2)<>0 then
       %>
        
   
	        <%if ed_iPas<> 4 then %>
	    	    <select size="1"  name="per" id="Select1"  onchange ="location.href=this.options[this.selectedIndex].value" >
            <%  if ed_sCombo(i,2)<>"" then   
    		       sP=ed_sPar(i,0)
    		       ed_sPar(i,0)=ed_sCombo(i,2)
    	    	   ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_iFil,ed_iMp,ed_iMs
    	    	   ed_sPar(i,0)=sP

            %>
        			<option value="<%=sPar%>"  <% if ed_sPar(i,0) =ed_sCombo(i,2) then response.Write"selected" %> style="width:100%; font-size:20px; font-family:Tahoma; padding:10px 0 10px 0 " >
				    <%="[" & ed_sCombo(i,2) & "]"%>
			        </option>

    		<%  end if
    		    wO=0
               
    		    for j=0 to ubound(gX,2)
    		   
    		       sP=ed_sPar(i,0)
    		       ed_sPar(i,0)=gx(0,j)
    	    	   ed_CalPar 1,ed_iCla,1,"",ed_iCol,ed_iOrd, ed_iFil,ed_iMp,ed_iMs
    	    	   ed_sPar(i,0)=sP
    	    	   
				   if isnumeric(ed_sPar(i,0)) then
						ix=ed_sPar(i,0) -gX(0,j)
				   else
					    ix=1
					    if ed_sPar(i,0) =gX(0,j) then ix=0
				   end if
     	    %>
			    <option value="<%=sPar%>"  
			    <% 
			    if ix=0 then 
			        response.Write"selected" 
			        wO=1
			    end if    
			        %>  >
				    <%=gX(1,j) %>
			    </option>
			    <%next%>
			   
	            </select>
	           
	              <%
	              
	              if wO=0 then 
			        if ed_sCombo(i,2)<>"" then 
                        ed_sPar(i,0)=ed_sCombo(i,2)
                    else   
			            ed_sPar(i,0)=ed_sPar(i,1) 
			           '  Response.write "<br />2509 wO:=" & WO & " " & ed_Spar(1,0)
			        end if    
			     end if%>
	       <% else 
	            %>
	                <%=ed_sPar(i,0) %>
	            <%
	          end if %>
	    <%
		'end if 
		%>         
<%	Next%>

<%    
            		       
End sub

sub VerificarData
	
	if idCliente <> 1 then
		%>
		<script language="JavaScript" type="text/javascript">
			Mensaje()
		</script>
		<%
		response.end
	end if
	'response.write "<br>36 Cliente:= " & Session("idCliente")
end sub

Sub Combos
 
	'response.write "<br>372 Combo1:=" & ed_sPar(1,0)
	'response.write " Combo2:=" & ed_sPar(2,0)
	'response.write " Combo3:=" & ed_sPar(3,0)
	'response.write " Combo3:=" & ed_sPar(4,0)
	'response.write " Combo3:=" & ed_sPar(5,0)
    ed_iCombo = 1
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	sql = sql & " FROM RS_DataProcSem "
	sql = sql & " GROUP BY "
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	sql = sql & " ORDER BY "
	sql = sql & " Categoria "
	'response.write "<br>372 Combo1:=" & sql
    ed_sCombo(1,0)="Categoria"
    ed_sCombo(1,1)=sql 
    'ed_sCombo(1,2)="Sin Selección"
	'response.write "<br>llego"
	'response.end
	
End Sub

Sub DataCombos
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	sql = sql & " FROM RS_DataProcSem "
	sql = sql & " GROUP BY "
	sql = sql & " Id_Categoria, "
	sql = sql & " Categoria "
	sql = sql & " ORDER BY "
	sql = sql & " Categoria "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexionRS
	if rsx1.eof then
		rsx1.close
	else
		gCategoria = rsx1.GetRows
		rsx1.close
	end if
	if ed_sPar(1,0) = "" then ed_sPar(1,0) = gCategoria(0,0)

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Area, "
	sql = sql & " Area "
	sql = sql & " FROM "
	sql = sql & " RS_DataProcSem "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_Area, "
	sql = sql & " Area "
	sql = sql & " ORDER BY "
	sql = sql & " Area "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexionRS
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
	sql = sql & " RS_DataProcSem "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_Fabricante, "
	sql = sql & " Fabricante "
	sql = sql & " HAVING "
	sql = sql & " Id_Fabricante <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Fabricante "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexionRS
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
	sql = sql & " RS_DataProcSem "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_Marca, "
	sql = sql & " Marca "
	sql = sql & " HAVING "
	sql = sql & " Id_Marca <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Marca "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexionRS
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
	sql = sql & " RS_DataProcSem "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_Segmento, "
	sql = sql & " Segmento "
	sql = sql & " HAVING "
	sql = sql & " Id_Segmento <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Segmento "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexionRS
	if rsx1.eof then
		rsx1.close
	else
		gSegmento = rsx1.GetRows
		rsx1.close
	end if

	iSemanaDes = 24
	iSemanaHas = 36
	'response.write "<br>310 Semana Desde:= " &  iSemanaDes
	'response.write "<br>310 Semana Hasta:= " &  iSemanaHas
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " IdSemana, "
	sql = sql & " Semana "
	sql = sql & " FROM "
	sql = sql & " ss_Semana "
	sql = sql & " WHERE "
	sql = sql & " IdSemana >= " & iSemanaDes
	sql = sql & " And IdSemana <= " & iSemanaHas
	sql = sql & " ORDER BY "
	sql = sql & " IdSemana DESC "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexionRS
	if rsx1.eof then
		rsx1.close
	else
		gSemanas = rsx1.GetRows
		rsx1.close
		iSemanaDes = iSemanaDes
		iSemanaHas = iSemanaHas
	end if
	'response.write "<br>246 DataCombos"
	'response.end

	sql = "" 
	sql = sql & " SELECT "
	sql = sql & " Id_Area, "
	sql = sql & " rtrim(Area) "
	sql = sql & " FROM "
	sql = sql & " RS_DataProcSem "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_Area, "
	sql = sql & " Area "
	sql = sql & " ORDER BY "
	sql = sql & " Area "
	'response.write "<br>372 Combo1:=" & sql
	'response.end
	rsx1.Open sql ,conexionRS
	if rsx1.eof then
		rsx1.close
	else
		gArea = rsx1.GetRows
		rsx1.close
	end if

	sql = "" 
	sql = sql & " SELECT "
	sql = sql & " Id_Canal, "
	sql = sql & " rtrim(Canal) "
	sql = sql & " FROM "
	sql = sql & " RS_DataProcSem "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria = " & ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_Canal, "
	sql = sql & " Canal "
	sql = sql & " ORDER BY "
	sql = sql & " Canal "
	'response.write "<br>372 Combo1:=" & sql
	'response.end
	rsx1.Open sql ,conexionRS
	if rsx1.eof then
		rsx1.close
	else
		gCanal = rsx1.GetRows
		rsx1.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Tamano, "
	sql = sql & " rtrim(Tamano) "
	sql = sql & " FROM "
	sql = sql & " RS_DataProcSem "
	sql = sql & " WHERE "
	sql = sql & " Id_Categoria =  " & ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_Tamano, "
	sql = sql & " Tamano "
	sql = sql & " HAVING "
	sql = sql & " Id_Tamano <> 0 "
	sql = sql & " ORDER BY "
	sql = sql & " Tamano "
	'response.write "<br>372 sql4:=" & sql
	rsx1.Open sql ,conexionRS
	if rsx1.eof then
		rsx1.close
	else
		gTamano = rsx1.GetRows
		rsx1.close
	end if
	
	sql = ""
	sql = sql & " SELECT "
	sql = sql & " CodigoBarra, "
	sql = sql & " Descripcion "
	sql = sql & " FROM "
	sql = sql & " RS_DataProcSem "
	sql = sql & " WHERE  "
	sql = sql & " Id_Categoria= " &  ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " CodigoBarra, "
	sql = sql & " Descripcion "
	sql = sql & " HAVING "
	sql = sql & " (CodigoBarra Is Not Null "
	sql = sql & " And CodigoBarra<>'') "
	sql = sql & " AND (Descripcion Is Not Null "
	sql = sql & " And Descripcion<>'' ) "
	sql = sql & " ORDER BY "
	sql = sql & " Descripcion "
	'response.write "<br>372 sql4:=" & sql
	'response.end
	rsx1.Open sql ,conexionRS
	if rsx1.eof then
		rsx1.close
	else
		gProducto = rsx1.GetRows
		rsx1.close
	end if


	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Indicador, "
	sql = sql & " Abreviatura "
	sql = sql & " FROM "
	sql = sql & " RS_Indicadores "
	sql = sql & " WHERE "
	sql = sql & " Ind_Sem = 1 " 
	sql = sql & " AND Ind_Activo = 1 " 
	sql = sql & " ORDER BY "
	sql = sql & " Id_Indicador "
	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexionRS
	if rsx1.eof then
		rsx1.close
	else
		gIndicadores = rsx1.GetRows
		rsx1.close
	end if

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_Zona, "
	sql = sql & " Zona "
	sql = sql & " FROM "
	sql = sql & " RS_DataProcSem "
	sql = sql & " WHERE  "
	sql = sql & " Id_Categoria= " &  ed_sPar(1,0)
	sql = sql & " GROUP BY "
	sql = sql & " Id_Zona, "
	sql = sql & " Zona "
	sql = sql & " ORDER BY "
	sql = sql & " Zona "

	'response.write "<br>372 Combo1:=" & sql
	rsx1.Open sql ,conexionRS
	if rsx1.eof then
		rsx1.close
	else
		gZona = rsx1.GetRows
		rsx1.close
	end if

	'response.write "<br>372 LLEGO"
	'response.end
End Sub

   
    LeePar
  
    
    if ed_iPas<>4 then 
        Encabezado
    end if    
	
	VerificarData
	Combos
	DataCombos

	'response.write "llego1"
	'response.end
    

%>
	<!--hidden-->
	<input type="hidden" name="Filtro" id="Filtro" align="right" size=250>
	<input type="hidden" name="Cliente" id="Cliente" align="right" size=4 value="<%=Session("idCliente")%>">
	<input type="hidden" name="Cat" id="Cat" align="right" size=4 value="<%=ed_sPar(1,0)%>">
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
			
			<div class="col-sm-4">
											
				<div class="form-group">
					<!--Categoria-->	
					 <label for="categoria"><i class="fas fa-shapes"></i>&nbsp;Categoría:</label>
					<%
					ed_vCombo2
					%>
				</div>
				
				<div class="form-group">
					<!--area-->	
					 <label for="area"><i class="	fas fa-globe-americas"></i>&nbsp;Área:</label>
					 <select id="Area" multiple="multiple">
						<% for iAre = 0 to  ubound(gArea,2) %>
							<option value="<%=gArea(0,iAre)%>"><%=gArea(1,iAre)%></option>
						<% next %>
					 </select>							                              					
				</div>

				<div class="form-group">
					<!--Zona-->	
					 <label for="zona"><i class="	fas fa-globe-europe"></i>&nbsp;Zona:</label>
					 <select id="Zona" multiple="multiple">
						<% for iZon = 0 to  ubound(gZona,2) %>
							<option value="<%=gZona(0,iZon)%>"><%=gZona(1,iZon)%></option>
						<% next %>
					 </select>							                              					
				</div>

				<div class="form-group">
					<!--canal-->	
					 <label for="canal"><i class="	fas fa-home"></i>&nbsp;Canal:</label>
					 <select id="Canal" multiple="multiple">
						<% for iCan = 0 to  ubound(gCanal,2) %>
							<option value="<%=gCanal(0,iCan)%>"><%=gCanal(1,iCan)%></option>
						<% next %>
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
					 <!--<img src="images/Select01.ico"  style="margin-left:0px;" title="Seleccionar Todos" alt="All" width="40px" onclick="TodoFab('Fabricante', true)"/>-->
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
					<!--Tamaño-->
				 	<label for="tamano"><i class="fas fa-ruler-combined"></i>&nbsp;Tamaño:</label>
				 	<select id="Tamano" multiple="multiple">
						<%	
						
						for iTam = 0 to  ubound(gTamano,2) %>
							<option value="<%=gTamano(0,iTam)%>"><%=gTamano(1,iTam)%></option>
						<% 
						next 
						
						%>
					</select>			 
				</div>


				<div class="form-group">
					<!--Producto-->
				 	<label for="producto"><i class="fas fa-barcode"></i>&nbsp;Producto:</label>
				 	<select id="Producto" multiple="multiple">
						<%	
						
						for iPro = 0 to  ubound(gProducto,2) 
							DescPro = gProducto(1,iPro) & " - " & gProducto(0,iPro)
							%>
							<option value="<%=gProducto(0,iPro)%>"><%=DescPro%></option>
							<% 
						next 
						
						%>
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
					<!--Semanas-->
				 	<label for="semanas"><i class="fas fa-calendar"></i>&nbsp;Semanas:</label>
				 	<select id="Semanas" multiple="multiple">
						<%	
							sMarcar = "selected"
							for iSem = 0 to  ubound(gSemanas,2) 
								sx = gSemanas(1,iSem) 
								if iSem > 4 then sMarcar = "" 
								%>
								<option value="<%=gSemanas(0,iSem)%>" <%=sMarcar%> ><%=sx%></option>
								<% 
							next 
						%>
					</select> 
				</div>
				
				<div class="form-group">
					
					<div class="col-sm-4">				
						<!--Borrar Filtros-->
						<button type="button" title="Borrar Pantalla"  class="btn btn-block btn-sm btn-danger" onclick="BorrarFiltros()">BORRAR FILTROS&nbsp;&nbsp;<i class="fas fa-times-circle"></i></button>
					</div>
					
					<div class="col-sm-4">				
						<!--Ejecutar-->
						<button type="button" title="Aplicar Selección" class="btn btn-block btn-sm btn-success" id="submit">APLICAR SELECCIÓN&nbsp;&nbsp;<i class="fas fa-check"></i></button>
						</div>
					
					<div class="col-sm-4">				
						<!--Exportar-->
						<button type="button" title="Exportar a Excel" class="btn btn-block btn-sm btn-primary" onclick="GenerarExcel();">EXPORTAR EXCEL&nbsp;&nbsp;<i class="fas fa-download"></i></button>
						<!--hidden-->
						<input type="hidden" name="Excel" id="Excel" align="right" size=0 value='<%=sExcel%>'>
					</div>
					
				</div>
							
			</div>  <!-- class="col-sm-6"> -->
			<div class="col-sm-2">
				<img alt="Logo de la Empresa" src="images/logo/LogoReatilScanning.png" style = "width:128px;  " class="img-responsive center-block" >
			</div>
	
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
		$('#Area').multiselect();
		$('#Zona').multiselect();
		$('#Canal').multiselect();
		$('#Fabricante').multiselect();
		$('#Marca').multiselect();
		$('#Segmento').multiselect();
		$('#Tamano').multiselect();
		$('#Producto').multiselect();
		$('#Rango').multiselect();
		$('#Indicadores').multiselect();
		$('#Semanas').multiselect();
		$('#SemanasAcum').multiselect();
		
		$('#submit').click(function() {
			debugger;
			
			//var categoria = $("#Categoria :selected").map((_,e) => e.value).get();
			var categ = document.getElementById("Cat").value;
			var area = $("#Area :selected").map((_,e) => e.value).get();
			var zona = $("#Zona :selected").map((_,e) => e.value).get();
			var canal = $("#Canal :selected").map((_,e) => e.value).get();
			var fabricante = $("#Fabricante :selected").map((_,e) => e.value).get();
			var marca = $("#Marca :selected").map((_,e) => e.value).get();
			var segmento = $("#Segmento :selected").map((_,e) => e.value).get();
			var tamano = $("#Tamano :selected").map((_,e) => e.value).get();
			var producto = $("#Producto :selected").map((_,e) => e.value).get();
			var indicadores = $("#Indicadores :selected").map((_,e) => e.value).get();
			var semanas = $("#Semanas :selected").map((_,e) => e.value).get();
			var semanasacumuladas = $("#SemanasAcum :selected").map((_,e) => e.value).get();
			
			var columnastotal = 5;
			if (semanasacumuladas.length > 0)
			{
				columnastotal = columnastotal - 1;
			}
			if (semanas.length > columnastotal)
			{
				//alert("Solo se pueden Seleccionar hasta un Maximo de 5 Semanas")
				
				swal("Alerta","Solo se pueden Seleccionar hasta un Maximo de 5 Semanas","error");
				return;
			}
		    $("#cargando").css("display", "block");		
			//alert(categ);
			//alert("fabricante:" + fabricante);
			//alert("marca:" + marca);
			//alert("segmento:" + segmento);
			//return;
			//alert(indicadores);
			//var stodo = "cat=" + categoria;
			var stodo = "cat=" + categ;
			stodo = stodo + "&are=" + area;
			stodo = stodo + "&zon=" + zona;
			stodo = stodo + "&can=" + canal;
			stodo = stodo + "&fab=" + fabricante;
			stodo = stodo + "&mar=" + marca;
			stodo = stodo + "&seg=" + segmento;
			stodo = stodo + "&tam=" + tamano;
			stodo = stodo + "&pro=" + producto;
			stodo = stodo + "&ran=";
			stodo = stodo + "&ind=" + indicadores;
			stodo = stodo + "&sem=" + semanas;
			stodo = stodo + "&semacum=" + semanasacumuladas;
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
