<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_MC_ValBuscarDetallesxProductosxFactura.asp - 26feb21 - 01mar21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idConsumo, rsDetalleProductos, arrDetalleProductos
	'
	idConsumo     = Request.QueryString("id_Consumo")		
	idTipoConsumo = Request.QueryString("id_TipConsumo")	
	'		
	IF ( CInt(idTipoConsumo) = 1 or Cint(idTipoConsumo) = 8 ) then
		'
		' Mercado y Medicinas
		'		
		set rsDetalleProductos			=	CreateObject("ADODB.Recordset")
		rsDetalleProductos.CursorType	=	adOpenKeyset 
		rsDetalleProductos.LockType		=	2 'adLockOptimistic 	
		'
		sql = vbnullstring	
		sql = sql & " SELECT"
		sql = sql & " PH_Consumo_Detalle_Productos.Id_Consumo,"						
		sql = sql & " PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos,"		
		sql = sql & " PH_Consumo_Detalle_Productos.Tipo_codigo_barras,"				
		sql = sql & " PH_Consumo_Detalle_Productos.Numero_codigo_barras,"			
		sql = sql & " PH_Consumo_Detalle_Productos.Cantidad,"						
		sql = sql & " PH_Consumo_Detalle_Productos.Precio_producto,"				
		sql = sql & " PH_Consumo_Detalle_Productos.Moneda,"							
		sql = sql & " PH_Consumo_Detalle_Productos.id_Moneda,"						
		sql = sql & " PH_Consumo_Detalle_Productos.Tasa_de_cambio,"					
		sql = sql & " PH_CB_Producto.Producto AS descripcion,"						
		sql = sql & " PH_CB_Producto.CodigoBarra,"
		sql = sql & " PH_Consumo_Detalle_Productos.Validado,"
		sql = sql & " PH_Consumo_Detalle_Productos.Pendiente,"
		sql = sql & " PH_Consumo_Detalle_Productos.id_categoria,"
		sql = sql & " CASE WHEN PH_Consumo_Detalle_Productos.Id_Moneda = 2"
		sql = sql & " THEN"
		sql = sql & " (PH_Consumo_Detalle_Productos.Cantidad * PH_Consumo_Detalle_Productos.Precio_producto )"
		sql = sql & " ELSE"
		sql = sql & " (PH_Consumo_Detalle_Productos.Cantidad * (PH_Consumo_Detalle_Productos.Precio_producto * PH_Consumo_Detalle_Productos.tasa_de_cambio ) ) END AS total"
		sql = sql & " FROM"
		sql = sql & " PH_Consumo_Detalle_Productos"
		sql = sql & " LEFT JOIN PH_CB_Producto ON PH_Consumo_Detalle_Productos.Numero_codigo_barras = PH_CB_Producto.CodigoBarra AND PH_CB_Producto.Ind_activo = 1"
		sql = sql & " WHERE"
		sql = sql & " PH_Consumo_Detalle_Productos.Id_Consumo = " & idConsumo
		'
		' Response.Write sql
		' Response.End
		'
		rsDetalleProductos.Open sql, conexion
		'
		if not rsDetalleProductos.eof then
			arrDetalleProductos = rsDetalleProductos.GetRows()  ' Convert recordset to 2D Array				
		end if
			'
		rsDetalleProductos.Close
		Set rsDetalleProductos = Nothing
		'
		if IsArray(arrDetalleProductos) then
		''
%>	
		<div class="container-fluid">  
		
			<div class="form-horizontal">

				<table class="table table-hover table-fixed table-striped table-condensed" cellspacing="0">		
				
					<thead class="thead-light">
				 
						<tr>
							<th class="text-center" title="Item Nro">#</th>						
							<th class="text-center" title="Tipo Codigo">Tipo Barras</th>
							<th class="text-center" title="Codigo Barras">C&oacute;digo Barras</th>
							<th class="text-left"   title="Pregunta">Descripci&oacute;n</th>
							<th class="text-center" title="Cantidad">Cantidad</th>
							<th class="text-right"  title="Precio">Precio Unitario</th>
							<th class="text-right"  title="Tasa de Cambio">Tasa Cambio</th>
							<th class="text-right"  title="Total Compra">Total Compra</th>					
							<th class="text-center"  title="Mondea de Pago">Moneda</th>
							
						</tr>
					
					</thead>			
					<tbody>			
<%
						Dim TotalCantidad, TotalPrecio, TotalCompra
						TotalCantidad=TotalPrecio=TotalCmpra=0
						'
						For i = 0 to ubound(arrDetalleProductos, 2)
						
							idConsumoDetalle	= arrDetalleProductos(1,i)	
							tipoCodigoBarras	= arrDetalleProductos(2,i)
							nroCodigoBarras		= arrDetalleProductos(3,i)
							cAntidad			= arrDetalleProductos(4,i)
							precioUnitario		= FormatNumber(arrDetalleProductos(5,i),2)
							idCategoria = CInt(arrDetalleProductos(13,i))
							''
							if(arrDetalleProductos(6,i)="" or isNull(arrDetalleProductos(6,i))) then
								monedaCambio	= "Sin Moneda"
							else
								monedaCambio	= arrDetalleProductos(6,i)						
							end if
							'
							if(arrDetalleProductos(8,i)="" or isNull(arrDetalleProductos(8,i))) then
								tasaCambio  	= "Sin Tasa"
							else
								tasaCambio		= FormatNumber(arrDetalleProductos(8,i),2)						
							end if						
							'
							' if(arrDetalleProductos(13,i)="" or isNull(arrDetalleProductos(14,i))) then
								' Total  	= 0
							' else
								' Total	= FormatNumber(arrDetalleProductos(14,i),2)						
							' end if	
							'							
							'01mar21
							if(arrDetalleProductos(13,i)="" or isNull(arrDetalleProductos(14,i))) then
								Total  	= 0
							else							
								if (idCategoria =9) then
									'Queso
									Total	= precioUnitario					
								else
									Total	= FormatNumber(arrDetalleProductos(14,i),2)										
								end if																								
							end if	
							''
							'monedaCambio		= arrDetalleProductos(6,i)						
							'tasaCambio			= FormatNumber(arrDetalleProductos(8,i),2)
							'
							dEscripcion 		= arrDetalleProductos(9,i)						
							Validado			= arrDetalleProductos(11,i)
							Pendiente			= arrDetalleProductos(12,i)
							Total 				= FormatNumber(Total,2)
							'
							' Calculos Totales
							'
							if (idCategoria = 9) then
								'Queso
								TotalCantidad = TotalCantidad + 1
							else
								TotalCantidad = TotalCantidad + Cint(arrDetalleProductos(4,i))
							end if
							'
							TotalPrecio = TotalPrecio + CDBl(arrDetalleProductos(5,i))
							TotalCompra = TotalCompra + CDBl(total)
							'							
							if( idCategoria > 0) then
								dEscripcion	= buscarCategoria(idCategoria)
							end if
							'									
%>
						<tr class="data">							
							<td class="text-center"><%=i+1%>
							<td class="text-center"><%=tipoCodigoBarras%></td>
							<td class="text-center"><%=nroCodigoBarras%></td>
							<td class="text-left"><%=dEscripcion%></td>
							<td class="text-center"><%=cAntidad%></td>
							<td class="text-right"><%=precioUnitario%></td>
							<td class="text-right"><%=tasaCambio%></td>
							<td class="text-right"><%=Total%></td>							
							<td class="text-center"><%=monedaCambio%></td>
							<!---->
							
							<!---->																		
							
							
						</tr> 
<%					next %>
						<tr>
							<td colspan="4" class="text-right text-primary"><strong>TOTALES:</strong></td>
							<td class="text-center text-primary"><strong><%=FormatNumber(TotalCantidad,0)%></strong></td>
							<td class="text-right text-primary"><strong><%=FormatNumber(TotalPrecio,2)%><strong></td>
							<td></td>
							<td class="text-right text-primary"><strong><%=FormatNumber(TotalCompra,2)%></strong></td>
							<td colspan="3"></td>							
						</tr>
				</tbody>
				
			</table>
					
		</div>	
	
	</div>
<% else %>
	<div class="container-fluid">
		<div class="form-horizontal">
			<table class="table table-hover table-fixed table-striped table-bordered table-condensed" cellspacing="0">
				<thead class="thead-light">
					<tr>
						<th class="text-center" title="Item Nro">#</th>
						<th class="text-center" title="Tipo Codigo">Tipo Barras</th>
						<th class="text-center" title="Codigo Barras">C&oacute;digo Barras</th>
						<th class="text-left"   title="Pregunta">Descripci&oacute;n</th>
						<th class="text-center" title="Cantidad">Cantidad</th>
						<th class="text-right"  title="Precio">Precio Unitario</th>
						<th class="text-right"  title="Tasa de Cambio">Tasa Cambio</th>
						<th class="text-right"  title="Total Compra">Total Compra</th>					
						<th class="text-right"  title="Mondea de Pago">Moneda</th>
					</tr>
				</thead>
				<tbody>			
					<tr>
						<td class="text-center text-danger" colspan="9"><h4>NO HAY REGISTRO DE PRODUCTOS DETALLADOS CON ESTA FACTURA..!</h4></td>			
					</tr>			
				</tbody>
			</table>
		</div>	
	</div>
<%
	end if
	
ELSE
	'Comida-Electrodomesticos-Vehiculos-juguetes
	'
	set rsDetalleProductos			=	CreateObject("ADODB.Recordset")
	rsDetalleProductos.CursorType	=	adOpenKeyset 
	rsDetalleProductos.LockType		=	2 'adLockOptimistic 	
	'
	sql = vbnullstring	
	sql = sql & " SELECT"
	sql = sql & " PH_Consumo.Id_Consumo,"
	sql = sql & " PH_TipoComida.Comida,"
	sql = sql & " PH_Consumo.Nombre_local,"
	sql = sql & " PH_Consumo.Total_Compra,"
	sql = sql & " PH_Moneda.Moneda" 
	sql = sql & " FROM"
	sql = sql & " PH_Consumo"
	sql = sql & " INNER JOIN PH_TipoComida ON PH_Consumo.Id_TipoComida = PH_TipoComida.Id_TipoComida"
	sql = sql & " INNER JOIN cacevedo_atenas.PH_Moneda ON cacevedo_atenas.PH_Consumo.Id_Moneda = cacevedo_atenas.PH_Moneda.Id_Moneda"
	sql = sql & " WHERE"
	sql = sql & " PH_Consumo.Id_Consumo = " & idConsumo
	'
	rsDetalleProductos.Open sql, conexion
	'
	if not rsDetalleProductos.eof then
		arrDetalleProductos = rsDetalleProductos.GetRows()  ' Convert recordset to 2D Array				
	end if
		'
	rsDetalleProductos.Close
	Set rsDetalleProductos = Nothing
	'
	if IsArray(arrDetalleProductos) then
%>
		<div class="container-fluid">
			<div class="form-horizontal">
				 
				<!--<button type="button" title="Agregar Producto" class="btn btn-primary btn-sm" id="agregarProd"  onclick="agregarProducto();" ><i class="fas fa-plus"></i>&nbsp;Agregar Productos</button>-->
				<button type="button" title="Eliminar Consumo" class="btn btn-danger  btn-sm" id="eliminarProd" onclick="eliminarProducto();"><i class="fas fa-times"></i>&nbsp;Eliminar Consumo</button>
				<button type="button" title="Validar Masivo"   class="btn btn-success btn-sm" id="validarMas"   onclick="validarMasivo();"><i class="fas fa-check-double"></i>&nbsp;Validar</button>
				<button type="button" title="Deshacer Masivo"  class="btn btn-warning btn-sm" id="deshacerMas"  onclick="deshacerMasivo();"><i class="fas fa-undo"></i>&nbsp;Deshacer Validar</button>
				<br><br>
				<table class="table table-hover table-fixed table-striped table-condensed" cellspacing="0">		
				
					<thead class="thead-light">
				 
						<tr>
							<th class="text-center" title="Item Nro">#</th>						
							<th class="text-center" title="Tipo Comida">Tipo Comida</th>
							<th class="text-center" title="Codigo Barras">Nombre Local</th>							
							<th class="text-right"  title="Total Compra">Total Compra</th>					
							<th class="text-center" title="Mondea de Pago">Moneda</th>							
							<th class="text-center" title="Editar">Acci&oacute;nes</th>
						</tr>
					
					</thead>			
					<tbody>		
<%
						'
						For i = 0 to ubound(arrDetalleProductos, 2)
						
							idConsumo      	= arrDetalleProductos(0,i)	
							tipoComida		= arrDetalleProductos(1,i)
							nombreLocal		= arrDetalleProductos(2,i)							
							'
							if(arrDetalleProductos(4,i)="" or isNull(arrDetalleProductos(4,i))) then
								monedaCambio	= "Sin Moneda"
							else
								monedaCambio	= arrDetalleProductos(4,i)						
							end if
							'							
							if(arrDetalleProductos(3,i)="" or isNull(arrDetalleProductos(3,i))) then
								Total  	= 0
							else
								Total	= FormatNumber(arrDetalleProductos(3,i),2)						
							end if						
							'								
%>	
						<tr>							
							<td class="text-center"><%=i+1%></td>
							<td class="text-center"><%=tipoComida%></td>
							<td class="text-center"><%=nombreLocal%></td>							
							<td class="text-right"><%=Total%></td>							
							<td class="text-center"><%=monedaCambio%></td>
							<td class="text-center" >
								<button type="button" title="Editar"   class="btn btn-info btn-xs"    id="submiteditar" onclick="obtener_DetallexProducto('<%=idConsumo%>');" ><i class="fas fa-edit"></i></button>																
							</td>							
						</tr>	
<%					next %>	
				</tbody>
				
			</table>
					
		</div>	
	
	</div>
<% else %>

	<div class="container-fluid">
		<div class="form-horizontal">
			<table class="table table-hover table-fixed table-striped table-bordered table-condensed" cellspacing="0">
				<thead class="thead-light">				 
					<tr>
						<th class="text-center" title="Item Nro">#</th>						
						<th class="text-center" title="Tipo Comida">Tipo Comida</th>
						<th class="text-center" title="Codigo Barras">Nombre Local</th>							
						<th class="text-right"  title="Total Compra">Total Compra</th>					
						<th class="text-center" title="Mondea de Pago">Moneda</th>							
						<th class="text-center" title="Editar">Acci&oacute;nes</th>
					</tr>					
					</thead>	
				<tbody>			
					<tr>
						<td class="text-center text-danger" colspan="9"><h4>NO HAY REGISTRO DE PRODUCTOS DETALLADOS CON ESTA FACTURA..!</h4></td>			
					</tr>			
				</tbody>
			</table>
		</div>	
	</div>
<%
	end if
	
END IF	
	
FUNCTION buscarCategoria(byval id)

	Dim QrySql, rsSql
	'
	' Buscar la categoria del campo otro producto en Registro consumos
	'
	QrySql = vbnullstring
	QrySql = " SELECT categoria FROM PH_Categoria WHERE Id_categoria = " & id
	Set rsSql = Server.CreateObject("ADODB.recordset")
	rsSql.Open QrySql, conexion
	'
	if not (rsSql.EOF and rsSql.BOF) then
	   categoria = rsSql(0)
	end if
	'
	rsSql.close
	set rsSql= Nothing	
	'
	buscarCategoria = TRIM(categoria)
	'
END FUNCTION
%>
	
	
	
