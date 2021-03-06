<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_ValBuscarDetallesxProductosxFactura.asp // 02ENE21  - 13ene21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idConsumo, rsDetalleProductos, arrDetalleProductos
	'
	idConsumo   = Request.QueryString("id_Consumo")		
	'	
	' Buscar Resultados
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
	'sql = sql & " (PH_Consumo_Detalle_Productos.Cantidad * (PH_Consumo_Detalle_Productos.Precio_producto * ISNULL(PH_Consumo_Detalle_Productos.tasa_de_cambio,1) ) ) END AS total"
	sql = sql & " FROM"
	sql = sql & " PH_Consumo_Detalle_Productos"
	sql = sql & " LEFT JOIN PH_CB_Producto ON PH_Consumo_Detalle_Productos.Numero_codigo_barras = PH_CB_Producto.CodigoBarra"
	sql = sql & " WHERE"
	sql = sql & " PH_Consumo_Detalle_Productos.Id_Consumo = " & idConsumo
	'sql = sql & " AND"
	'sql = sql & " PH_Consumo_Detalle_Productos.Validado = 0"
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
			 
			<button type="button" title="Agregar Producto" class="btn btn-primary btn-sm" id="agregarProd"  onclick="agregarProducto();" ><i class="fas fa-plus"></i>&nbsp;Agregar Productos</button>
			<button type="button" title="Eliminar Consumo" class="btn btn-danger  btn-sm" id="eliminarProd" onclick="eliminarProducto();"><i class="fas fa-times"></i>&nbsp;Eliminar Consumo</button>
			<button type="button" title="Validar Masivo"   class="btn btn-success btn-sm" id="validarMas"   onclick="validarMasivo();"><i class="fas fa-check-double"></i>&nbsp;Validar Masivo</button>
			<button type="button" title="Deshacer Masivo"  class="btn btn-warning btn-sm" id="deshacerMas"  onclick="deshacerMasivo();"><i class="fas fa-undo"></i>&nbsp;Deshacer Masivo</button>
			<br><br>
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
						<th class="text-center" title="Item Nro">Status</th>
						<th class="text-center" title="Editar">Acci&oacute;nes</th>
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
						''
						if(arrDetalleProductos(13,i)="" or isNull(arrDetalleProductos(14,i))) then
							Total  	= 0
						else
							Total	= FormatNumber(arrDetalleProductos(14,i),2)						
						end if						
						'monedaCambio		= arrDetalleProductos(6,i)						
						'tasaCambio			= FormatNumber(arrDetalleProductos(8,i),2)
						dEscripcion 		= arrDetalleProductos(9,i)						
						Validado			= arrDetalleProductos(11,i)
						Pendiente			= arrDetalleProductos(12,i)
						'Total 				= FormatNumber(arrDetalleProductos(13,i),2)
						Total 				= FormatNumber(Total,2)
						'Calculos Totales
						TotalCantidad = TotalCantidad + Cint(arrDetalleProductos(4,i))
						TotalPrecio = TotalPrecio + CDBl(arrDetalleProductos(5,i))
						TotalCompra = TotalCompra + CDBl(total)
						'
						idCategoria = CInt(arrDetalleProductos(13,i))
						if( idCategoria > 0) then
							dEscripcion	= buscarCategoria(idCategoria)
						end if
						'									
%>
						<tr>							
							<td class="text-center"><%=i+1%></td>
							<td class="text-center"><%=tipoCodigoBarras%></td>
							<td class="text-center"><%=nroCodigoBarras%></td>
							<td class="text-left"><%=dEscripcion%></td>
							<td class="text-center"><%=cAntidad%></td>
							<td class="text-right"><%=precioUnitario%></td>
							<td class="text-right"><%=tasaCambio%></td>
							<td class="text-right"><%=Total%></td>							
							<td class="text-center"><%=monedaCambio%></td>
							<!---->
							<%if(validado=True and Pendiente=False) then%>
								<td class="text-center"><i class="fas fa-check"></i></td>
							<%elseif(validado=False and Pendiente=True) then%>
								<td class="text-center"><i class="fas fa-crosshairs"></i></td>
							<%elseif(validado=True and Pendiente=True) then%>
								<td class="text-center"><i class="fas fa-check"></i></td>
							<%else%>
								<td class="text-center"><i class="fas fa-eye"></i></td>						
							<%end if%>
							<!--
							<td class="text-center" >
								<button type="button" title="Editar"   class="btn btn-info btn-xs"    id="submitedit" onclick="obtener_DetallexProducto('<%=idConsumoDetalle%>');" ><i class="fas fa-edit"></i></button>
								<button type="button" title="Validar"  class="btn btn-success btn-xs" id="submitval" onclick="validar_Directo('<%=idConsumoDetalle%>');" ><i class="fas fa-check"></i></button>
								<button type="button" title="Eliminar" class="btn btn-danger btn-xs"  id="submitdel" onclick="eliminar_Detalle_Producto('<%=idConsumoDetalle%>');" ><i class="fas fa-times"></i></button>
								<button type="button" title="Marcar Pendiente"  class="btn btn-warning btn-xs" id="submitpend" onclick="marcar_Producto_Pendiente('<%=idConsumoDetalle%>');" ><i class="fas fa-crosshairs"></i></button>
								<button type="button" title="Borrar Status"  class="btn btn-primary btn-xs" id="submitReset" onclick="eliminar_Status_Producto('<%=idConsumoDetalle%>');" ><i class="fas fa-redo"></i></button>
							</td>
							-->				
														
							<td class="text-center" >
								<div class="dropdown">
								  <button class="btn btn-info btn-sm dropdown-toggle" type="button" data-toggle="dropdown">Menu&nbsp;<span class="caret"></span> <span class="sr-only">Desplegar men??</span></button>
								  <ul class="dropdown-menu pull-right" role="menu">
								  	<li><a href="#" title="Editar" onclick="obtener_DetallexProducto('<%=idConsumoDetalle%>');"><i class="fas fa-edit"></i>&nbsp;Editar</a></li>
									<li><a href="#" title="Validar"  onclick="validar_Directo('<%=idConsumoDetalle%>');" ><i class="fas fa-check"></i>&nbsp;Validar</a></li>
									<li><a href="#" title="Eliminar" onclick="eliminar_Detalle_Producto('<%=idConsumoDetalle%>');" ><i class="fas fa-times"></i>&nbsp;Eliminar</a></li>
									<li><a href="#" title="Marcar Pendiente" onclick="marcar_Producto_Pendiente('<%=idConsumoDetalle%>');"><i class="fas fa-crosshairs"></i>&nbsp;Pendiente</a></li>
									<li><a href="#" title="Borrar Status"  onclick="eliminar_Status_Producto('<%=idConsumoDetalle%>');" ><i class="fas fa-undo"></i>&nbsp;Deshacer</a></li>									
								  </ul>
								</div>
							</td>
							
						</tr> 
<%
					next
%>
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
<%

	else
%>
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
	
	
	
