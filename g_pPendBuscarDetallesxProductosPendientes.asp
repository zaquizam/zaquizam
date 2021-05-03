<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_pPendBuscarDetallesxProductosPendientes -- 03mar21 - 17abr21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idCodigoBarras, rsDetalleProductos, arrDetalleProductos, Promedio, PromedioCantidad
	'
	Dim instSql, Sql, precio, rsSql
	Dim dd, mm, yy, hh, nn, ss, Moda
	Dim datevalue, timevalue, dtsnow
	'
	dtsnow = Now()
	dd = Right("00" & Day(dtsnow), 2)
	mm = Right("00" & Month(dtsnow), 2)
	yy = Year(dtsnow)
	hh = Right("00" & Hour(dtsnow), 2)
	nn = Right("00"  & Minute(dtsnow), 2)
	ss = Right("00" & Second(dtsnow), 2)
	datevalue = yy  & "-" & mm & "-" & dd
	timevalue = hh  & ":" & nn & ":" & ss
	sUpdate = datevalue & " " & timevalue
	'
	sIP       = Request.ServerVariables("REMOTE_ADDR")
	'
	' Captura Variables
	'
	idCodigoBarras	= Request.QueryString("idBarcode")		
	'
	 '
    ' Buscar Id de la Semana segun la fecha de consumo
    '    
    dtsnow = Now()
	dd = Right("00" & Day(dtsnow), 2)
	mm = Right("00" & Month(dtsnow), 2)
	yy = Year(dtsnow)
	hh = Right("00" & Hour(dtsnow), 2)
	nn = Right("00" & Minute(dtsnow), 2)
	ss = Right("00" & Second(dtsnow), 2)
	datevalue = yy & "-" & mm & "-" & dd
	timevalue = hh & ":" & nn & ":" & ss
	sUpdate = datevalue & " " & timevalue    
	'	
    QrySql = vbnullstring
    QrySql = QrySql & " SELECT idsemana FROM ss_semana  WHERE '" & datevalue & "' BETWEEN fec_inicio AND fec_fin"
    '    
    Set rsSql = Server.CreateObject("ADODB.recordset")
    rsSql.Open QrySql, conexion
    '
    if not (rsSql.EOF and rsSql.BOF) then
        idSemana=rsSql(0)
    else
        idSemana=0
    end if
    '
    rsSql.close
    set rsSql= Nothing	    
	'
	' Mercado y Medicinas
	'		
	set rsDetalleProductos			=	CreateObject("ADODB.Recordset")
	rsDetalleProductos.CursorType	=	adOpenKeyset 
	rsDetalleProductos.LockType		=	2 'adLockOptimistic 	
	'
	sql = vbnullstring	
	sql = sql & " SELECT"
	sql = sql & " PH_Consumo_Detalle_Productos.Id_Hogar,"
	sql = sql & " PH_GArea.Area,"
	sql = sql & " ss_Estado.Estado,"
	sql = sql & " ss_Semana.Semana,"
	sql = sql & " PH_Consumo_Detalle_Productos.Cantidad,"
	sql = sql & " PH_Consumo_Detalle_Productos.Precio_producto AS precio,"
	sql = sql & " PH_Consumo_Detalle_Productos.Moneda,"
	sql = sql & " PH_Consumo_Detalle_Productos.Id_Consumo_Detalle_Productos,"
	sql = sql & " ss_Semana.idSemana,"
	sql = sql & " PH_Consumo_Detalle_Productos.id_Moneda,"
	sql = sql & " PH_Consumo_Detalle_Productos.tasa_de_cambio,"
	sql = sql & " CASE"
	sql = sql & " WHEN PH_Consumo_Detalle_Productos.Id_Moneda <> 2 THEN"
	sql = sql & " ( PH_Consumo_Detalle_Productos.Precio_producto * PH_Consumo_Detalle_Productos.Tasa_de_cambio )"
	sql = sql & " ELSE PH_Consumo_Detalle_Productos.Precio_producto"
	sql = sql & " END AS Precio_Conversion"
	sql = sql & " FROM"
	sql = sql & " PH_Consumo_Detalle_Productos"
	sql = sql & " INNER JOIN PH_Consumo ON PH_Consumo_Detalle_Productos.Id_Consumo = PH_Consumo.Id_Consumo"
	sql = sql & " INNER JOIN ss_Semana ON PH_Consumo.Id_Semana = ss_Semana.IdSemana"
	sql = sql & " INNER JOIN PH_PanelHogar ON PH_Consumo.Id_Hogar = PH_PanelHogar.Id_PanelHogar"
	sql = sql & " INNER JOIN PH_GAreaEstado ON PH_PanelHogar.Id_Estado = PH_GAreaEstado.Id_Estado"
	sql = sql & " INNER JOIN PH_GArea ON PH_GAreaEstado.Id_Area = PH_GArea.Id_Area"
	sql = sql & " INNER JOIN ss_Estado ON PH_GAreaEstado.Id_Estado = ss_Estado.Id_Estado"
	sql = sql & " WHERE"
	sql = sql & " ss_Semana.idSemana > 14" 
	sql = sql & " AND PH_Consumo_Detalle_Productos.Precio_producto is not null"
	sql = sql & " AND PH_Consumo_Detalle_Productos.tasa_de_cambio is not null"
	sql = sql & " AND PH_Consumo_Detalle_Productos.Cantidad is not null"
	sql = sql & " AND PH_Consumo_Detalle_Productos.Id_Moneda is not null"
	sql = sql & " AND PH_Consumo_Detalle_Productos.Moneda is not null"
	sql = sql & " AND PH_Consumo_Detalle_Productos.Numero_codigo_barras = '" & idCodigoBarras & "'"
	sql = sql & " ORDER BY"
	sql = sql & " Precio_Conversion ASC"
	'
	'Response.Write sql
	'Response.End
	'
	rsDetalleProductos.Open sql, conexion
	'
	if not rsDetalleProductos.eof then
		arrDetalleProductos = rsDetalleProductos.GetRows()  ' Convert recordset to 2D Array				
	end if
		'
	rsDetalleProductos.Close: Set rsDetalleProductos = Nothing
	'
	if IsArray(arrDetalleProductos) then	
		'
		' Borrar datos del calculo de Moda de Productos 
		'
		Sql = vbnullstring
		Sql = Sql & " DELETE FROM PH_Temporal_Productos_Pendientes_Moda WHERE USR='" & Session("Usuario") & "'"
		Set objExec = conexion.Execute(Sql)
		Set objExec = Nothing
		'	
		' Calcular el Promedio de Precios Unificados a Bs.
		'
		Promedio = PromedioCantidad = 0				
		For i = 0 to ubound(arrDetalleProductos, 2)
			'
			Promedio = Promedio + CDbl(arrDetalleProductos(11,i))
			PromedioCantidad = PromedioCantidad + 1
			arrDetalleProductos(6,i)="Bolivar Soberano"
			'
			precio=0
			precio=Replace(arrDetalleProductos(11,i),",",".")
			'
			instSql = vbnullstring	
			instSql = instSql & " INSERT INTO PH_Temporal_Productos_Pendientes_Moda "
			instSql = instSql & " ( Id_Consumo_Detalle_Producto, Precio_producto,"
			instSql = instSql & "  USR, idsession, IP, Fec_Ult_Mod )"
			instSql = instSql & " VALUES "
			instSql = instSql & "(" & arrDetalleProductos(7,i) & ","
			instSql = instSql & ""  & precio & ","
			instSql = instSql & "'" & Session("Usuario")  & "',"			
			instSql = instSql & ""  & Session.SessionID & ","
			instSql = instSql & "'" & sIp & "',"			
			instSql = instSql & "'" & sUpdate & "')"
			'			
			Set objExec = conexion.Execute(instSql)
			'
		Next
		'
		Set objExec = Nothing
		'
		if CInt(PromedioCantidad) > 0 then
			Promedio = FormatNumber(CDbl(Promedio) / CInt(PromedioCantidad),2)		
		end if
		'
		' Calculo de la Moda
		'
		instSql = vbnullstring
		instSql = instSql & " SELECT "
		instSql = instSql & " PH_Temporal_Productos_Pendientes_Moda.Precio_producto,"
		instSql = instSql & " COUNT ( PH_Temporal_Productos_Pendientes_Moda.Id_Consumo_Detalle_Producto ) AS CuentaDeId_Consumo "
		instSql = instSql & " FROM"
		instSql = instSql & " PH_Temporal_Productos_Pendientes_Moda "
		instSql = instSql & " WHERE"
		instSql = instSql & " PH_Temporal_Productos_Pendientes_Moda.USR = '" & Session("Usuario") & "'"
		instSql = instSql & " GROUP BY"
		instSql = instSql & " PH_Temporal_Productos_Pendientes_Moda.Precio_producto"
		instSql = instSql & " ORDER BY"
		instSql = instSql & " COUNT ( PH_Temporal_Productos_Pendientes_Moda.Id_Consumo_Detalle_Producto ) DESC"
		'
		Set rsSql = Server.CreateObject("ADODB.recordset")
		rsSql.Open instSql, conexion
		'
		if not rsSql.EOF then
			arrModa = rsSql.GetRows()  ' Convert recordset to 2D Array
			Moda    = FormatNumber(CDbl(arrModa(0,0)),2)
		else
			Moda = 0
		end if
		'
		rsSql.close : set rsSql= Nothing	
		'		
%>	
		 	
	<div class="form-horizontal">
	
			<div>
				<div class="col-sm-3">
          			<h4 class="text-danger text-center"><strong><i class="fas fa-barcode"></i>&nbsp;<%=idCodigoBarras%></strong></h4>
				</div>	
				<div class="col-sm-3">
          			<h4 class="text-primary text-center"><strong>Promedio:&nbsp;<%=Promedio%></strong></h4>
					<input type="hidden" name="Promedio" id="Promedio" disabled maxlength=15 size=15 value=<%=Promedio%> />					
				</div>	
				<div class="col-sm-3">
          			<h4 class="text-secondary text-center"><strong>Moda:&nbsp;<%=Moda%></strong></h4>
					<input type="hidden" name="Moda" id="Moda" disabled maxlength=15 size=15 value=<%=Moda%> />					
				</div>	
				<div class="col-sm-3">
          			<h4 class="text-danger text-center"><strong><i class="fas fa-money-bill-alt"></i>&nbsp;Precios expresados en Bolivares</strong></h4>
				</div>	
        	</div>	
			<table class="table table-scroll table-striped">
				<input type="hidden" value="" id="Hiddenfield2"/>
				<thead>							 
					<tr>
						<th class="text-center" title="Item Nro">#</th>						
						<th class="text-center" title="ID Hogar">ID Hogar</th>
						<th class="text-center" title="Area">Area</th>
						<th class="text-center" title="Estado">Estado</th>
						<th class="text-center" title="Estado">Semana</th>
						<th class="text-center" title="Cantidad">Cantidad</th>
						<th class="text-center" title="Precio">Precio Bs</th>
						<th class="text-center" title="Mondea de Pago">Moneda</th>							
						<th class="text-center" title="Promedio">Promedio</th>
						<th class="text-center" title="Moda">Moda</th>
						<th class="text-center" title="Valor" colspan="2">Valor Manual</th>											
						<th class="text-center" title="Cantidad" colspan="2">Cantidad Manual</th>											
						<th class="text-center" title="Mostrar" >Chequear</th>											
					</tr>				
				</thead>			
				<tbody>			
<%
					Dim idHogar, sArea, sEstado,sSemana, iCantidad, sMoneda, precioUnitario	
					'
					For i = 0 to ubound(arrDetalleProductos, 2)
						'
						iCantidad = 0
						precioUnitario = 0
						'
						idHogar				= arrDetalleProductos(0,i)	
						sArea				= arrDetalleProductos(1,i)
						sEstado				= arrDetalleProductos(2,i)
						sSemana				= arrDetalleProductos(3,i)
						iCantidad			= arrDetalleProductos(4,i)
						'precioUnitario		= arrDetalleProductos(5,i)
						sMoneda				= arrDetalleProductos(6,i)	
						idConsumoDetalle	= arrDetalleProductos(7,i)
						idSemanaDetalle 	= arrDetalleProductos(8,i)
						idMoneda			= arrDetalleProductos(9,i)	
						TasadeCambio		= arrDetalleProductos(10,i)
						precioUnitario		= arrDetalleProductos(11,i)
						'													
%>
					<tr class="data">							
						<!--<td class="text-center"><%=i+1%></td>-->
						<%if idSemana <> idSemanaDetalle then %>
							<td class="text-center"><%=i+1%>&nbsp;<input type="checkbox" name="CambioMasivo" id="chkbox_<%=idConsumoDetalle%>" value="<%=idConsumoDetalle%>" checked /></td>
						<%else %>
							<td class="text-center"><%=i+1%>&nbsp;<input type="checkbox" name="CambioMasivo" id="chkbox_<%=idConsumoDetalle%>" value="<%=idConsumoDetalle%>" /></td>
						<%end if%>
						<td class="text-center"><%=idHogar%></td>
						<td class="text-center"><%=sArea%></td>
						<td class="text-center"><%=sEstado%></td>
						<td class="text-center"><%=sSemana%></td>							
						<td class="text-center"><%=iCantidad%></td>
						<td  class="text-right"><%=FormatNumber(precioUnitario,2)%></td>
						<td class="text-center">
							<%=sMoneda%>
							<input type="hidden" id="idMon_<%=idConsumoDetalle%>" disabled value="<%=idMoneda%>" />
							<input type="hidden" id="tasa_<%=idConsumoDetalle%>"  disabled value="<%=TasadeCambio%>" />							
							<input type="hidden" id="cant_<%=idConsumoDetalle%>"  disabled value="<%=iCantidad%>" />
						</td>							
						<!---->																		
						<td class="text-center">							
							<img src="images/Boton01.png"  title="Actualizar precio por Promedio" width="24px" height="24px" onclick="ActualizarPromedio('<%=idConsumoDetalle%>')"/>
						</td>						
						<td class="text-center">
							<img src="images/Boton02.png"  title="Actualizar precio por Moda" width="24px" height="24px" onclick="ActualizarModa('<%=idConsumoDetalle%>')"/>
						</td>						
						<td class="text-center" colspan="2">							
							<input type="text" id="valor_<%=idConsumoDetalle%>" maxlength=15 size=15 >
							<img src="images/Boton03.png"  title="Actualizar precio manual" width="24px" height="24px" onclick="ActualizarManual('<%=idConsumoDetalle%>')"/>						
						</td>					
						<td class="text-center" colspan="2">							
							<input type="text" id="cantmod_<%=idConsumoDetalle%>" maxlength=8 size=8 >
							<img src="images/Boton04.png"  title="Actualizar Cantidad manual" width="24px" height="24px" onclick="ActualizarCantidad('<%=idConsumoDetalle%>')"/>						
						</td>					
						<td class="text-center">
							<!--<a href="#" title="Ver Detalle del Registro"  onclick="MostrarDetalleRegistro('<%=idConsumoDetalle%>');" ><i class="fas fa-search"></i></a>-->
							<img src="images/buscarpend2.png"  title="Ver Detalle del Registro" width="24px" height="24px" onclick="MostrarDetalleRegistro('<%=idConsumoDetalle%>')"/>						
						</td>
					</tr> 
					
<%					next %>
				</tbody>
			
			</table>				
		
	</div>	
	
<% else %>
	
	<div class="form-horizontal">
		<table class="table table-hover table-fixed table-striped table-bordered table-condensed" cellspacing="0">
			<thead class="thead-light">
				<tr>
					<tr>
						<th class="text-center" title="Item Nro">#</th>						
						<th class="text-center" title="ID Hogar">ID Hogar</th>
						<th class="text-center" title="Area">Area</th>
						<th class="text-center" title="Estado">Estado</th>
						<th class="text-center" title="Estado">Semana</th>
						<th class="text-center" title="Cantidad">Cantidad</th>
						<th class="text-right"  title="Precio">Precio Unitario</th>
						<th class="text-center"  title="Mondea de Pago">Moneda</th>							
						<th class="text-center" title="Editar">Acci&oacute;nes</th>
					</tr>
				</tr>
			</thead>
			<tbody>			
				<tr>
					<td class="text-center text-danger" colspan="9"><h4>NO HAY REGISTRO DE PRODUCTOS PENDIENTES CON ESE CODIGO DE BARRAS..!</h4></td>			
				</tr>			
			</tbody>
		</table>
	</div>	
	
<%end if%>
	

	
	
	
