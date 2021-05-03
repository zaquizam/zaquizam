<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_ValCalcularResumenSemanalResuelto.asp // 20ene21 - 
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet  = "utf-8"
	'	
	Dim idHogar, idTipCons, sql
	Dim idSemana, idSemana1, idSemana2, idSemana3, idSemana4
	Dim rsSemana, rsSemana1, arrSemana1, rsSemana2, arrSemana2
	Dim rsSemana3, arrSemana3, rsSemana4, arrSemana4
	Dim precio, precio1, precio2, precio3, precio4
	Dim cantidad, cantidad1, cantidad2, cantidad3, cantidad4
	Dim variacion, variacion1, variacion2, variacion3, variacion4
	'
	variacion = variacion1 = variacion2 = variacion3 = variacion4 = 0
	''	
	idSemana	= Request.QueryString("id_semana")
	idHogar		= Request.QueryString("id_Hogar")
	idConsumo 	= Request.QueryString("id_Consumo")
	'
	' Buscar el Tipo de Consumo seleccionado
	'
	set rsTipoConsumo			=	CreateObject("ADODB.Recordset")
	rsTipoConsumo.CursorType	=	adOpenKeyset 
	rsTipoConsumo.LockType		=	2 'adLockOptimistic 	
	'
	sql = vbnullstring	
	sql = sql & " SELECT id_tipoconsumo FROM PH_Consumo WHERE PH_Consumo.id_Consumo = " & idConsumo	
	'	
	' Response.Write sql
	' Response.End
	'
    rsTipoConsumo.Open sql, conexion
	'
	if not rsTipoConsumo.EOF then		
		idTipCons = CInt(rsTipoConsumo("id_tipoconsumo"))		
	 else
		 idTipCons = 0
	end if	
	'
	rsTipoConsumo.Close
	Set rsTipoConsumo = Nothing
	'
	' Calcular los Resultados de las ultimas 5 semanas
	'
	' Semana Actual
	'
	set rsSemana		=	CreateObject("ADODB.Recordset")
	rsSemana.CursorType	=	adOpenKeyset 
	rsSemana.LockType	=	2 'adLockOptimistic 	
	'
	sql = vbnullstring	
	sql = sql & " SELECT"
	sql = sql & " Count(PH_Consumo_Detalle_Productos.Cantidad) as Cantidad,"
	sql = sql & " Sum(PH_Consumo_Detalle_Productos.Precio_producto) as Precio"
	sql = sql & " FROM"
	sql = sql & " PH_Consumo"
	sql = sql & " INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar"	
	sql = sql & " AND PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo"
	sql = sql & " WHERE"
	sql = sql & " PH_Consumo.Id_Hogar = " & idHogar
	sql = sql & " AND"
	sql = sql & " PH_Consumo.id_TipoConsumo = " & idTipCons
	sql = sql & " AND"
	sql = sql & " PH_Consumo.Id_Semana = " & idSemana
	sql = sql & " GROUP BY"
	sql = sql & " PH_Consumo.Id_Hogar,"
	sql = sql & " PH_Consumo.id_TipoConsumo,"
	sql = sql & " PH_Consumo.Id_Semana"
	'	
	' Response.Write sql
	' Response.End
	'
    rsSemana.Open sql, conexion
	'
	if not rsSemana.EOF then		
		cantidad = rsSemana("cantidad")
		precio   = rsSemana("precio")		
	 else
		 precio  = 0
		 cantidad = 0	
	end if
	'
	rsSemana.Close
	Set rsSemana = Nothing
	'
	' Buscar semana
	'
	set rsSemana		=	CreateObject("ADODB.Recordset")
	rsSemana.CursorType	=	adOpenKeyset 
	rsSemana.LockType	=	2 'adLockOptimistic 	
	'
	sql = vbnullstring	
	sql = sql & " SELECT semana FROM ss_semana WHERE idsemana=" & idSemana
	'
    rsSemana.Open sql, conexion
	'
	if not rsSemana.EOF then		
		semana= rsSemana("semana")	
	end if
	'
	rsSemana.Close
	Set rsSemana = Nothing	
	'
	' Primera Semana
	'
	idSemana1 = idSemana - 1
	'
	set rsSemana1			=	CreateObject("ADODB.Recordset")
	rsSemana1.CursorType	=	adOpenKeyset 
	rsSemana1.LockType		=	2 'adLockOptimistic 	
	'
	sql = vbnullstring	
	sql = sql & " SELECT"
	sql = sql & " Count(PH_Consumo_Detalle_Productos.Cantidad) as Cantidad,"
	sql = sql & " Sum(PH_Consumo_Detalle_Productos.Precio_producto) as Precio"
	sql = sql & " FROM"
	sql = sql & " PH_Consumo"
	sql = sql & " INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar"
	sql = sql & " AND PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo"
	sql = sql & " WHERE"
	sql = sql & " PH_Consumo.Id_Hogar = " & idHogar
	sql = sql & " AND"
	sql = sql & " PH_Consumo.id_TipoConsumo = " & idTipCons
	sql = sql & " AND"
	sql = sql & " PH_Consumo.Id_Semana = " & idSemana1
	sql = sql & " GROUP BY"
	sql = sql & " PH_Consumo.Id_Hogar,"
	sql = sql & " PH_Consumo.id_TipoConsumo,"
	sql = sql & " PH_Consumo.Id_Semana"
	'	
	'Response.Write sql
	' Response.End
	'
    rsSemana1.Open sql, conexion
	'
	if not rsSemana1.EOF then		
		cantidad1 = rsSemana1("cantidad")
		precio1   = rsSemana1("precio")		
	 else
		 precio1  = 0
		 cantidad1 = 0	
	end if
		'
	rsSemana1.Close
	Set rsSemana1 = Nothing
	'
	' Buscar semana
	'
	set rsSemana1			=	CreateObject("ADODB.Recordset")
	rsSemana1.CursorType	=	adOpenKeyset 
	rsSemana1.LockType		=	2 'adLockOptimistic 	
	'
	sql = vbnullstring	
	sql = sql & " SELECT semana FROM ss_semana WHERE idsemana=" & idSemana1
	'
    rsSemana1.Open sql, conexion
	'
	if not rsSemana1.EOF then		
		semana1= rsSemana1("semana")	
	end if
	'
	rsSemana1.Close
	Set rsSemana1 = Nothing	
	'
	' Segunda Semana
	'
	idSemana2 = idSemana - 2
	'
	set rsSemana2			=	CreateObject("ADODB.Recordset")
	rsSemana2.CursorType	=	adOpenKeyset 
	rsSemana2.LockType		=	2 'adLockOptimistic 	
	'
	sql = vbnullstring	
	sql = sql & " SELECT"
	sql = sql & " Count(PH_Consumo_Detalle_Productos.Cantidad) as Cantidad,"
	sql = sql & " Sum(PH_Consumo_Detalle_Productos.Precio_producto) as Precio"
	sql = sql & " FROM"
	sql = sql & " PH_Consumo"
	sql = sql & " INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar"	
	sql = sql & " AND PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo"
	sql = sql & " WHERE"
	sql = sql & " PH_Consumo.Id_Hogar = " & idHogar
	sql = sql & " AND"
	sql = sql & " PH_Consumo.id_TipoConsumo = " & idTipCons
	sql = sql & " AND"
	sql = sql & " PH_Consumo.Id_Semana = " & idSemana2
	sql = sql & " GROUP BY"
	sql = sql & " PH_Consumo.Id_Hogar,"
	sql = sql & " PH_Consumo.id_TipoConsumo,"
	sql = sql & " PH_Consumo.Id_Semana"
	'
	' Response.Write sql
	' Response.End
	'
    rsSemana2.Open sql, conexion
	'
	if not rsSemana2.EOF then		
		cantidad2 = rsSemana2("cantidad")
		precio2   = rsSemana2("precio")		
	 else
		 precio2  = 0
		 cantidad2 = 0	
	end if
	'
	rsSemana2.Close
	Set rsSemana2 = Nothing
	'
	' Buscar semana
	'
	set rsSemana2			=	CreateObject("ADODB.Recordset")
	rsSemana2.CursorType	=	adOpenKeyset 
	rsSemana2.LockType		=	2 'adLockOptimistic 	
	'
	sql = vbnullstring	
	sql = sql & " SELECT semana FROM ss_semana WHERE idsemana=" & idSemana2
	'
    rsSemana2.Open sql, conexion
	'
	if not rsSemana2.EOF then		
		semana2= rsSemana2("semana")	
	end if
	'
	rsSemana2.Close
	Set rsSemana2 = Nothing	
	'
	' Tercera Semana
	'
	idSemana3 = idSemana - 3
	'
	set rsSemana3			=	CreateObject("ADODB.Recordset")
	rsSemana3.CursorType	=	adOpenKeyset 
	rsSemana3.LockType		=	2 'adLockOptimistic 	
	'
	sql = vbnullstring	
	sql = sql & " SELECT"
	sql = sql & " Count(PH_Consumo_Detalle_Productos.Cantidad) as Cantidad,"
	sql = sql & " Sum(PH_Consumo_Detalle_Productos.Precio_producto) as Precio"
	sql = sql & " FROM"
	sql = sql & " PH_Consumo"
	sql = sql & " INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar"	
	sql = sql & " AND PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo"
	sql = sql & " WHERE"
	sql = sql & " PH_Consumo.Id_Hogar = " & idHogar
	sql = sql & " AND"
	sql = sql & " PH_Consumo.id_TipoConsumo = " & idTipCons
	sql = sql & " AND"
	sql = sql & " PH_Consumo.Id_Semana = " & idSemana3
	sql = sql & " GROUP BY"
	sql = sql & " PH_Consumo.Id_Hogar,"
	sql = sql & " PH_Consumo.id_TipoConsumo,"
	sql = sql & " PH_Consumo.Id_Semana"
	'
	'Response.Write sql
	' Response.End
	'
    rsSemana3.Open sql, conexion
	'
	if not rsSemana3.EOF then		
		cantidad3 = rsSemana3("cantidad")
		precio3   = rsSemana3("precio")		
	 else
		 precio3  = 0
		 cantidad3 = 0	
	end if
	'
	rsSemana3.Close
	Set rsSemana3 = Nothing
	'
	' Buscar semana
	'
	set rsSemana3			=	CreateObject("ADODB.Recordset")
	rsSemana3.CursorType	=	adOpenKeyset 
	rsSemana3.LockType		=	2 'adLockOptimistic 	
	'
	sql = vbnullstring	
	sql = sql & " SELECT semana FROM ss_semana WHERE idsemana=" & idSemana3
	'
    rsSemana3.Open sql, conexion
	'
	if not rsSemana3.EOF then		
		semana3= rsSemana3("semana")	
	end if
	'
	rsSemana3.Close
	Set rsSemana3 = Nothing	
	'
	' Cuarta Semana
	'
	idSemana4 = idSemana - 4
	'
	set rsSemana4			=	CreateObject("ADODB.Recordset")
	rsSemana4.CursorType	=	adOpenKeyset 
	rsSemana4.LockType		=	2 'adLockOptimistic 	
	'
	sql = vbnullstring	
	sql = sql & " SELECT"
	sql = sql & " Count(PH_Consumo_Detalle_Productos.Cantidad) as Cantidad,"
	sql = sql & " Sum(PH_Consumo_Detalle_Productos.Precio_producto) as Precio"
	sql = sql & " FROM"
	sql = sql & " PH_Consumo"
	sql = sql & " INNER JOIN PH_Consumo_Detalle_Productos ON PH_Consumo.Id_Hogar = PH_Consumo_Detalle_Productos.Id_Hogar"	
	sql = sql & " AND PH_Consumo.Id_Consumo = PH_Consumo_Detalle_Productos.Id_Consumo"
	sql = sql & " WHERE"
	sql = sql & " PH_Consumo.Id_Hogar = " & idHogar
	sql = sql & " AND"
	sql = sql & " PH_Consumo.id_TipoConsumo = " & idTipCons
	sql = sql & " AND"
	sql = sql & " PH_Consumo.Id_Semana = " & idSemana4
	sql = sql & " GROUP BY"
	sql = sql & " PH_Consumo.Id_Hogar,"
	sql = sql & " PH_Consumo.id_TipoConsumo,"
	sql = sql & " PH_Consumo.Id_Semana"
	'
	' Response.Write sql
	' Response.End
	'
    rsSemana4.Open sql, conexion
	'
	if not rsSemana4.EOF then		
		cantidad4 = rsSemana4("cantidad")
		precio4   = rsSemana4("precio")		
	 else
		 precio4  = 0
		 cantidad4 = 0	
	end if
	'
	rsSemana4.Close
	Set rsSemana4 = Nothing
	'
	' Buscar semana
	'
	set rsSemana4			=	CreateObject("ADODB.Recordset")
	rsSemana4.CursorType	=	adOpenKeyset 
	rsSemana4.LockType		=	2 'adLockOptimistic 	
	'
	sql = vbnullstring	
	sql = sql & " SELECT semana FROM ss_semana WHERE idsemana=" & idSemana4
	'
    rsSemana4.Open sql, conexion
	'
	if not rsSemana4.EOF then		
		semana4= rsSemana4("semana")	
	end if
	'
	rsSemana4.Close
	Set rsSemana4 = Nothing	
	'
	''
	' CALCULO DE VARIACIONES
	''
	PromedioMonto = 0
	PromedioMonto = ( CDbl(precio1) + CDbl(precio2) + CDbl(precio3) + CDbl(precio4) ) / 4
	
	' Response.Write " Prom: " & PromedioMonto & "<br>"
	'response.end
	
	VariacionMonto = 0 
	' VariacionMonto = (PromedioMonto - precio)
	' VariacionMonto = (VariacionMonto  / PromedioMonto)
	' VariacionMonto = (VariacionMonto  * 100)
	'
	if(PromedioMonto>0) then
		VariacionMonto = ( CDbl(precio) - CDbl(PromedioMonto) ) / CDbl(PromedioMonto) * 100
	else
		VariacionMonto = 0
	end if	
	''
	PromedioUnidades = 0
	PromedioUnidades = ( CDbl(cantidad1) + CDbl(cantidad2) + CDbl(cantidad3) + CDbl(cantidad4) ) / 4
	
	' Response.Write " Prom: " & PromedioMonto & "<br>"
	'response.end
	
	VariacionUnidades = 0 
	' VariacionMonto = (PromedioMonto - precio)
	' VariacionMonto = (VariacionMonto  / PromedioMonto)
	' VariacionMonto = (VariacionMonto  * 100)
	'
	if(PromedioUnidades>0) then
		VariacionUnidades = ( CDbl(cantidad) - CDbl(PromedioUnidades) ) / CDbl(PromedioUnidades) * 100
	else
		VariacionUnidades = 0
	end if	
	'
	' Response.Write "paso"
	'response.End
	'
%>	
	<div class="form-horizontal">
	
		<table class="table table-hover table-fixed table-striped table-condensed " cellspacing="0">		
			<thead>
			<tr>
				<th colspan="2" class="text-center info">Sem: <%=Semana%> </th>
				<th class="default"></th>
				<th colspan="2" class="text-center info">Sem: <%=Semana1%></th>
				<th class="default"></th>
				<th colspan="2" class="text-center info">Sem: <%=Semana2%></th>
			</tr>
			</thead>
			<tbody>
				<tr>
					<td class="text-center">Monto</td>
					<td class="text-center">Unidades</td>					
					<td class="info"></td>
					<td class="text-center">Monto</td>
					<td class="text-center">Unidades</td>
					<td class="info"></td>
					<td class="text-center">Monto</td>
					<td class="text-center">Unidades</td>
				</tr>			   
				<tr>
					<td class="text-center"><%=FormatNumber(precio,2)%></td>
					<td class="text-center"><%=FormatNumber(cantidad,0)%></td>
					<td class="info"></td>
					<td class="text-center"><%=FormatNumber(precio1,2)%></td>
					<td class="text-center"><%=FormatNumber(cantidad1,0)%></td>
					<td class="info"></td>
					<td class="text-center"><%=FormatNumber(precio2,2)%></td>
					<td class="text-center"><%=FormatNumber(cantidad2,0)%></td>
				</tr>               
			</tbody>
		</table>
		<br>
		<table class="table table-hover table-fixed table-striped table-condensed" cellspacing="0">		
			<thead>
			<tr>
				<th colspan="2" class="text-center info">Sem: <%=Semana3%> </th>
				<th class="default"></th>
				<th colspan="2" class="text-center info">Sem: <%=Semana4%></th>
			</tr>
			</thead>
			<tbody>
				<tr>
					<td class="text-center">Monto</td>
					<td class="text-center">Unidades</td>
					<td class="info"></td>
					<td class="text-center">Monto</td>
					<td class="text-center">Unidades</td>         
				</tr>			   
				<tr>
					<td class="text-center"><%=FormatNumber(precio3,2)%></td>
					<td class="text-center"><%=FormatNumber(cantidad3,0)%></td>
					<td class="info"></td>
					<td class="text-center"><%=FormatNumber(precio4,2)%></td>
					<td class="text-center"><%=FormatNumber(cantidad4,0)%></td>
				</tr>               
			</tbody>
		</table>
		<!-- VARIACION SEMANAL -->
		<h4><i class="fas fa-tachometer-alt"></i><strong>&nbsp;Porcentajes de Variaci&oacute;n:</strong></h4>
		<table class="table table-hover table-fixed table-striped table-condensed " cellspacing="0">		
			<thead>
			<tr>
				<th colspan="7" class="text-center warning">VARIACI&Oacute;N EN MONTOS</th>				
			</tr>
			</thead>
			<tbody>
				<tr>
					<td class="text-center">Sem: <%=left(Semana1,4)%> </td>
					<td class="text-center">Sem: <%=left(Semana2,4)%> </td>
					<td class="text-center">Sem: <%=left(Semana3,4)%> </td>
					<td class="text-center">Sem: <%=left(Semana4,4)%> </td>
					<td class="text-center text-primary"><strong>Promedio</strong></td>
					<td class="text-center">Sem: <%=left(Semana,4)%> </td>					
					<% if (VariacionMonto<0) then %>
						<td class="text-center text-danger"><strong>% Variaci&oacute;n<strong></td>					
					<% else %>
						<td class="text-center text-primary"><strong>% Variaci&oacute;n<strong></td>
					<% end if%>										
				</tr>			   
				<tr>
					<td class="text-center"><%=FormatNumber(precio1,2)%></td>
					<td class="text-center"><%=FormatNumber(precio2,0)%></td>
					<td class="text-center"><%=FormatNumber(precio3,0)%></td>
					<td class="text-center"><%=FormatNumber(precio4,2)%></td>
					<td class="text-center text-primary"><strong><%=FormatNumber(PromedioMonto,2)%></strong></td>
					<td class="text-center"><%=FormatNumber(Precio,2)%></td>					
					<% if (VariacionMonto<0) then %>
						<td class="text-center text-danger"><strong><%=FormatNumber(VariacionMonto,2)%></strong></td>
					<% else %>
						<td class="text-center text-primary"><strong><%=FormatNumber(VariacionMonto,2)%></strong></td>
					<% end if%>						
				</tr>               
			</tbody>
		</table>
		<br>
		<table class="table table-hover table-fixed table-striped table-condensed " cellspacing="0">		
			<thead>
			<tr>
				<th colspan="7" class="text-center warning">VARIACI&Oacute;N EN UNIDADES</th>				
			</tr>
			</thead>
			<tbody>
				<tr>
					<td class="text-center">Sem: <%=left(Semana1,4)%> </td>
					<td class="text-center">Sem: <%=left(Semana2,4)%> </td>
					<td class="text-center">Sem: <%=left(Semana3,4)%> </td>
					<td class="text-center">Sem: <%=left(Semana4,4)%> </td>
					<td class="text-center text-primary"><strong>Promedio</strong></td>
					<td class="text-center">Sem: <%=left(Semana,4)%> </td>					
					<% if (VariacionUnidades<0) then %>
						<td class="text-center text-danger"><strong>% Variaci&oacute;n</strong></td>					
					<% else %>
						<td class="text-center text-primary"><strong>% Variaci&oacute;n</strong></td>
					<% end if%>					
				</tr>			   
				<tr>
					<td class="text-center"><%=FormatNumber(cantidad1,2)%></td>
					<td class="text-center"><%=FormatNumber(cantidad2,0)%></td>
					<td class="text-center"><%=FormatNumber(cantidad3,0)%></td>
					<td class="text-center"><%=FormatNumber(cantidad4,2)%></td>
					<td class="text-center text-primary"><strong><%=FormatNumber(PromedioUnidades,2)%></strong></td>
					<td class="text-center"><%=FormatNumber(cantidad,2)%></td>
					<% if (VariacionUnidades<0) then %>
						<td class="text-center text-danger"><strong><%=FormatNumber(VariacionUnidades,2)%><strong></td>
					<% else %>
						<td class="text-center text-primary"><strong><%=FormatNumber(VariacionUnidades,2)%><strong></td>
					<% end if%>		
				</tr>               
			</tbody>
		</table>
		
		
		
		
	</div>	

	
	
	
