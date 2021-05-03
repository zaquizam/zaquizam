<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_ValCalcularResumenSemanal.asp // 09ene21 - 12ene21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idHogar, idTipCons, sql
	Dim idSemana1, idSemana2, idSemana3, idSemana4
	Dim rsSemana1, arrSemana1, rsSemana2, arrSemana2
	Dim rsSemana3, arrSemana3, rsSemana4, arrSemana4
	Dim precio1, precio2, precio3, precio4
	Dim cantidad1, cantidad2, cantidad3, cantidad4
	'	
	idSemana	= Request.QueryString("id_semana")
	idHogar		= Request.QueryString("id_Hogar")
	idTipCons	= Request.QueryString("id_TipCons")
	'	
	' Calcular los Resultados de las ultimas 4 semanas
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
	'Response.Write sql
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
%>	
	<div class="form-horizontal">
	
		<table class="table table-hover table-fixed table-striped table-condensed " cellspacing="0">		
			<thead>
			<tr>
				<th colspan="2" class="text-center warning">Semana: <%=Semana1%> </th>
				<th class="warning"></th>
				<th colspan="2" class="text-center warning">Semana: <%=Semana2%></th>
			</tr>
			</thead>
			<tbody>
				<tr>
					<td class="text-center">Monto</td>
					<td class="text-center">Unidades</td>
					<td class="warning"></td>
					<td class="text-center">Monto</td>
					<td class="text-center">Unidades</td>                  
				</tr>			   
				<tr>
					<td class="text-center"><%=FormatNumber(precio1,2)%></td>
					<td class="text-center"><%=FormatNumber(cantidad1,0)%></td>
					<td class="warning"></td>
					<td class="text-center"><%=FormatNumber(precio2,2)%></td>
					<td class="text-center"><%=FormatNumber(cantidad2,0)%></td>
				</tr>               
			</tbody>
		</table>
		<br>
		<table class="table table-hover table-fixed table-striped table-condensed " cellspacing="0">		
			<thead>
			<tr>
				<th colspan="2" class="text-center info">Semana: <%=Semana3%> </th>
				<th class="info"></th>
				<th colspan="2" class="text-center info">Semana: <%=Semana4%></th>
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
		
	</div>	

	
	
	
