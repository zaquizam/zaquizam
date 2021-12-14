<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValBuscarDetallesxProductosxUnico.asp
	' 02ene21
	'
	Session.lcid=1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'	
	Dim idConsumo, rsDetalleUnico, arrDetalleUnico
	'
	idConsumoDetalle	=	Request.QueryString("id_ConsumoDetalle")		
	'	
	' Buscar Resultados
	'
	set rsDetalleUnico			=	CreateObject("ADODB.Recordset")
	rsDetalleUnico.CursorType	=	adOpenKeyset 
	rsDetalleUnico.LockType		=	2 'adLockOptimistic 
	'		
	sql = vbnullstring	
	sql = sql & " SELECT"
	sql = sql & " PH_Consumo_Detalle_Productos.Numero_codigo_barras,"
	sql = sql & " PH_Consumo_Detalle_Productos.Cantidad,"
	sql = sql & " PH_Consumo_Detalle_Productos.Precio_producto,"
	sql = sql & " PH_Consumo_Detalle_Productos.Tipo_codigo_barras,"	
	sql = sql & " PH_Consumo_Detalle_Productos.Tasa_de_cambio,"
	sql = sql & " PH_Consumo_Detalle_Productos.Moneda,"
	sql = sql & " PH_Consumo_Detalle_Productos.Total_compra,"	
	sql = sql & " PH_Consumo_Detalle_Productos.Id_Consumo_detalle_Productos,"
	sql = sql & " PH_Consumo_Detalle_Productos.id_Moneda,"
	sql = sql & " PH_Consumo_Detalle_Productos.id_Categoria,"
	sql = sql & " PH_Consumo_Detalle_Productos.Unidad_empaque"
	sql = sql & " FROM"
	sql = sql & " PH_Consumo_Detalle_Productos"
	sql = sql & " WHERE"
	sql = sql & " PH_Consumo_Detalle_Productos.Id_Consumo_detalle_Productos = " & idConsumoDetalle
	'
	'Response.Write sql
	'Response.End
	'
    rsDetalleUnico.Open sql, conexion
	'
	if not rsDetalleUnico.eof then
		arrDetalleUnico = rsDetalleUnico.GetRows()  ' Convert recordset to 2D Array					
	end if
		'
	rsDetalleUnico.Close
	Set rsDetalleUnico = Nothing
	'
	'Response.ContentType = "application/json"		
	'
	sTabla=vbnullstring
	
	if IsArray(arrDetalleUnico) then
	
		For i = 0 to ubound(arrDetalleUnico, 2)
			sTabla    =   chr(123)&  chr(34) & "barcode"	& chr(34) & ":" & chr(34) & arrDetalleUnico(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "cantidad" 	& chr(34) & ":" & chr(34) & arrDetalleUnico(1,i) & chr(34) & chr(44)
			tasa = replace(arrDetalleUnico(4,i),",",".")
			sTabla    =    sTabla &  chr(34) & "tasa" 	    & chr(34) & ":" & chr(34) & tasa & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "moneda"	    & chr(34) & ":" & chr(34) & arrDetalleUnico(5,i) & chr(34) & chr(44)			
			sTabla    =    sTabla &  chr(34) & "idmoneda"	& chr(34) & ":" & chr(34) & Cstr(arrDetalleUnico(8,i)) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "idcateg"	& chr(34) & ":" & chr(34) & Cstr(arrDetalleUnico(9,i)) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "unidademp"	& chr(34) & ":" & chr(34) & arrDetalleUnico(10,i) & chr(34) & chr(44)
			'			
			precio = replace(arrDetalleUnico(2,i),",",".")
			sTabla    =    sTabla &  chr(34) & "precio"    & chr(34) & ":" & chr(34) & precio & chr(34) & chr(125) & chr(44)
			'
			sTablaJson = sTablaJson & sTabla
			sTabla=vbnullstring
		next
		'
		sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
		JsonData= chr(91) & sTabla & chr(93) '& chr(125)
		'
	else
		'Eof()
		sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	end if
	Response.Write(JsonData)
	conexion.close
	set conexion = nothing	
	
%>