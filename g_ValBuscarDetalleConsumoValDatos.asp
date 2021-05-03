<%@language=vbscript%>
<!--#include file="conexion.asp"-->
<%
	' g_ValBuscarDetalleConsumoValDatos.asp - 07abr21
	'
	Session.lcid = 2057
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'
	Dim idConsumo, strSQL, rsDetalleProducto, arrDetProd
	'
	idConsumo = Request.QueryString("idConsumo")
	'
	Set rsDetalleProducto = CreateObject("ADODB.Recordset")
	rsDetalleProducto.CursorType = adOpenKeyset 
	rsDetalleProducto.LockType = 2 'adLockOptimistic 	
	'
	strSQL = vbnullString
	strSQL = strSQL & " SELECT"
	strSQL = strSQL & " cantidad,"
	strSQL = strSQL & " Precio_producto,"
	strSQL = strSQL & " Tasa_de_cambio,"
	strSQL = strSQL & " Total_compra, "
	strSQL = strSQL & " Moneda, "
	strSQL = strSQL & " Numero_codigo_barras"
	strSQL = strSQL & " FROM"
	strSQL = strSQL & " PH_Consumo_Detalle_Productos"
	strSQL = strSQL & " WHERE"
	strSQL = strSQL & " Id_Consumo_Detalle_Productos = " & idConsumo
	'
	rsDetalleProducto.open strSQL, conexion	
	'
	'Response.write strSQL 
	'Response.End
	'
	Response.ContentType = "application/json"	
	'
	If not rsDetalleProducto.EOF  Then
    	arrDetProd = rsDetalleProducto.GetRows()  ' Convert recordset to 2D Array
	end if
	'	
	sTabla=vbnullstring
	
	if IsArray(arrDetProd) then
	
		For i = 0 to ubound(arrDetProd, 2)
		
							sTabla  =   chr(123) &  chr(34) & "cantidad"  & chr(34) & ":" & chr(34) & (arrDetProd(0,i))  & chr(34) & chr(44)
			'
			sTabla    =   	sTabla  &   chr(34)  & "precio" & chr(34)   & ":" & chr(34) & (arrDetProd(1,i)) & chr(34) & chr(44)
			
			sTabla    =   	sTabla  &   chr(34)  & "tasa"   & chr(34)   & ":" & chr(34) & (arrDetProd(2,i)) & chr(34) & chr(44)
			
			sTabla    =   	sTabla  &   chr(34)  & "total"   & chr(34)   & ":" & chr(34) & (arrDetProd(3,i)) & chr(34) & chr(44)
			
			sTabla    =   	sTabla  &   chr(34)  & "moneda"   & chr(34)   & ":" & chr(34) & (arrDetProd(4,i)) & chr(34) & chr(44)
			
			sTabla    =   	sTabla  &   chr(34)  & "barcode"  & chr(34) & ":" & chr(34) & (arrDetProd(5,i)) & chr(34) & chr(125) & chr(44)
			
			sTablaJson = sTablaJson & sTabla
			sTabla=vbnullstring
			
		next				
		'
	else
		'Eof()
		'sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		'
						sTabla  =   chr(123) &  chr(34) & "motivo"     & chr(34) & ":" & chr(34) & "No Aplica" & chr(34) & chr(44)
		'
		sTabla    =   	sTabla  &   chr(34)  & "comentario" & chr(34)  & ":" & chr(34) & "No Aplica" & chr(34) & chr(44)
		
		sTabla    =   	sTabla  &   chr(34)  & "respuesta"  & chr(34)  & ":" & chr(34) & "No Aplica" & chr(34) & chr(125) & chr(44)
		'				
		sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		
		'JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	end if
	'	
	sTabla   = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData = chr(91) & sTabla & chr(93) 
	'JsonData = sTabla 
	Response.Write(JsonData)
	'
	rsDetalleProducto.close : set rsDetalleProducto = Nothing 
	conexion.Close : Set conexion = Nothing
	'
%>