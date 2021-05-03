<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_pPendBuscarDetallexProductos.asp - 10mar21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
	idConDetalle		= Request.QueryString("idConDetalle")
    '
    ' Buscar Detalle Regsitro
    '
	Dim rsDetalleRegistro
	set rsDetalleRegistro = CreateObject("ADODB.Recordset")
	rsDetalleRegistro.CursorType = adOpenKeyset 
	rsDetalleRegistro.LockType = 2 'adLockOptimistic 
	'	
	sql = vbnullString
	sql = sql & " SELECT"
	sql = sql & " numero_codigo_barras,"
	sql = sql & " cantidad,"
	sql = sql & " Precio_producto,"
	sql = sql & " tasa_de_cambio,"
	sql = sql & " total_compra,"
	sql = sql & " moneda,"
	sql = sql & " FORMAT (fecha_creacion, 'dd/MM/yyyy ') AS fecha"
	sql = sql & " FROM"
	sql = sql & " PH_Consumo_Detalle_Productos"
	sql = sql & " WHERE"
	sql = sql & " Id_Consumo_Detalle_Productos = "& idConDetalle
	'
	'response.write sql
	'response.end
	'
    rsDetalleRegistro.Open sql ,conexion
	'
	if not rsDetalleRegistro.EOF then
		arrDetalleRegistro = rsDetalleRegistro.GetRows()  ' Convert recordset to 2D Array
	end if
	'			
	Response.ContentType = "application/json"		
	'
	sTabla=vbnullstring
	
	if IsArray(arrDetalleRegistro) then
	
		For i = 0 to ubound(arrDetalleRegistro, 2)
		
			sTabla    =    chr(123) &  chr(34) & "codigobar" & chr(34) & ":" & chr(34) & arrDetalleRegistro(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla   &  chr(34) & "cantidad"	 & chr(34) & ":" & chr(34) & arrDetalleRegistro(1,i) & chr(34) & chr(44)
			sTabla    =    sTabla   &  chr(34) & "precio"	 & chr(34) & ":" & chr(34) & arrDetalleRegistro(2,i) & chr(34) & chr(44)
			sTabla    =    sTabla   &  chr(34) & "tasa"	     & chr(34) & ":" & chr(34) & arrDetalleRegistro(3,i) & chr(34) & chr(44)
			sTabla    =    sTabla   &  chr(34) & "total"	 & chr(34) & ":" & chr(34) & arrDetalleRegistro(4,i) & chr(34) & chr(44)
			sTabla    =    sTabla   &  chr(34) & "moneda"	 & chr(34) & ":" & chr(34) & arrDetalleRegistro(5,i) & chr(34) & chr(44)
			'
			sTabla    =    sTabla  &  chr(34) & "fecha"     & chr(34) & ":" & chr(34) & arrDetalleRegistro(6,i) & chr(34) & chr(125) & chr(44)
			
			sTablaJson = sTablaJson & sTabla
			sTabla=vbnullstring
			
		next				
		'
	else
		'Eof()
		'sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		'
		sTabla    =    chr(123) &  chr(34) & "codigobar" & chr(34) & ":" & chr(34) & "NO APLICA" & chr(34) & chr(44)
		sTabla    =    sTabla   &  chr(34) & "cantidad"	 & chr(34) & ":" & chr(34) & "NO APLICA" & chr(34) & chr(44)
		sTabla    =    sTabla   &  chr(34) & "precio"	 & chr(34) & ":" & chr(34) & "NO APLICA" & chr(34) & chr(44)
		sTabla    =    sTabla   &  chr(34) & "tasa"	     & chr(34) & ":" & chr(34) & "NO APLICA" & chr(34) & chr(44)
		sTabla    =    sTabla   &  chr(34) & "total"	 & chr(34) & ":" & chr(34) & "NO APLICA" & chr(34) & chr(44)
		sTabla    =    sTabla   &  chr(34) & "moneda"	 & chr(34) & ":" & chr(34) & "NO APLICA" & chr(34) & chr(44)
		'
		sTabla    =    sTabla  &  chr(34) & "fecha"      & chr(34) & ":" & chr(34) & "NO APLICA" & chr(34) & chr(125) & chr(44)
		'		
		sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		'
	end if
	'	
	sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData= chr(91) & sTabla & chr(93) '& chr(125)
	Response.Write(JsonData)
	'
	conexion.Close    
    Set conexion = Nothing
	'
%>