<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_pPendbuscarDatosCodigoBarras.asp - 11mar21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
	CodigoBarras		= Request.QueryString("CodigoBarras")
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
	sql = sql & " PH_CB_Categoria.Categoria,"
	sql = sql & " PH_CB_Producto.Producto"
	sql = sql & " FROM"
	sql = sql & " PH_CB_Categoria"
	sql = sql & " INNER JOIN PH_CB_Producto ON PH_CB_Producto.Id_Categoria = PH_CB_Categoria.id_Categoria"
	sql = sql & " WHERE"
	sql = sql & " PH_CB_Producto.CodigoBarra = '" & CodigoBarras & "'"	
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
		
			sTabla    =    chr(123) &  chr(34) & "categoria" & chr(34) & ":" & chr(34) & arrDetalleRegistro(0,i) & chr(34) & chr(44)			
			'
			sTabla    =    sTabla  &  chr(34)  & "descripcion"  & chr(34) & ":" & chr(34) & arrDetalleRegistro(1,i) & chr(34) & chr(125) & chr(44)
			
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
			
		next				
		'
	else
		'Eof()
		'sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		'
		sTabla    =    chr(123) &  chr(34) & "categoria" & chr(34) & ":" & chr(34) & "NO APLICA" & chr(34) & chr(44)		
		'
		sTabla    =    sTabla  &  chr(34) & "descripcion"   & chr(34) & ":" & chr(34) & "NO APLICA" & chr(34) & chr(125) & chr(44)
		'		
		sTablaJson = sTablaJson & sTabla
		sTabla = vbnullstring
		'
	end if
	'	
	sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData= chr(91) & sTabla & chr(93) '& chr(125)
	Response.Write(JsonData)
	'
	conexion.Close  : Set conexion = Nothing
	'
%>