<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_pPendValidarProductoPendienteExista.asp - 15mar21
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
	CodigoBarras		= Request.QueryString("barcode")
    '
    ' Buscar Si ek Producto Existe
    '
	Dim rsBuscarProducto
	set rsBuscarProducto = CreateObject("ADODB.Recordset")
	rsBuscarProducto.CursorType = adOpenKeyset 
	rsBuscarProducto.LockType = 2 'adLockOptimistic 
	'	
	sql = vbnullString
	sql = sql & " SELECT"
	sql = sql & " id_Producto"
	sql = sql & " FROM"
	sql = sql & " PH_CB_Producto"
	sql = sql & " WHERE"
	sql = sql & " CodigoBarra = '" & CodigoBarras & "'"	
	'
	'response.write sql
	'response.end
	'	
    rsBuscarProducto.Open sql ,conexion
	'
	if not rsBuscarProducto.EOF then
		Response.Write True
	else	
		Response.Write False
	end if	
	'
	rsBuscarProducto.close : Set rsBuscarProducto = Nothing
	conexion.Close  : Set conexion = Nothing
	'
%>