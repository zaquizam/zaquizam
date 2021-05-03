<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValUpdateDetallesxProductosxUnicoValDatos.asp - 07abr21 -
	'
	Dim updSql	
	'
	Session.lcid = 2057
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
    ' Capturar las variables
    '
	idConsDetalle 	= Request.QueryString("idConsumoDetalle")
	precio			= Request.QueryString("precio")
	cantidad		= Request.QueryString("cantidad")
	barcode			= Request.QueryString("barcode")	
	idMoneda		= Request.QueryString("idMoneda")
	moneda			= Request.QueryString("moneda")
	tasaCambio		= Request.QueryString("tasa")
	TotalCompra     = Request.QueryString("total")
    '
    ' Actualizar Datos Validando....
    '
    updSql = vbnullstring
	updSql = updSql & " UPDATE PH_Consumo_Detalle_Productos "
    updSql = updSql & " SET"
    updSql = updSql & " Precio_producto=" & precio & ","
    updSql = updSql & " Cantidad=" & cantidad & ","
    updSql = updSql & " Numero_codigo_barras= '" & barcode & "',"    
	updSql = updSql & " Moneda= '"  & moneda & "',"
	updSql = updSql & " id_moneda=" & idMoneda & ","
	updSql = updSql & " Total_Compra=" & TotalCompra & ","
	updSql = updSql & " Tasa_de_Cambio=" & tasaCambio & ","
    updSql = updSql & " Validado='1'"
	'
    updSql = updSql & " WHERE"
    updSql = updSql & " Id_Consumo_detalle_Productos =" & idConsDetalle
    '
    'Response.Write updSql
	'Response.end
    '
    Set objExec = conexion.Execute(updSql)
    Set objExec = Nothing
    '
    If Err.Number = 0 Then
		'
		' Buscar el id del consumo
		'		
		QrySql = vbnullstring
		QrySql = " SELECT id_consumo FROM PH_Consumo_Detalle_Productos WHERE Id_Consumo_detalle_Productos = " & idConsDetalle
		Set rsSql = Server.CreateObject("ADODB.recordset")
		'
		rsSql.Open QrySql, conexion
		'
		if not (rsSql.EOF and rsSql.BOF) then
		   idConsumo = rsSql(0)
		end if		
		'
		rsSql.close : Set rsSql = Nothing	
		'
		Set rsscroll = CreateObject("ADODB.Recordset")
		Dim strSQL, rsscroll, intRow
		strSQL = "SELECT COUNT(Cantidad) AS TOTAL FROM PH_Consumo_Detalle_Productos WHERE Validado='0' AND Id_Consumo =" & idConsumo
		rsscroll.open strSQL, conexion
		intRow = rsscroll("Total")
		rsscroll.close: set rsscroll = nothing 
		'	
		if CInt(intRow) = 0 then			
			'
			updSql = vbnullstring
			updSql = updSql & " UPDATE PH_Consumo"
			updSql = updSql & " SET"
			updSql = updSql & " Validado='1'"		'
			updSql = updSql & " WHERE"
			updSql = updSql & " Id_Consumo =" & idConsumo
			'		
			Set objExec = conexion.Execute(updSql)
			Set objExec = Nothing
			'
		end if	
		'
		If Err.Number = 0 Then
			Response.Write True
		Else
			Response.write "error" '(Err.Description)
		End If	
		'		
    Else
        Response.write "error" '(Err.Description)
    End If
	'
	conexion.Close
    Set objExec = Nothing	
    Set conexion = Nothing
	'
%>