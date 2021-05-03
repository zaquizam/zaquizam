<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_ValUpdateDetallesxProductosxUnico.asp // 03ene21 - 13ene21
	'
	Dim updSql	
	'
	Session.lcid = 2057
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
	'
	' Dim dd, mm, yy, hh, nn, ss, datevalue, dtsnow
	' '
	' dtsnow = Now()
	' dd = Right("00" & Day(dtsnow), 2)
	' mm = Right("00" & Month(dtsnow), 2)
	' yy = Year(dtsnow)
	' hh = Right("00" & Hour(dtsnow), 2)
	' nn = Right("00" & Minute(dtsnow), 2)
	' ss = Right("00" & Second(dtsnow), 2)
	' datevalue = yy & "-" & mm & "-" & dd
	' timevalue = hh & ":" & nn & ":" & ss
	' sUpdate = datevalue & " " & timevalue
    '
    ' Capturar las variables
    '
	idConsDetalle 	= Request.QueryString("idConsumoDetalle")
	precio			= Request.QueryString("precio")
	cantidad		= Request.QueryString("cantidad")
	barcode			= Request.QueryString("barcode")	
	idConsumo		= Request.QueryString("idConsumo")
	idMoneda		= Request.QueryString("idMoneda")
	moneda			= Request.QueryString("moneda")
	tasaCambio		= Request.QueryString("tasa")
	TotalCompra     = Request.QueryString("total")
    '
    ' Actualizar Datos Validando....
    '
    updSql = vbnullstring
	'updSql = updSql & " SET DATEFORMAT MDY"
	updSql = updSql & " UPDATE PH_Consumo_Detalle_Productos "
    updSql = updSql & " SET"
    updSql = updSql & " Precio_producto=" & Precio & ","
    updSql = updSql & " Cantidad=" & cantidad & ","
    updSql = updSql & " Numero_codigo_barras= '" & barcode & "',"    
	updSql = updSql & " Moneda= '" & moneda & "',"
	updSql = updSql & " id_moneda=" & idMoneda & ","
	updSql = updSql & " Total_Compra=" & TotalCompra & ","
	updSql = updSql & " Tasa_de_Cambio=" & tasaCambio & ","
    'updSql = updSql & " Fec_Inactivo='" & sUpdate  & "',"
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
		' Actualizar Datos Validando....
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
		' '
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