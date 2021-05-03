<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
'
' g_ValAgregarProductosalConsumo.asp // 08-ene-21 - 13ene21
'
Dim dd, mm, yy, hh, nn, ss, instSql 
Dim datevalue, timevalue, dtsnow
'
Session.lcid = 1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'				
'
'Capturar las variables
'
idConsumo    = Request.Form("idConsumo")
idHogar      = Request.Form("idHogar")
tieneCodigo  = 1
idCategoria  = 0
NroBarcode   = Request.Form("bArcode")
Cantidad     = Request.Form("cAntidad")
Precio       = Request.Form("pRecio")
idMoneda	 = Request.Form("idMoneda")
Moneda	 	 = Request.Form("mOneda")
tasaCambio 	 = Request.Form("tasaCambio")
'
tipoCodigo	=	"Manual"
sNuevoReg 	=	"G"
bValidado  	=	1
bPendiente	= 	0
bResuelto	=	0
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
'
' Formatear precio
' Precio = Replace(Precio, ".", "")
' Precio = Replace(Precio, ",", ".")
' '
' tasaCambio = Replace(tasaCambio, ".", "")
' tasaCambio = Replace(tasaCambio, ",", ".")
'
TotalCompra = 0	
TotalCompra = ( Precio * tasaCambio ) * Cantidad 	
'	
' Variables General
'
sIP      = Request.ServerVariables("REMOTE_ADDR")	
'
' Insertar Datos ....
'
instSql = vbnullstring
instSql = instSql & " INSERT INTO PH_Consumo_Detalle_Productos "
instSql = instSql & " ("
instSql = instSql & " id_Consumo, id_Hogar, Tiene_Codigo_Barras, Numero_codigo_barras,"
instSql = instSql & " Id_Categoria, Cantidad, Precio_producto, id_moneda, moneda,"
instSql = instSql & " total_compra, tasa_de_cambio, Tipo_codigo_barras,"
instSql = instSql & " IP, idsession, validado, Pendiente, Resuelto, Fecha_Creacion, Fec_ult_mod, Status_registro"
instSql = instSql & " )"
'
instSql = instSql & " VALUES "
'
instSql = instSql & "("  & idConsumo & ","
instSql = instSql & ""   & idHogar & ","
instSql = instSql & ""   & tieneCodigo & ","
instSql = instSql & "'"  & NroBarcode & "',"
instSql = instSql & ""   & idCategoria & ","
instSql = instSql & ""   & Cantidad & ","
instSql = instSql & ""   & (Precio) & ","
'
instSql = instSql & ""   & idMoneda & ","
instSql = instSql & "'"  & Moneda & "',"
instSql = instSql & ""   & (TotalCompra) & ","
instSql = instSql & ""   & (tasaCambio) & ","
instSql = instSql & "'"  & (tipoCodigo) & "',"
'	
instSql = instSql & "'"  & sIp & "',"
instSql = instSql & ""   & Session.SessionID & ","
instSql = instSql & ""   & bValidado & ","
instSql = instSql & ""   & bPendiente & ","
instSql = instSql & ""   & bResuelto & ","
instSql = instSql & "'"  & datevalue & "',"
instSql = instSql & "'"  & sUpdate & "',"
'
instSql = instSql & "'"  & sNuevoReg & "')"
'
'Response.Write  instSql
'Response.End
'
Set objExec = conexion.Execute(instSql)
'
If Err.Number = 0 Then
	Response.write True
Else
	Response.write (Err.Description)
End If
'
conexion.Close
Set objExec = Nothing
Set conexion = Nothing
'
%>
