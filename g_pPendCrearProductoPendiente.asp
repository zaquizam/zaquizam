<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
''
' g_pPendCrearProductosPendiente.asp  - 11mar21 - 22mar21
''
Dim instSql, sNuevoReg, Sql
Dim dd, mm, yy, hh, nn, ss
Dim datevalue, timevalue, dtsnow
'
Session.lcid      = 1034
Response.CodePage = 65001
Response.CharSet  = "utf-8"
'
' Capturar las variables
'
barcode 	= Request.QueryString("barcode")
icategoria	= Request.QueryString("categoria")
ifabricante	= Request.QueryString("fabricante")
imarca		= Request.QueryString("marca")
isegmento	= Request.QueryString("segmento")
itamano		= Request.QueryString("tamano")
irango		= Request.QueryString("rango")
iunidad		= Request.QueryString("unidad")
descProducto= Request.QueryString("descProducto")
fragmento   = Request.QueryString("fragmento")
iactivo     = 1
indPendiente= 1
'
sCreacion = mydate(DATE)
sIP       = Request.ServerVariables("REMOTE_ADDR")
'    
' Store DateTimeStamp once. / Fecha Formato Americano
'
dtsnow = Now()
dd = Right("00" & Day(dtsnow), 2)
mm = Right("00" & Month(dtsnow), 2)
yy = Year(dtsnow)
hh = Right("00" & Hour(dtsnow), 2)
nn = Right("00"  & Minute(dtsnow), 2)
ss = Right("00" & Second(dtsnow), 2)
datevalue = yy  & "-" & mm & "-" & dd
timevalue = hh  & ":" & nn & ":" & ss
sUpdate = datevalue & " " & timevalue
'
' Insertar Datos ....
'
instSql = ""
instSql = instSql & " INSERT INTO PH_CB_Producto "
instSql = instSql & " ( codigobarra, id_categoria, id_fabricante, id_marca, id_segmento,"
instSql = instSql & " id_tamano, id_rangotamano, id_unidadmedida, fec_alta,"
instSql = instSql & " producto, fragmentacion,"
instSql = instSql & " IP, ind_activo, ind_pendiente, idsession, Fec_Ult_Mod )"
'
instSql = instSql & " VALUES "
'
instSql = instSql & "('" & barcode & "',"
instSql = instSql & ""  & icategoria & ","
instSql = instSql & ""  & ifabricante & ","
instSql = instSql & ""  & imarca & ","
instSql = instSql & ""  & isegmento & ","
instSql = instSql & ""  & itamano & ","
instSql = instSql & ""  & irango & ","
instSql = instSql & ""  & iunidad & ","
instSql = instSql & "'" & sCreacion & "',"
instSql = instSql & "'" & descProducto & "',"
instSql = instSql & "'" & fragmento & "',"
'
instSql = instSql & "'" & sIp & "',"
instSql = instSql & "'" & iactivo & "',"
instSql = instSql & "'" & indPendiente & "',"
instSql = instSql & ""  & Session.SessionID & ","
instSql = instSql & "'" & sUpdate & "')"
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

Function myDate(dt)
    ' Formatear Fecha consumo
    Dim d, m, y, sep
    sep = "-"
    ' right(..) here works as rpad(x,2,"0")
    d = right("0" & datePart("d", dt), 2)
    m = right("0" & datePart("m", dt), 2)
    y = datePart("yyyy", dt)
    myDate = y & sep & m & sep & d
End Function

%>