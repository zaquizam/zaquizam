<!--#include file="Conexion.asp" -->
<%
'
' llenar_cmb_convivencias_x_fab.asp - 16jun21 - 
'
Session.lcid = 1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'
if conexion.errors.count <> 0 Then
  Response.Write ("No hay Conexion con la BD...!")
  Response.End
end if

Dim opcion, id, QrySql, idCat, idFab
'
'opcion  = Cint(Request.Querystring("opcion"))
'idQuery = Cint(Request.Querystring("id"))
'
opcion  = Request.Form("opcion")
idCat = Request.Form("idCat")
idFab = Request.Form("idFab")
'
'Response.write(opcion) & "<br>"
'Response.write(id)  & "<br>"
'
IF (Cint(opcion) = 1) THEN
	'
	'Fill combo Marca por Categoria + Fabricante A y B
	'	
	Dim rsMarca, arrMarca
	'
	' Buscar Datos de todos los Marcas segun la categoria y Fabricante
	'
	QrySql = vbnullstring
	QrySql = " SELECT DISTINCT Id_Marca AS id, Marca AS nombre FROM PH_DataCrudaMensual WHERE " & _
	" id_Categoria  = " & idCat & _
	" AND id_Fabricante = " & idFab & _
	" ORDER BY Marca ASC"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsMarca = Server.CreateObject("ADODB.recordset")
	rsMarca.Open QrySql, conexion
	'
	if not rsMarca.EOF then
		arrMarca = rsMarca.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsMarca.Close : Set rsMarca = Nothing
	'
	'Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
	sTabla = vbnullstring

	if IsArray(arrMarca) then

		For i = 0 to ubound(arrMarca, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrMarca(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrMarca(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbnullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexion.close : set conexion = nothing
	'	
ELSEIF (Cint(opcion) = 2) THEN
	'
	'Fill combo Segmento A y B
	'
	Dim rsSegmento, arrSegmento
	'
	' Buscar Datos de todos los Segmentos segun la categoria
	'
	QrySql = vbnullstring	
	QrySql = " SELECT DISTINCT Id_segmento AS id, segmento AS nombre FROM PH_DataCrudaMensual WHERE " & _
	" id_Categoria  = " & idCat & _
	" AND id_Fabricante = " & idFab & _
	" ORDER BY segmento ASC"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsSegmento = Server.CreateObject("ADODB.recordset")
	rsSegmento.Open QrySql, conexion
	'
	if not rsSegmento.EOF then
		arrSegmento = rsSegmento.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsSegmento.Close : Set rsSegmento = Nothing
	'
	'Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
	sTabla = vbnullstring

	if IsArray(arrSegmento) then

		For i = 0 to ubound(arrSegmento, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrSegmento(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrSegmento(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbnullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexion.close : set conexion = nothing
	'
ELSEIF (Cint(opcion) = 3) THEN 	
	'
	'Fill combo Rango Tamaño A y B
	'
	Dim rsRangTamano, arrRangTamano
	'
	' Buscar Datos de todos los RangTamanos segun la categoria
	'		
	QrySql = vbnullstring	
	QrySql = " SELECT DISTINCT Id_rangotamano AS id, rangotamano AS nombre FROM PH_DataCrudaMensual WHERE " & _
	" id_Categoria  = " & idCat & _
	" AND id_Fabricante = " & idFab & _
	" ORDER BY rangotamano ASC"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsRangTamano = Server.CreateObject("ADODB.recordset")
	rsRangTamano.Open QrySql, conexion
	'
	if not rsRangTamano.EOF then
		arrRangTamano = rsRangTamano.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsRangTamano.Close : Set rsRangTamano = Nothing
	'
	'Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
	sTabla = vbnullstring

	if IsArray(arrRangTamano) then

		For i = 0 to ubound(arrRangTamano, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrRangTamano(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrRangTamano(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		 & chr(34)& ":" & chr(34)  & "0" 		 & chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbnullstring

	end if
	''
	sTabla  = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexion.close : set conexion = nothing
	'		
ELSE
	' de lo Contrario
	Response.write "error"
END IF
'

FUNCTION RemoverSaltodeLinea(byval str)
	IF isNull(str) THEN str = "" END IF
	str = REPLACE(str,vbCr,"")			'Chr(13)
	str = REPLACE(str,vbLf,"")			'Chr(10)
	str = REPLACE(str,VbCrlf,"")		'Chr(13)+Chr(10)
	str = REPLACE(str,vbNewLine,"")		'vbNewLine
	str = REPLACE(str,vbFormFeed,"")	'Chr(12)
	str = REPLACE(str,vbTab,"")			'Chr(9)
	str = REPLACE(str,vbTab,"")			'Chr(11)
	''
	RemoverSaltodeLinea = TRIM(str)

END FUNCTION

%>