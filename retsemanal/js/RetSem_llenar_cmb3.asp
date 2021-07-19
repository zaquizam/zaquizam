<!--#include file="conexionRS.asp" -->
<%
'
' RetSem_llenar_cmb3.asp - 15jul21 - 
'
' Cambio en combo Marca - 
'
Session.lcid = 1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'
if conexionRS.errors.count <> 0 Then
  Response.Write ("No hay conexionRS con la BD...!")
  Response.End
end if

Dim opcion, QrySql, idCat, idCliente, idFab
'
'opcion  = Cint(Request.Form("opcion"))
'idQuery = Cint(Request.Form("id"))
'
opcion = Request.Form("opcion")
idCat  = Request.Form("idCat")
idFab  = Request.Form("idFab")
idCliente = Request.Form("idCli")
'
IF (Cint(opcion) = 5) THEN
	'
	'ReFill combo Marca
	'			
	Dim rsMarca, arrMarca
	'
	' Buscar Datos de todas las Marcas
	'
	QrySql = vbnullstring	
	QrySql = QrySql & " SELECT DISTINCT"
	QrySql = QrySql & " Id_Marca as id, "
	QrySql = QrySql & " Marca as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria = " & idCat
	QrySql = QrySql & " AND"
	QrySql = QrySql & " Id_Fabricante = " & idFab
	QrySql = QrySql & " AND"
	QrySql = QrySql & " Id_Marca <> 0"
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Marca"	
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsMarca = Server.CreateObject("ADODB.recordset")
	rsMarca.Open QrySql, conexionRS
	'
	if not rsMarca.EOF then
		arrMarca = rsMarca.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsMarca.Close : Set rsMarca = Nothing
	'	
	'Crear Archivo Array Json
	'
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
	conexionRS.close : set conexionRS = nothing
	'		
ELSEIF (Cint(opcion) = 6) THEN
	'
	'ReFill combo Segmento
	'			
	Dim rsSegmento, arrSegmento		
	'
	' Buscar Datos de todas las Segmento
	'
	QrySql = vbnullstring	
	QrySql = QrySql & " SELECT DISTINCT"
	QrySql = QrySql & " Id_Segmento as id, "
	QrySql = QrySql & " Segmento as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria = " & idCat
	QrySql = QrySql & " AND"
	QrySql = QrySql & " Id_Fabricante = " & idFab
	QrySql = QrySql & " AND"
	QrySql = QrySql & " Id_Segmento <> 0"
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Segmento"	
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsSegmento = Server.CreateObject("ADODB.recordset")
	rsSegmento.Open QrySql, conexionRS
	'
	if not rsSegmento.EOF then
		arrSegmento = rsSegmento.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsSegmento.Close : Set rsSegmento = Nothing
	'	
	'Crear Archivo Array Json
	'
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
	conexionRS.close : set conexionRS = nothing
	'	
ELSEIF (Cint(opcion) = 7) THEN
	'
	'ReFill combo Tamaño
	'			
	Dim rsTamano, arrTamano		
	'
	' Buscar Datos de todas las Tamano
	'
	QrySql = vbnullstring	
	QrySql = QrySql & " SELECT DISTINCT"
	QrySql = QrySql & " Id_Tamano as id, "
	QrySql = QrySql & " Tamano as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria = " & idCat
	QrySql = QrySql & " AND"
	QrySql = QrySql & " Id_Fabricante = " & idFab
	QrySql = QrySql & " AND"
	QrySql = QrySql & " Id_Tamano <> 0"
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Tamano"	
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsTamano = Server.CreateObject("ADODB.recordset")
	rsTamano.Open QrySql, conexionRS
	'
	if not rsTamano.EOF then
		arrTamano = rsTamano.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsTamano.Close : Set rsTamano = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrTamano) then

		For i = 0 to ubound(arrTamano, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrTamano(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrTamano(1,i) & chr(34) & chr(125) &chr(44)
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
	conexionRS.close : set conexionRS = nothing
	'	
ELSEIF (Cint(opcion) = 8) THEN
	'
	'Fill combo Productos
	'			
	Dim rsProducto, arrProducto
	'
	' Buscar Datos de todas las Productos
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT DISTINCT"
	QrySql = QrySql & " CodigoBarra as id, "
	QrySql = QrySql & " Descripcion as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE  "
	QrySql = QrySql & " Id_Categoria= " &  idCat
	QrySql = QrySql & " AND id_Fabricante = " & idFab
	QrySql = QrySql & " AND CodigoBarra IS NOT NULL"
	QrySql = QrySql & " AND CodigoBarra <> ''"
	QrySql = QrySql & " AND Descripcion IS NOT NULL"
	QrySql = QrySql & " AND Descripcion <> ''"
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " Descripcion"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsProducto = Server.CreateObject("ADODB.recordset")
	rsProducto.Open QrySql, conexionRS
	'
	if not rsProducto.EOF then
		arrProducto = rsProducto.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsProducto.Close : Set rsProducto = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrProducto) then

		For i = 0 to ubound(arrProducto, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrProducto(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrProducto(1,i) & chr(34) & chr(125) &chr(44)
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
	conexionRS.close : set conexionRS = nothing
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