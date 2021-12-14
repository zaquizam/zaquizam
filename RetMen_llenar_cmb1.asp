<!--#include file="conexionRS.asp" -->
<%
'
'RetMen_llenar_cmb1.asp - 12jul21 - 13dic21
'
Session.lcid = 1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'
if conexionRS.errors.count <> 0 Then
  Response.Write ("No hay conexionRS con la BD...!")
  Response.End
end if

Dim opcion, QrySql, idCat, idCliente
'
'opcion  = Cint(Request.Querystring("opcion"))
'idQuery = Cint(Request.Querystring("id"))
'
opcion = Request.Querystring("opcion")
idCat  = Request.Querystring("idCat")
idCliente = Request.Querystring("idCli")
'

IF (Cint(opcion) = 1) THEN
	'
	'Fill combo Categoria
	'				
	Dim rsCategoria, arrCategoria
	'
	' Buscar Datos de todas las Categorias
	' 03nov21
	'
	IF (Cint(idCliente) = 1 ) THEN
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT RS_DataProcSem.Id_Categoria AS id,  RS_DataProcSem.Categoria AS nombre FROM dbo.RS_DataProcSem"
		QrySql = QrySql & " GROUP BY RS_DataProcSem.Id_Categoria, RS_DataProcSem.Categoria ORDER BY RS_DataProcSem.Categoria ASC"
	ELSE
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT RS_DataProcSem.Id_Categoria AS id,  RS_DataProcSem.Categoria AS nombre FROM dbo.RS_DataProcSem"		
		QrySql = QrySql & " INNER JOIN dbo.ss_ClienteCategoria ON  RS_DataProcSem.Id_Categoria = ss_ClienteCategoria.Id_Categoria"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " ss_ClienteCategoria.Id_Cliente = " & idCliente
		QrySql = QrySql & " and ss_ClienteCategoria.Ind_Mensual = 1 "
		QrySql = QrySql & " GROUP BY RS_DataProcSem.Id_Categoria, RS_DataProcSem.Categoria ORDER BY RS_DataProcSem.Categoria ASC"
	END IF			
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsCategoria = Server.CreateObject("ADODB.recordSet")
	rsCategoria.Open QrySql, conexionRS
	'
	if not rsCategoria.EOF then
		arrCategoria = rsCategoria.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsCategoria.Close : Set rsCategoria = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrCategoria) then

		For i = 0 to ubound(arrCategoria, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrCategoria(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrCategoria(1,i) & chr(34) & chr(125) &chr(44)
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
	conexionRS.Close : Set conexionRS = Nothing
	'
ELSEIF (Cint(opcion) = 2) THEN
	'
	'Fill combo Area
	'			
	Dim rsArea, arrArea
	'
	' Buscar Datos de todas las Areas
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " Id_Area as id, "
	QrySql = QrySql & " Area as mombre "
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria = " & idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Area, "
	QrySql = QrySql & " Area "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Id_Area "	
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsArea = Server.CreateObject("ADODB.recordSet")
	rsArea.Open QrySql, conexionRS
	'
	if not rsArea.EOF then
		arrArea = rsArea.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsArea.Close : Set rsArea = Nothing
	'
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrArea) then

		For i = 0 to ubound(arrArea, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrArea(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrArea(1,i) & chr(34) & chr(125) &chr(44)
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
	conexionRS.Close : Set conexionRS = Nothing
	'	
ELSEIF (Cint(opcion) = 3) THEN
	'
	'Fill combo Zona
	'			
	Dim rsZona, arrZona
	
	'idCat = Request.Form("idCat")
	'
	' Buscar Datos de todas las Zonas
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " Id_Zona, "
	QrySql = QrySql & " Zona "
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE  "
	QrySql = QrySql & " Id_Categoria= " & idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Zona, "
	QrySql = QrySql & " Zona "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Zona "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsZona = Server.CreateObject("ADODB.recordSet")
	rsZona.Open QrySql, conexionRS
	'
	if not rsZona.EOF then
		arrZona = rsZona.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsZona.Close : Set rsZona = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrZona) then

		For i = 0 to ubound(arrZona, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrZona(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrZona(1,i) & chr(34) & chr(125) &chr(44)
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
	conexionRS.Close : Set conexionRS = Nothing
	'	
	
ELSEIF (Cint(opcion) = 4) THEN
	'
	'Fill combo Canal
	'			
	Dim rsCanal, arrCanal
	
	'idCat = Request.Form("idCat")
	'
	' Buscar Datos de todas las Canales
	'
	QrySql = vbnullstring
	
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " Id_Canal as id, "
	QrySql = QrySql & " rtrim(Canal) as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria = " & idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Canal, "
	QrySql = QrySql & " Canal "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Canal "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsCanal = Server.CreateObject("ADODB.recordSet")
	rsCanal.Open QrySql, conexionRS
	'
	if not rsCanal.EOF then
		arrCanal = rsCanal.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsCanal.Close : Set rsCanal = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrCanal) then

		For i = 0 to ubound(arrCanal, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrCanal(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrCanal(1,i) & chr(34) & chr(125) &chr(44)
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
	conexionRS.Close : Set conexionRS = Nothing
	'		
ELSEIF (Cint(opcion) = 5) THEN
	'
	' Fill combo Fabricante
	'			
	Dim rsFabricante, arrFabricante	
	'
	' Buscar Datos de todas las Fabricantes
	'
	QrySql = vbnullstring	
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " Id_Fabricante as id, "
	QrySql = QrySql & " Fabricante as nombre "
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria = " & idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Fabricante, "
	QrySql = QrySql & " Fabricante "
	QrySql = QrySql & " HAVING "
	QrySql = QrySql & " Id_Fabricante <> 0 "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Fabricante "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsFabricante = Server.CreateObject("ADODB.recordSet")
	rsFabricante.Open QrySql, conexionRS
	'
	if not rsFabricante.EOF then
		arrFabricante = rsFabricante.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsFabricante.Close : Set rsFabricante = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrFabricante) then

		For i = 0 to ubound(arrFabricante, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrFabricante(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrFabricante(1,i) & chr(34) & chr(125) &chr(44)
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
	conexionRS.Close : Set conexionRS = Nothing
	'		
ELSEIF (Cint(opcion) = 6) THEN
	'
	'Fill combo Marca
	'			
	Dim rsMarca, arrMarca
	
	'idCat = Request.Form("idCat")
	'
	' Buscar Datos de todas las Canales
	'
	'QrySql = vbnullstring	
	'QrySql = QrySql & " SELECT "
	'QrySql = QrySql & " Id_Marca as id, "
	'QrySql = QrySql & " Marca as nombre"
	'QrySql = QrySql & " FROM "
	'QrySql = QrySql & " RS_DataProcSem "
	'QrySql = QrySql & " WHERE "
	'QrySql = QrySql & " Id_Categoria = " & idCat
	'QrySql = QrySql & " GROUP BY "
	'QrySql = QrySql & " Id_Marca, "
	'QrySql = QrySql & " Marca "
	'QrySql = QrySql & " HAVING "
	'QrySql = QrySql & " Id_Marca <> 0 "
	'QrySql = QrySql & " ORDER BY "
	'QrySql = QrySql & " Marca "
	'
	if idCat >= 127 and idCat <= 145 then
		QrySql = vbnullstring	
		QrySql = QrySql & " SELECT "
		QrySql = QrySql & " Id_Marca as id, "
		QrySql = QrySql & " Marca+'('+Fabricante+')' as nombre "
		QrySql = QrySql & " FROM "
		QrySql = QrySql & " RS_DataProcSem "
		QrySql = QrySql & " WHERE "
		QrySql = QrySql & " Id_Categoria = " & idCat
		QrySql = QrySql & " GROUP BY "
		QrySql = QrySql & " Id_Marca, "
		QrySql = QrySql & " Marca+'('+Fabricante+')'"
		QrySql = QrySql & " HAVING "
		QrySql = QrySql & " Id_Marca <> 0 "
		QrySql = QrySql & " ORDER BY "
		QrySql = QrySql & " Marca+'('+Fabricante+')'"
	else 
		QrySql = vbnullstring	
		QrySql = QrySql & " SELECT "
		QrySql = QrySql & " Id_Marca as id, "
		QrySql = QrySql & " Marca as nombre"
		QrySql = QrySql & " FROM "
		QrySql = QrySql & " RS_DataProcSem "
		QrySql = QrySql & " WHERE "
		QrySql = QrySql & " Id_Categoria = " & idCat
		QrySql = QrySql & " GROUP BY "
		QrySql = QrySql & " Id_Marca, "
		QrySql = QrySql & " Marca "
		QrySql = QrySql & " HAVING "
		QrySql = QrySql & " Id_Marca <> 0 "
		QrySql = QrySql & " ORDER BY "
		QrySql = QrySql & " Marca "
	end if
	
	
	
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsMarca = Server.CreateObject("ADODB.recordSet")
	rsMarca.Open QrySql, conexionRS
	'
	if not rsMarca.EOF then
		arrMarca = rsMarca.GetRows()  ' Convert recordSet to 2D Array
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
	conexionRS.Close : Set conexionRS = Nothing
	'		
ELSEIF (Cint(opcion) = 7) THEN
	'
	'Fill combo Segmento
	'			
	Dim rsSegmento, arrSegmento
	'
	' Buscar Datos de todas las Canales
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " Id_Segmento as id, "
	QrySql = QrySql & " Segmento as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria = " & idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Segmento, "
	QrySql = QrySql & " Segmento "
	QrySql = QrySql & " HAVING "
	QrySql = QrySql & " Id_Segmento <> 0 "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Segmento "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsSegmento = Server.CreateObject("ADODB.recordSet")
	rsSegmento.Open QrySql, conexionRS
	'
	if not rsSegmento.EOF then
		arrSegmento = rsSegmento.GetRows()  ' Convert recordSet to 2D Array
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
	conexionRS.Close : Set conexionRS = Nothing
	'		
	
ELSEIF (Cint(opcion) = 8) THEN
	'
	'Fill combo Tamaño
	'			
	Dim rsTamano, arrTamano
	'
	' Buscar Datos de todas las Tamanos
	'
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " Id_Tamano as id, "
	QrySql = QrySql & " CONVERT(DECIMAL(10,0),Tamano) as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_DataProcSem "
	QrySql = QrySql & " WHERE "
	QrySql = QrySql & " Id_Categoria =  " & idCat
	QrySql = QrySql & " GROUP BY "
	QrySql = QrySql & " Id_Tamano, "
	QrySql = QrySql & " Tamano "
	QrySql = QrySql & " HAVING "
	QrySql = QrySql & " Id_Tamano <> 0 "
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " CONVERT(DECIMAL(10,0),Tamano) "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsTamano = Server.CreateObject("ADODB.recordSet")
	rsTamano.Open QrySql, conexionRS
	'
	if not rsTamano.EOF then
		arrTamano = rsTamano.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsTamano.Close : Set rsTamano = Nothing
	'
	'Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
	sTabla = vbnullstring

	if IsArray(arrTamano) then

		For i = 0 to ubound(arrTamano, 2)
			'
			'value=Split(arrTamano(1,i),".")			
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
	conexionRS.Close : Set conexionRS = Nothing
	'
ELSEIF (Cint(opcion) = 9) THEN
	'
	'Fill combo Productos
	'			
	Dim rsProducto, arrProducto
	'
	' Buscar Datos de todas las Productos
	'
	QrySql = vbnullstring	
	QrySql = QrySql & " SELECT"
	QrySql = QrySql & " RS_DataProcSem.CodigoBarra as id,"	
	QrySql = QrySql & " TRIM(RS_DataProcSem.Descripcion) as nombre"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " RS_DataProcSem INNER JOIN PH_CB_Fabricante ON RS_DataProcSem.Id_Fabricante = PH_CB_Fabricante.id_Fabricante"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " RS_DataProcSem.Id_Categoria = " & idCat
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_CB_Fabricante.Ind_MarcaPropia = 0"
	QrySql = QrySql & " GROUP BY"
	QrySql = QrySql & " RS_DataProcSem.CodigoBarra,"
	QrySql = QrySql & " RS_DataProcSem.Descripcion"
	QrySql = QrySql & " HAVING"	
	QrySql = QrySql & " ( RS_DataProcSem.CodigoBarra IS NOT NULL AND RS_DataProcSem.CodigoBarra <> '' )"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " ( RS_DataProcSem.Descripcion IS NOT NULL AND RS_DataProcSem.Descripcion <> '' )"	
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " nombre"	
	'	
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsProducto = Server.CreateObject("ADODB.recordSet")
	rsProducto.Open QrySql, conexionRS
	'
	if not rsProducto.EOF then
		arrProducto = rsProducto.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	'Response.end
	rsProducto.Close : Set rsProducto = Nothing
	'	
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrProducto) then

		For i = 0 to ubound(arrProducto, 2)
			'
			sTabla     = chr(123) &  chr(34) & "id" 	& chr(34) & ":" & chr(34) & arrProducto(0,i) & chr(34) & chr(44)
			sTabla     = sTabla   &  chr(34) & "nombre" & chr(34) & ":" & chr(34) & RemoverSaltodeLinea(arrProducto(1,i)) &  " - "  & RemoverSaltodeLinea(arrProducto(0,i)) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34) & ":" & chr(34)  & "0" 		 & chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"  & chr(34) & ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
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
	conexionRS.Close : Set conexionRS = Nothing
	'	
ELSEIF (Cint(opcion) = 10) THEN
	'
	'Fill combo Indicadores
	'			
	Dim rsIndicadores, arrIndicadores
	'
	' Buscar Datos de todas las Indicadores
	'	
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "	
	QrySql = QrySql & " Id_Indicador as id, "
	QrySql = QrySql & " Abreviatura as nombre"
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " RS_Indicadores "
	QrySql = QrySql & " WHERE "	
	if idCliente = 1 then 
		QrySql = QrySql & " Ind_Atenas = 1 " 
	else
		QrySql = QrySql & " Ind_Men = 1 " 
	end if
	'
 	if (idCat > 126 and idCat < 146) or (idCat = 41 or idCat = 18) then
		QrySql = QrySql & " AND ( Id_Indicador <> 3 and Id_Indicador <> 15 and Id_Indicador <> 9 ) "
	end if
	'
	QrySql = QrySql & " AND Ind_Activo = 1 " 
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " Id_Indicador "		
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsIndicadores = Server.CreateObject("ADODB.recordSet")
	rsIndicadores.Open QrySql, conexionRS
	'
	if not rsIndicadores.EOF then
		arrIndicadores = rsIndicadores.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsIndicadores.Close : Set rsIndicadores = Nothing
	'
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrIndicadores) then

		For i = 0 to ubound(arrIndicadores, 2)
			'
			sTabla     =   chr(123)&  chr(34) & "id" 	& chr(34) & ":" & chr(34) & arrIndicadores(0,i) & chr(34) & chr(44)
			sTabla     =   sTabla &  chr(34) & "nombre" & chr(34) & ":" & chr(34) & arrIndicadores(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		 & chr(34) & ":" & chr(34) & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34) & ":" & chr(34) & "No Aplica" & chr(34) & chr(125) & chr(44)
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
	conexionRS.Close : Set conexionRS = Nothing
	'
ELSEIF (Cint(opcion) = 11) THEN
	'
	'Fill combo Semanas
	'			
	Dim rsSemanas, arrSemanas, iSemanaDes, iSemanaHas, rsSemanario
	'	
	if Cint(idCliente=1) then
		'atenas
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT Min(RS_DataProcSem.Id_Semana) AS desde, Max(RS_DataProcSem.Id_Semana) AS hasta, RS_DataProcSem.Id_Categoria FROM RS_DataProcSem"
		QrySql = QrySql & " GROUP BY RS_DataProcSem.Id_Categoria HAVING (((RS_DataProcSem.Id_Categoria)=" & idCat & "));"
		'		
	else
		QrySql = vbnullstring		
		QrySql = QrySql & " SELECT ss_ClienteCategoria.Id_PeriodoDesde AS desde, ss_ClienteCategoria.Id_PeriodoPub AS hasta FROM ss_ClienteCategoria"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " ss_ClienteCategoria.Id_Cliente = " & idCliente
		QrySql = QrySql & " AND"
		QrySql = QrySql & " ss_ClienteCategoria.Ind_Mensual = 1"
		QrySql = QrySql & " AND"
		QrySql = QrySql & " ss_ClienteCategoria.Id_Categoria = " & idCat
	end if
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsSemanario = Server.CreateObject("ADODB.recordSet")
	rsSemanario.Open QrySql, conexionRS
			
	if not (rsSemanario.EOF and rsSemanario.BOF) then		
		iSemanaDes = rsSemanario("desde").value 'rsSemanario(0)
		iSemanaHas = rsSemanario("hasta").value 'rsSemanario(1)
	else
		iSemanaDes = 0
		iSemanaHas = 0
	end if		
	'
	rsSemanario.Close : Set rsSemanario = Nothing
	'
	' Buscar Datos de todas las Semanas
	'	
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT "
	QrySql = QrySql & " IdSemana as id, "
	QrySql = QrySql & " Semana as nombre "
	QrySql = QrySql & " FROM "
	QrySql = QrySql & " ss_Semana "
	if( iSemanaDes <> 0 and iSemanaHas <> 0 ) then
		QrySql = QrySql & " WHERE "
		QrySql = QrySql & " IdSemana >= " & iSemanaDes
		QrySql = QrySql & " And IdSemana <= " & iSemanaHas
	end if	
	QrySql = QrySql & " ORDER BY "
	QrySql = QrySql & " IdSemana DESC "
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsSemanas = Server.CreateObject("ADODB.recordSet")
	rsSemanas.Open QrySql, conexionRS
	'
	if not rsSemanas.EOF then
		arrSemanas = rsSemanas.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsSemanas.Close : Set rsSemanas = Nothing
	'
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrSemanas) then

		For i = 0 to ubound(arrSemanas, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrSemanas(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrSemanas(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34)& ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
		'
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
	conexionRS.Close : Set conexionRS = Nothing
	'		
ELSEIF (Cint(opcion) = 12) THEN	
 	'
	' Verificar cliente contrato el servicio				
	'
	Dim rsCliente, arrCliente
	'
	if (CInt(idCliente) = 1) then
	
		Response.Write CInt(idCliente)
	
	else
		QrySql = vbnullstring		
		QrySql = QrySql & " SELECT COUNT(Id_Cliente) as total FROM dbo.ss_ClienteCategoria"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " dbo.ss_ClienteCategoria.Id_Cliente = " & idCliente
		QrySql = QrySql & " AND"
		QrySql = QrySql & " dbo.ss_ClienteCategoria.Ind_Mensual = 1"
		'
		Response.Write QrySql & "<BR><BR>"
		'Response.end
		'
		Set rsCliente = Server.CreateObject("ADODB.recordSet")
		rsCliente.Open QrySql, conexionRS
		'
		if not (rsCliente.EOF and rsCliente.BOF) then
			Response.Write rsCliente(0)
		else			
			Response.Write 0			
		end if		
		'
		rsCliente.Close : Set rsCliente = Nothing
		'
		conexionRS.Close : Set conexionRS = Nothing		
	end if	'

ELSEIF (Cint(opcion) = 13) THEN
	'
	'Fill combo Meses
	'			
	Dim rsMeses, arrMeses
	'
	' Buscar Datos de todas las Meses
	'
	Dim iSemDes, iSemHas, rsMensual
	'	
	if Cint(idCliente=1) then
		'Atenas
		QrySql = vbnullstring
		QrySql = QrySql & " SELECT Min(RS_DataProcSem.Id_Semana) AS desde, Max(RS_DataProcSem.Id_Semana) AS hasta, RS_DataProcSem.Id_Categoria FROM RS_DataProcSem"
		QrySql = QrySql & " GROUP BY RS_DataProcSem.Id_Categoria HAVING RS_DataProcSem.Id_Categoria=" & idCat
		'		
	else
		QrySql = vbnullstring		
		QrySql = QrySql & " SELECT ss_ClienteCategoria.Id_PeriodoDesde AS desde, ss_ClienteCategoria.Id_PeriodoPub AS hasta FROM ss_ClienteCategoria"
		QrySql = QrySql & " WHERE"
		QrySql = QrySql & " ss_ClienteCategoria.Id_Cliente = " & idCliente
		QrySql = QrySql & " AND"
		QrySql = QrySql & " ss_ClienteCategoria.Ind_Mensual = 1"
		QrySql = QrySql & " AND"
		QrySql = QrySql & " ss_ClienteCategoria.Id_Categoria = " & idCat
	end if
	'
	' Response.Write QrySql & "<BR><BR>"
	' Response.end
	'
	Set rsMensual = Server.CreateObject("ADODB.recordSet")
	rsMensual.Open QrySql, conexionRS
			
	if not (rsMensual.EOF and rsMensual.BOF) then		
		iSemDes = rsMensual("desde").value 'rsSemanario(0)
		iSemHas = rsMensual("hasta").value 'rsSemanario(1)
	else
		iSemDes = 0
		iSemHas = 0
	end if		
	'
	rsMensual.Close : Set rsMensual = Nothing
	'		
	QrySql = vbnullstring
	QrySql = QrySql & " SELECT ss_Periodo.IdPeriodo as id, ss_Periodo.Periodo as nombre"
	QrySql = QrySql & " FROM ss_Periodo INNER JOIN ss_Semana ON ss_Periodo.IdPeriodo = ss_Semana.Id_Periodo"
	'QrySql = QrySql & " WHERE (((ss_Semana.IdSemana)>= " & iSemDes & ") AND ((ss_Semana.IdSemana)<= " & iSemHas  & "))"
	QrySql = QrySql & " WHERE ss_Semana.IdSemana >= " & iSemDes & " AND ss_Semana.IdSemana<= " & iSemHas 
	QrySql = QrySql & " GROUP BY ss_Periodo.IdPeriodo, ss_Periodo.Periodo"
	QrySql = QrySql & " ORDER BY ss_Periodo.IdPeriodo DESC;"	
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsMeses = Server.CreateObject("ADODB.recordSet")
	rsMeses.Open QrySql, conexionRS
	'
	if not rsMeses.EOF then
		arrMeses = rsMeses.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	rsMeses.Close : Set rsMeses = Nothing
	'
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrMeses) then

		For i = 0 to ubound(arrMeses, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrMeses(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & RemoverSaltodeLinea(arrMeses(1,i)) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34) & ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"  & chr(34) & ":" & chr(34)  & "No Aplica" & chr(34) & chr(125) & chr(44)
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
	conexionRS.Close : Set conexionRS = Nothing	
	'
ELSEIF (Cint(opcion) = 14) THEN
'
	'Find Categoria es Medicina
	'				
	Dim rsMedicina
	'	
	QrySql = vbnullstring
	QrySql = " SELECT PH_CB_Categoria.Ind_Medicina as Medicina FROM dbo.PH_CB_Categoria WHERE PH_CB_Categoria.id_Categoria = " & idCat
	'		
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsMedicina = Server.CreateObject("ADODB.recordSet")
	rsMedicina.Open QrySql, conexionRS
			
	if not (rsMedicina.EOF and rsMedicina.BOF) then	
		Response.Write rsMedicina("Medicina").value							
	end if		
	'
	rsMedicina.Close : Set rsMedicina = Nothing	
	'
	' Cerrar conexiones
	'	
	conexionRS.Close : Set conexionRS = Nothing	
	'
ELSE
	' de lo Contrario
	Response.Write "error"
END IF
'
%>