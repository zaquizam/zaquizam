<!--#include file="conexion.asp" -->
<%
'
'PH_Cte_HomePantryRpMen_Fill_cmb1.asp - 09feb22 - 09feb22
'
Server.ScriptTimeout = 10000
Response.Buffer = True	
Session.lcid = 1034
Response.CodePage = 65001
Response.CharSet = "UTF-8"	
'
if conexion.errors.count <> 0 Then
  Response.Write ("No hay conexion con la BD...!")
  Response.End
end if

Dim opcion, QrySql, idCat, idCliente
'
opcion = Request.Querystring("opcion")
idCat  = Request.Querystring("idCat")
idCliente = Request.Querystring("idCli")
'
IF (Cint(opcion) = 1) THEN
	'
	'Fill combo Categoria
	'				
	Dim hpCategoria, arrCategoria
	'
	' Buscar Datos de todas las Categorias
	' 03nov21
	'
	QrySql = vbnullstring
	IF (Cint(idCliente) = 1 ) THEN		
		QrySql = " SELECT PH_DataProcesadaSem.Id_Categoria AS id,  PH_DataProcesadaSem.Categoria AS nombre FROM PH_DataProcesadaSem GROUP BY PH_DataProcesadaSem.Id_Categoria, PH_DataProcesadaSem.Categoria ORDER BY PH_DataProcesadaSem.Categoria ASC"
	ELSE		
		QrySql = " SELECT PH_DataProcesadaSem.Id_Categoria AS id, PH_DataProcesadaSem.Categoria AS nombre FROM PH_DataProcesadaSem INNER JOIN ss_ClienteCategoria ON  PH_DataProcesadaSem.Id_Categoria = ss_ClienteCategoria.Id_Categoria WHERE ss_ClienteCategoria.Id_Cliente = " & idCliente & " and ss_ClienteCategoria.Ind_Semanal = 1 GROUP BY PH_DataProcesadaSem.Id_Categoria, PH_DataProcesadaSem.Categoria ORDER BY PH_DataProcesadaSem.Categoria ASC"		
	END IF			
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set hpCategoria = Server.CreateObject("ADODB.recordSet")
	hpCategoria.Open QrySql, conexion, 0, 1
	'
	if not hpCategoria.EOF then
		arrCategoria = hpCategoria.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	hpCategoria.Close : Set hpCategoria = Nothing
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
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No hay Datos" & chr(34) & chr(125) & chr(44)
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
	conexion.Close : Set conexion = Nothing
	'
ELSEIF (Cint(opcion) = 2) THEN
	'
	'Fill combo Fabricante
	'			
	Dim hpFabricante, arrFabricante	
	'
	' Buscar Datos de todas las Fabricantes
	'
	QrySql = vbnullstring		
	'27ene22
	QrySql = "SELECT Id_Fabricante AS id, Fabricante AS nombre FROM PH_DataProcesadaSem WHERE Id_Categoria = " & idCat & " GROUP BY Id_Fabricante, Fabricante HAVING Id_Fabricante <> 0 ORDER BY Fabricante"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set hpFabricante = Server.CreateObject("ADODB.recordSet")
	hpFabricante.Open QrySql, conexion, 0, 1
	'
	if not hpFabricante.EOF then
		arrFabricante = hpFabricante.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	hpFabricante.Close : Set hpFabricante = Nothing
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
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No hay Datos" & chr(34) & chr(125) & chr(44)
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
	conexion.Close : Set conexion = Nothing
	'		
ELSEIF (Cint(opcion) = 3) THEN
	'
	'Fill combo Marca
	'			
	Dim hpMarca, arrMarca
	
	'idCat = Request.Form("idCat")
	'
	' Buscar Datos de todas las Canales
	'
	' if idCat >= 127 and idCat <= 145 then
		' QrySql = vbnullstring			
		' QrySql = " SELECT Id_Marca as id, Trim(Marca)+'('+Trim(Fabricante)+')' as nombre FROM PH_DataProcesadaSem WHERE Id_Fabricante <> 0 AND Id_Categoria = " & idCat
		' QrySql = QrySql & " GROUP BY Id_Marca, Trim(Marca)+'('+Trim(Fabricante)+')' HAVING Id_Marca <> 0 ORDER BY Trim(Marca) + '('+Trim(Fabricante)+')'"
	' else 
		QrySql = vbnullstring			
		QrySql = " SELECT Id_Marca as id, Marca as nombre FROM PH_DataProcesadaSem WHERE Id_Categoria = " & idCat & " GROUP BY Id_Marca, Marca HAVING Id_Marca <> 0 ORDER BY Marca"
	' end if
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set hpMarca = Server.CreateObject("ADODB.recordSet")
	hpMarca.Open QrySql, conexion, 0, 1
	'
	if not hpMarca.EOF then
		arrMarca = hpMarca.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	hpMarca.Close : Set hpMarca = Nothing
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
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No hay Datos" & chr(34) & chr(125) & chr(44)
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
	conexion.Close : Set conexion = Nothing
	'		
ELSEIF (Cint(opcion) = 4) THEN
	'
	'Fill combo Segmento
	'			
	Dim hpSegmento, arrSegmento
	'
	' Buscar Datos de todas las Canales
	'
	QrySql = vbnullstring		
	QrySql = " SELECT Id_Segmento as id, Segmento as nombre FROM PH_DataProcesadaSem WHERE Id_Categoria = " & idCat & " GROUP BY Id_Segmento, Segmento HAVING Id_Segmento <> 0 ORDER BY Segmento "
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set hpSegmento = Server.CreateObject("ADODB.recordSet")
	hpSegmento.Open QrySql, conexion, 0, 1
	'
	if not hpSegmento.EOF then
		arrSegmento = hpSegmento.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	hpSegmento.Close : Set hpSegmento = Nothing
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
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34)& ":" & chr(34)  & "No hay Datos" & chr(34) & chr(125) & chr(44)
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
	conexion.Close : Set conexion = Nothing
	'		
ELSEIF (Cint(opcion) = 5) THEN
	'
	'Fill combo Indicadores
	'			
	Dim hpIndicadores, arrIndicadores
	'
	' Buscar Datos de todas las Indicadores
	'	
	QrySql = vbnullstring	
	QrySql = " SELECT Id_Indicador AS id, Abreviatura AS nombre FROM PH_Indicadores WHERE "
	'if idCliente = 1 then 
	if CInt(idCliente) = 1 then
		QrySql = QrySql & " Ind_Atenas = 1 " 
	else
		QrySql = QrySql & " Ind_Sem = 1 " 
	end if	
	'
	QrySql = QrySql & " AND Ind_Activo = 1 ORDER BY Id_Indicador "		
	'
	'Response.Write QrySql & "<BR><BR>" '& idCat
	'Response.end
	'
	Set hpIndicadores = Server.CreateObject("ADODB.recordSet")
	hpIndicadores.Open QrySql, conexion, 0, 1
	'
	if not hpIndicadores.EOF then
		arrIndicadores = hpIndicadores.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	hpIndicadores.Close : Set hpIndicadores = Nothing
	'
	'Crear Archivo Array Json
	'
	sTabla = vbnullstring

	if IsArray(arrIndicadores) then

		For i = 0 to ubound(arrIndicadores, 2)
			'
			sTabla     = chr(123)&  chr(34) & "id" 	& chr(34) & ":" & chr(34) & arrIndicadores(0,i) & chr(34) & chr(44)
			sTabla     = sTabla &  chr(34) & "nombre" & chr(34) & ":" & chr(34) & arrIndicadores(1,i) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbnullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		 & chr(34) & ":" & chr(34) & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"   & chr(34) & ":" & chr(34) & "No hay Datos" & chr(34) & chr(125) & chr(44)
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
	conexion.Close : Set conexion = Nothing
	'
ELSEIF (Cint(opcion) = 6) THEN
	'
	' Fill combo Semanas
	'			
	Dim hpSemanas, arrSemanas, iSemanaDes, iSemanaHas, hpSemanario
	'	
	if Cint(idCliente=1) then
		'atenas
		QrySql = vbnullstring
		QrySql = " SELECT Min(PH_DataProcesadaSem.Id_Semana) AS desde, Max(PH_DataProcesadaSem.Id_Semana) AS hasta, PH_DataProcesadaSem.Id_Categoria FROM PH_DataProcesadaSem GROUP BY PH_DataProcesadaSem.Id_Categoria HAVING PH_DataProcesadaSem.Id_Categoria=" & idCat
		'		
	else		
		'27ene22
		QrySql = vbnullstring		
		QrySql = " SELECT ss_ClienteCategoria.Id_semanaDesde AS desde, ss_ClienteCategoria.Id_semanaPub AS hasta FROM ss_ClienteCategoria WHERE ss_ClienteCategoria.Ind_semanal = 1 AND ss_ClienteCategoria.Id_Cliente = " & idCliente & " AND ss_ClienteCategoria.Id_Categoria = " & idCat
	end if
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set hpSemanario = Server.CreateObject("ADODB.recordSet")
	hpSemanario.Open QrySql, conexion, 0, 1
			
	if not (hpSemanario.EOF and hpSemanario.BOF) then		
		iSemanaDes = hpSemanario("desde").value 'hpSemanario(0)
		iSemanaHas = hpSemanario("hasta").value 'hpSemanario(1)
	else
		iSemanaDes = 0
		iSemanaHas = 0
	end if		
	'
	hpSemanario.Close : Set hpSemanario = Nothing
	'
	' Buscar Datos de todas las Semanas
	'		
	'27ene22
	QrySql = vbnullstring
	QrySql = " SELECT IdSemana as id, Semana as nombre FROM ss_Semana "
	if( iSemanaDes <> 0 and iSemanaHas <> 0 ) then
		QrySql = QrySql & " WHERE IdSemana >= " & iSemanaDes & " And IdSemana <= " & iSemanaHas
	end if	
	QrySql = QrySql & " ORDER BY IdSemana DESC "
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set hpSemanas = Server.CreateObject("ADODB.recordSet")
	hpSemanas.Open QrySql, conexion, 0, 1
	'
	if not hpSemanas.EOF then
		arrSemanas = hpSemanas.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	hpSemanas.Close : Set hpSemanas = Nothing
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
	conexion.Close : Set conexion = Nothing
	'
ELSEIF (Cint(opcion) = 7) THEN
	'
	'Fill combo Meses
	'			
	Dim hpMeses, arrMeses
	'
	' Buscar Datos de todas las Meses
	'
	Dim iSemDes, iSemHas, rsMensual
	'	
	if Cint(idCliente=1) then
		'Atenas
		qrySql = vbNullstring
		qrySql = " SELECT Min(PH_DataProcesadaSem.Id_Semana) AS desde, Max(PH_DataProcesadaSem.Id_Semana) AS hasta, PH_DataProcesadaSem.Id_Categoria FROM PH_DataProcesadaSem GROUP BY PH_DataProcesadaSem.Id_Categoria HAVING PH_DataProcesadaSem.Id_Categoria=" & idCat
		'		
	else
		qrySql = vbNullstring		
		qrySql = " SELECT ss_ClienteCategoria.Id_PeriodoDesde AS desde, ss_ClienteCategoria.Id_PeriodoPub AS hasta FROM ss_ClienteCategoria WHERE ss_ClienteCategoria.Id_Cliente = " & idCliente & " AND ss_ClienteCategoria.Ind_Mensual = 1 AND ss_ClienteCategoria.Id_Categoria = " & idCat
	end if
	'
	' Response.Write qrySql & "<BR><BR>"
	' Response.end
	'
	Set rsMensual = Server.CreateObject("ADODB.recordSet")
	rsMensual.Open qrySql, conexion, 0, 1
			
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
	qrySql = vbNullstring
	qrySql = " SELECT ss_Periodo.IdPeriodo as id, ss_Periodo.Periodo as nombre FROM ss_Periodo INNER JOIN ss_Semana ON ss_Periodo.IdPeriodo = ss_Semana.Id_Periodo"
	qrySql = qrySql & " WHERE ss_Semana.IdSemana >= " & iSemDes & " AND ss_Semana.IdSemana<= " & iSemHas & " GROUP BY ss_Periodo.IdPeriodo, ss_Periodo.Periodo ORDER BY ss_Periodo.IdPeriodo DESC;"	
	'
	'Response.Write qrySql & "<BR><BR>"
	'Response.end
	'
	Set hpMeses = Server.CreateObject("ADODB.recordSet")
	hpMeses.Open qrySql, conexion, 0, 1
	'
	if not hpMeses.EOF then
		arrMeses = hpMeses.GetRows()  ' Convert recordSet to 2D Array
	end if
	'
	hpMeses.Close : Set hpMeses = Nothing
	'
	'Crear Archivo Array Json
	'
	sTabla = vbNullstring

	if IsArray(arrMeses) then

		For i = 0 to ubound(arrMeses, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrMeses(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & RemoverSaltodeLinea(arrMeses(1,i)) & chr(34) & chr(125) &chr(44)
			sTablaJson = sTablaJson & sTabla
			sTabla = vbNullstring
			'
		Next

	else
		'Eof()
		sTabla  =   chr(123) &  chr(34) & "id" 		& chr(34) & ":" & chr(34)  & "0" 		& chr(34) & chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"  & chr(34) & ":" & chr(34)  & "No hay datos" & chr(34) & chr(125) & chr(44)
		''
		sTablaJson = sTablaJson & sTabla
		sTabla = vbNullstring

	end if
	''
	sTabla   = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData = chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'	
	conexion.Close : Set conexion = Nothing	
	'
ELSEIF (Cint(opcion) = 12) THEN	
 	'
	' Verificar cliente contrato el servicio				
	'
	Dim hpCliente, arrCliente
	'
	if (CInt(idCliente) = 1) then
	
		Response.Write CInt(idCliente)
	
	else
		QrySql = vbnullstring		
		QrySql = " SELECT COUNT(Id_Cliente) as total FROM ss_ClienteCategoria WHERE ss_ClienteCategoria.Id_Cliente = " & idCliente & " AND ss_ClienteCategoria.Ind_semanal = 1"
		'
		'Response.Write QrySql & "<BR><BR>"
		'Response.end
		'
		Set hpCliente = Server.CreateObject("ADODB.recordSet")
		hpCliente.Open QrySql, conexion, 0, 1
		'
		if not (hpCliente.EOF and hpCliente.BOF) then
			Response.Write hpCliente(0)
		else			
			Response.Write 0			
		end if		
		'
		hpCliente.Close : Set hpCliente = Nothing
		'
		conexion.Close : Set conexion = Nothing		
	end if	
	'
ELSE
	' de lo Contrario
	Response.Write "error"
END IF
'
%>