<!--#include file="conexion.asp" -->
<%
'
' PH_Cte_HomePantryRpSem_Fill_cmb2.asp - 02feb22 - 02feb22
'
' Cambio en combo Marca - 
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

Dim opcion, QrySql, idCat, idCliente, idMar
'
'opcion  = Cint(Request.Querystring("opcion"))
'idQuery = Cint(Request.Querystring("id"))
'
opcion = Request.Querystring("opcion")
idCat  = Request.Querystring("idCat")
idFab   = Request.Querystring("idFab")
idCliente = Request.Querystring("idCli")
'
IF (Cint(opcion) = 2) THEN
	'
	'ReFill combo Marca
	'			
	Dim hpMarca, arrMarca
	'
	' Buscar Datos de todas las Marcas
	'		
	QrySql = vbnullstring	
	QrySql = " SELECT Id_Marca as id, Marca as nombre FROM PH_DataProcesadaSem WHERE Id_Categoria = " & idCat
	if Len(idFab) <> 0 then 
		QrySql = QrySql & " AND Id_Fabricante in (" & idFab & ")"
	end if	
	QrySql = QrySql & " GROUP BY Id_Marca, Marca HAVING Id_Marca <> 0 ORDER BY Marca "	
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set hpMarca = Server.CreateObject("ADODB.recordset")
	hpMarca.Open QrySql, conexion, 0, 1
	'
	if not hpMarca.EOF then
		arrMarca = hpMarca.GetRows()  ' Convert recordset to 2D Array
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
	conexion.close : set conexion = nothing
	'		
ELSEIF (Cint(opcion) = 3) THEN
	'
	'ReFill combo Segmento
	'			
	Dim hpSegmento, arrSegmento		
	'
	' Buscar Datos de todas las Segmento
	'
	QrySql = vbnullstring	
	QrySql = " SELECT Id_Segmento AS id, Segmento AS nombre FROM cacevedo_atenas.PH_DataProcesadaSem WHERE PH_DataProcesadaSem.Id_Segmento <> 0 AND PH_DataProcesadaSem.Id_Categoria = " & idCat	
	if Len(idFab) <> 0 then 
		QrySql = QrySql & " AND 	PH_DataProcesadaSem.Id_Fabricante in (" & idFab & ")"
	end if	
	QrySql = QrySql & " GROUP BY	PH_DataProcesadaSem.Id_Segmento, PH_DataProcesadaSem.Segmento ORDER BY PH_DataProcesadaSem.Segmento ASC"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set hpSegmento = Server.CreateObject("ADODB.recordset")
	hpSegmento.Open QrySql, conexion, 0, 1
	'
	if not hpSegmento.EOF then
		arrSegmento = hpSegmento.GetRows()  ' Convert recordset to 2D Array
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
		sTabla  =   chr(123) &  chr(34) & "id"		& chr(34)& ":" & chr(34)	& "0"				& chr(34)	& chr(44)
		sTabla  =   sTabla   &  chr(34) & "nombre"	& chr(34)& ":" & chr(34)	& "No hay Datos"	& chr(34)	& chr(125)	& chr(44)
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

%>