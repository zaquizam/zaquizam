<!--#include file="conexion.asp" -->
<%
'
' llenar_cmb_convivencias.asp - 24may21 - 16jun21
'
Session.lcid = 1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'
if conexion.errors.count <> 0 Then
  Response.Write ("No hay Conexion con la BD...!")
  Response.End
end if

Dim opcion, QrySql
'
'opcion  = Cint(Request.Querystring("opcion"))
'idQuery = Cint(Request.Querystring("id"))
'
opcion  = Request.Form("opcion")
'opcion  = Request.Querystring("opcion")
'
'Response.write(opcion) & "<br>"
'Response.END

'
IF (Cint(opcion) = 1) THEN
	'
	'Fill combo Categoria A y B
	'				
	Dim rsCategoria, arrCategoria
	'
	' Buscar Datos de todas las Categorias
	'
	QrySql = vbnullstring
	QrySql = " SELECT DISTINCT Id_categoria AS id, categoria AS nombre" & _
	" FROM PH_DataCrudaMensual" & _
	" ORDER BY Categoria ASC"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsCategoria = Server.CreateObject("ADODB.recordset")
	rsCategoria.Open QrySql, conexion
	'
	if not rsCategoria.EOF then
		arrCategoria = rsCategoria.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsCategoria.Close : Set rsCategoria = Nothing
	'
	'Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
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
	conexion.close : set conexion = nothing
	'
ELSEIF (Cint(opcion) = 2) THEN
'
	'Fill combo Areas
	'			
	Dim rsArea, arrArea
	'
	' Buscar Datos de todas las Areas
	'
	QrySql = vbnullstring
	QrySql = " SELECT DISTINCT Id_Area AS id, Area AS nombre " & _
	" FROM PH_DataCrudaMensual WHERE id_area <> 0" & _
	" ORDER BY Area ASC"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsArea = Server.CreateObject("ADODB.recordset")
	rsArea.Open QrySql, conexion
	'
	if not rsArea.EOF then
		arrArea = rsArea.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsArea.Close : Set rsArea = Nothing
	'
	'Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
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
	conexion.close : set conexion = nothing
	'	
ELSEIF (Cint(opcion) = 3) THEN
	'
	'Fill combo Periodo
	'			
	Dim rsPeriodo, arrPeriodo
	'
	' Buscar Datos de todas las Periodos
	'
	QrySql = vbnullstring
	QrySql = " SELECT ss_Periodo.Semanas, ss_Periodo.Periodo" & _
	" FROM" & _
	" (PH_DataCrudaMensual INNER JOIN ss_Semana ON PH_DataCrudaMensual.Id_Semana = ss_Semana.IdSemana)" & _
	" INNER JOIN ss_Periodo ON ss_Semana.Id_Periodo = ss_Periodo.IdPeriodo" & _
	" GROUP BY" & _
	" ss_Periodo.Semanas, ss_Periodo.Periodo, ss_Periodo.IdPeriodo" & _
	" ORDER BY" & _
	" ss_Periodo.IdPeriodo DESC;"
	'
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsPeriodo = Server.CreateObject("ADODB.recordset")
	rsPeriodo.Open QrySql, conexion
	'
	if not rsPeriodo.EOF then
		arrPeriodo = rsPeriodo.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	rsPeriodo.Close : Set rsPeriodo = Nothing
	'
	'Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
	sTabla = vbnullstring

	if IsArray(arrPeriodo) then

		For i = 0 to ubound(arrPeriodo, 2)
			'
			sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrPeriodo(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrPeriodo(1,i) & chr(34) & chr(125) &chr(44)
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