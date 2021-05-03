<!--#include file="../includes/inc_validar_user.asp" -->
<!--#include file="../includes/inc_dbconexion.asp" -->
<%
' 06-10-2020 18:04:03 - 27-10-2020 19:48:05
'
Session.lcid=1034
Response.CodePage = 65001
Response.CharSet = "utf-8"
'
Dim opcion, id, QrySql
'
'opcion  = Request.Querystring("opcion")
'idQuery = Request.Querystring("id")
'
opcion  = Request.Form("opcion")
idQuery = Request.Form("id")

IF (opcion=1) THEN
	'
	'Fill combo Forma de Pago
	'
	Dim rsFormaPago
	'
    if conexionBD.errors.count <> 0 Then
		Response.Write ("no hay conexion...!")
		Response.End
  	end If
	'
	' Buscar Datos de todas las Formas de Pago
	'
	QrySql = vbnullstring
    QrySql = QrySql & " SELECT"
    QrySql = QrySql & " PH_FormaPago.Id_FormaPago AS id,"
    QrySql = QrySql & " PH_FormaPago.FormaPago AS nombre"
    QrySql = QrySql & " FROM"
    QrySql = QrySql & " PH_FormaPago"
    QrySql = QrySql & " WHERE"
    QrySql = QrySql & " PH_FormaPago.Ind_Activo = 1"
    QrySql = QrySql & " AND"
    QrySql = QrySql & " PH_FormaPago.Id_Moneda =" & idQuery
    QrySql = QrySql & " ORDER BY"
    QrySql = QrySql & " PH_FormaPago.FormaPago ASC"
    '
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsFormaPago = Server.CreateObject("ADODB.recordset")
	rsFormaPago.Open QrySql,conexionBD
	'
	if not rsFormaPago.EOF then
    	arrFormaPago = rsFormaPago.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
	sTabla=""

    if IsArray(arrFormaPago) then

        For i = 0 to ubound(arrFormaPago, 2)
            '
            sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrFormaPago(0,i)  & chr(34) & chr(44)
            sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrFormaPago(1,i)  & chr(34) & chr(125)&chr(44)
            sTablaJson = sTablaJson & sTabla
            sTabla=""
            '
        next

    else
        'Eof()
        sTabla    =   chr(123)&  chr(34) & "id" 			& chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
        sTabla    =   sTabla &  chr(34) & "nombre"         & chr(34)& ":" & chr(34)  & "No Aplica" 	& chr(34) & chr(125)&chr(44)
        ''
        sTablaJson = sTablaJson & sTabla
        sTabla=""

    end if
	''
	sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'
	rsFormaPago.Close
	Set rsFormaPago = Nothing
	'
	conexionBD.close
	set conexionBD = nothing
	'

ELSEIF (opcion=2) THEN
 '
	Dim rsCadena
	'
    if conexionBD.errors.count <> 0 Then
		Response.Write ("no hay conexion...!")
		Response.End
  	end If
	'
	' Buscar Datos de todas las Cadenas Registrados
	'
	QrySql = vbnullstring
    QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_Cadena.Id_Cadena AS id,"
	QrySql = QrySql & " PH_Cadena.Cadena AS nombre"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Cadena"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Cadena.Ind_Activo = 1"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_Cadena.Id_Canal =" & idQuery
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_Cadena.Cadena ASC"
    '
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsCadena = Server.CreateObject("ADODB.recordset")
	rsCadena.Open QrySql,conexionBD
	'
	if not rsCadena.EOF then
    	arrCadena = rsCadena.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
	sTabla=""

    if IsArray(arrCadena) then

        For i = 0 to ubound(arrCadena, 2)
            '
            sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrCadena(0,i)  & chr(34) & chr(44)
            sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrCadena(1,i)  & chr(34) & chr(125)&chr(44)
            sTablaJson = sTablaJson & sTabla
            sTabla=""
            '
        next

    else
        'Eof()
        sTabla    =   chr(123)&  chr(34) & "id" 			& chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
        sTabla    =   sTabla &  chr(34) & "nombre"         & chr(34)& ":" & chr(34)  & "No Aplica" 	& chr(34) & chr(125)&chr(44)
        ''
        sTablaJson = sTablaJson & sTabla
        sTabla=""

    end if
	''
	sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'
	rsCadena.Close
	Set rsCadena = Nothing
	'
	conexionBD.close
	set conexionBD = nothing
	'
ELSEIF (opcion=3) THEN
 	' Fill combo Categoria
	Dim rsCategoria
	'
    if conexionBD.errors.count <> 0 Then
		Response.Write ("no hay conexion...!")
		Response.End
  	end If
	'
	' Buscar Datos de todas las Cadenas Registrados
	'
	QrySql = vbnullstring
    QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_Categoria.Id_Categoria AS id,"
	QrySql = QrySql & " PH_Categoria.Categoria AS nombre"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_Categoria"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_Categoria.Ind_Activo = 1"
	'QrySql = QrySql & " AND"
	'QrySql = QrySql & " PH_Cadena.Id_Canal =" & idQuery
	QrySql = QrySql & " ORDER BY"
	QrySql = QrySql & " PH_Categoria.Categoria ASC"
    '
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsCategoria = Server.CreateObject("ADODB.recordset")
	rsCategoria.Open QrySql,conexionBD
	'
	if not rsCategoria.EOF then
    	arrCategoria = rsCategoria.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
	sTabla=""

    if IsArray(arrCategoria) then

        For i = 0 to ubound(arrCategoria, 2)
            '
            sTabla    =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrCategoria(0,i)  & chr(34) & chr(44)
            sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrCategoria(1,i)  & chr(34) & chr(125)&chr(44)
            sTablaJson = sTablaJson & sTabla
            sTabla=""
            '
        next

    else
        'Eof()
        sTabla    =   chr(123)&  chr(34) & "id" 			& chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
        sTabla    =   sTabla &  chr(34) & "nombre"         & chr(34)& ":" & chr(34)  & "No Aplica" 	& chr(34) & chr(125)&chr(44)
        ''
        sTablaJson = sTablaJson & sTabla
        sTabla=""

    end if
	''
	sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	Response.Write(JsonData)
	'
	' Cerrar conexiones
	'
	rsCategoria.Close
	Set rsCategoria = Nothing
	'
	conexionBD.close
	set conexionBD = nothing
	'
ELSEIF (opcion=4) THEN
	'
	' Fill Recolectar Todos los Datos
	'
	Dim rsRecolectar
	'
    if conexionBD.errors.count <> 0 Then
		Response.Write ("no hay conexion...!")
		Response.End
  	end If
	'
	' Buscar Datos de todas las Cadenas Registrados
	'
	QrySql = vbnullstring
    QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_TipoConsumo.Recoleccion_completa"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_TipoConsumo"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_TipoConsumo.Ind_Activo = 1"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_TipoConsumo.Id_TipoConsumo =" & idQuery
    '
	'Response.Write QrySql & "<BR><BR>"
	'Response.end
	'
	Set rsRecolectar = Server.CreateObject("ADODB.recordset")
	rsRecolectar.Open QrySql,conexionBD
	'
	if not rsRecolectar.EOF then
    	arrRecolectar = rsRecolectar.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
	sTabla=""

    if IsArray(arrRecolectar) then

        For i = 0 to ubound(arrRecolectar, 2)
            '
            sTabla     =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrRecolectar(0,i)  & chr(34) & chr(125)&chr(44) '& chr(44)
            'sTabla    =    sTabla &  chr(34) & "nombre" & chr(34)& ":" & chr(34) & arrCategoria(1,i)  & chr(34) & chr(125)&chr(44)
            sTablaJson = sTablaJson & sTabla
            sTabla=""
            '
			Response.Write arrRecolectar(0,i)
        next

    else
        'Eof()
        ' sTabla    =   chr(123)&  chr(34) & "id" 			& chr(34)& ":" & chr(34) & "0" 			& chr(34) & chr(44)
        ' sTabla    =   sTabla &  chr(34) & "nombre"         & chr(34)& ":" & chr(34)  & "No Aplica" 	& chr(34) & chr(125)&chr(44)
		sTabla     =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & "0" & chr(34) & chr(125)&chr(44) '& chr(44)
        ''
        sTablaJson = sTablaJson & sTabla
        sTabla=""

		Response.Write 0

    end if
	''
	'sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	'JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
	'Response.Write(JsonData)
	'
	' Cerrar conexiones
	'
	rsRecolectar.Close
	Set rsRecolectar = Nothing
	'
	conexionBD.close
	set conexionBD = nothing
	'
ELSEIF (opcion=5) THEN
	'
	' Fill Recolectar Todos los Datos
	'
	Dim rsComida
	'
    if conexionBD.errors.count <> 0 Then
		Response.Write ("no hay conexion...!")
		Response.End
  	end If
	'
	' Buscar Datos de todas las Cadenas Registrados
	'
	QrySql = vbnullstring
    QrySql = QrySql & " SELECT"
	QrySql = QrySql & " PH_TipoConsumo.Comida_status"
	QrySql = QrySql & " FROM"
	QrySql = QrySql & " PH_TipoConsumo"
	QrySql = QrySql & " WHERE"
	QrySql = QrySql & " PH_TipoConsumo.Ind_Activo = 1"
	QrySql = QrySql & " AND"
	QrySql = QrySql & " PH_TipoConsumo.Id_TipoConsumo =" & idQuery
    '
	Set rsComida = Server.CreateObject("ADODB.recordset")
	rsComida.Open QrySql,conexionBD
	'
	if not rsComida.EOF then
    	arrComida = rsComida.GetRows()  ' Convert recordset to 2D Array
	end if
	'
	Response.ContentType = "application/json"
	''
	'Crear Archivo Array Json
	''
	sTabla=""

    if IsArray(arrComida) then

        For i = 0 to ubound(arrComida, 2)
            '
            sTabla     =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & arrComida(0,i)  & chr(34) & chr(125)&chr(44)
            sTablaJson = sTablaJson & sTabla
            sTabla=""
            '
			Response.Write arrComida(0,i)
        next

    else
        'Eof()
		sTabla     =   chr(123)&  chr(34) & "id" 	& chr(34)& ":" & chr(34) & "0" & chr(34) & chr(125)&chr(44) '& chr(44)
        ''
        sTablaJson = sTablaJson & sTabla
        sTabla=""

		Response.Write 0

    end if	
	'
	' Cerrar conexiones
	'
	rsComida.Close
	Set rsComida = Nothing
	'
	conexionBD.close
	set conexionBD = nothing
	'	
 ELSE

	' de lo Contrario

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