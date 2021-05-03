<%@language=vbscript%>
<!--#include file="Conexion.asp"-->
<%
	' g_rRevInvBuscarAltaHogar.asp
	'
	' 05ene21 - 08ene21 
	'
	Session.lcid = 1034
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
    '
	idHogar		= Request.QueryString("id_Hogar")
    '
    ' Buscar Alta Hogar 
    '
	Dim rsAltaHogar
	set rsAltaHogar = CreateObject("ADODB.Recordset")
	rsAltaHogar.CursorType = adOpenKeyset 
	rsAltaHogar.LockType = 2 'adLockOptimistic 
	'
	'sql = vbnullString
	'sql = sql & " SELECT FORMAT (Fec_Registro, 'dd-MM-yyyy ') FROM PH_PanelHogar WHERE Id_PanelHogar= " & CInt(idHogar)
	'
	sql = vbnullString
	sql = sql & " SELECT"
	sql = sql & " PH_Panelistas.Nombre1,"
	sql = sql & " PH_Panelistas.Apellido1,"
	sql = sql & " PH_Panelistas.Celular,"
	sql = sql & " FORMAT (PH_PanelHogar.Fec_Registro, 'dd/MM/yyyy ') AS fecha"
	sql = sql & " FROM"
	sql = sql & " PH_PanelHogar"
	sql = sql & " LEFT JOIN PH_Panelistas ON PH_Panelistas.Id_Hogar = PH_PanelHogar.Id_PanelHogar"
	sql = sql & " WHERE"
	sql = sql & " PH_Panelistas.ResponsablePanel = 1"
	sql = sql & " AND"
	sql = sql & " PH_PanelHogar.Id_PanelHogar ="& idHogar
	'
	'response.write sql
	'response.end
	'
    rsAltaHogar.Open sql ,conexion
	'
	if not rsAltaHogar.EOF then
		arrAltaHogar = rsAltaHogar.GetRows()  ' Convert recordset to 2D Array
	end if
	'			
	Response.ContentType = "application/json"		
	'
	sTabla=vbnullstring
	
	if IsArray(arrAltaHogar) then
	
		For i = 0 to ubound(arrAltaHogar, 2)
		
			sTabla    =    chr(123) &  chr(34) & "nombre"   & chr(34) & ":" & chr(34) & arrAltaHogar(0,i) & chr(34) & chr(44)
			sTabla    =    sTabla   &  chr(34) & "apellido"	& chr(34) & ":" & chr(34) & arrAltaHogar(1,i) & chr(34) & chr(44)
			sTabla    =    sTabla   &  chr(34) & "celular"	& chr(34) & ":" & chr(34) & arrAltaHogar(2,i) & chr(34) & chr(44)
			'
			sTabla    =    sTabla  &  chr(34) & "fecha"     & chr(34) & ":" & chr(34) & arrAltaHogar(3,i) & chr(34) & chr(125) & chr(44)
			
			sTablaJson = sTablaJson & sTabla
			sTabla=vbnullstring
			
		next				
		'
	else
		'Eof()
		'sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		'
		sTabla    =    chr(123) &  chr(34) & "nombre"   & chr(34) & ":" & chr(34) & "NO APLICA" & chr(34) & chr(44)
		sTabla    =    sTabla   &  chr(34) & "apellido"	& chr(34) & ":" & chr(34) & "NO APLICA" & chr(34) & chr(44)
		sTabla    =    sTabla   &  chr(34) & "celular"	& chr(34) & ":" & chr(34) & "NO APLICA" & chr(34) & chr(44)
		'
		sTabla    =    sTabla  &  chr(34) & "fecha"     & chr(34) & ":" & chr(34) & "NO APLICA" & chr(34) & chr(125) & chr(44)
		'		
		sTablaJson = sTablaJson & sTabla
		sTabla=vbnullstring
		'
	end if
	'	
	sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
	JsonData= chr(91) & sTabla & chr(93) '& chr(125)
	Response.Write(JsonData)
	'
	conexion.Close    
    Set conexion = Nothing
	'
%>