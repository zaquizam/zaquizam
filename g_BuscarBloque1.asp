<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ynum=Request.Form("id")
	'ynum=Request.QueryString("id")

	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Nombre1, "				' 0
	sql = sql & " Nombre2, "				' 1
	sql = sql & " Apellido1, "				' 2
	sql = sql & " Apellido2, "				' 3
	sql = sql & " Id_Nacionalidad, "		' 4
	sql = sql & " Cedula, "					' 5
	sql = sql & " Celular, "				' 6
	sql = sql & " CelularAdicional, "		' 7
	sql = sql & " NumeroCortesia, "			' 8
	sql = sql & " Correo, "					' 9
	sql = sql & " CorreoAlterno, "			'10
	sql = sql & " Id_Parentesco, "			'11
	sql = sql & " Id_EstadoCivil, "			'12
	sql = sql & " Fec_Nacimiento, "			'13
	sql = sql & " Id_Sexo, "				'14
	sql = sql & " Id_Educacion, "			'15
	sql = sql & " Id_TipoIngreso, "			'16
	sql = sql & " CantidadPersonas, "		'17
	sql = sql & " Id_FrecuenciaCompra, "	'18
	sql = sql & " Titular, "				'19
	sql = sql & " Id_Banco, "				'20
	sql = sql & " NumeroCuenta, "			'21
	sql = sql & " Id_PagoRapido, "			'22
	sql = sql & " CedulaTitular "			'23
	sql = sql & " FROM "
	sql = sql & " PH_Panelistas "
	sql = sql & " WHERE "
	sql = sql & " Id_Hogar = " & ynum 
	sql = sql & " AND ResponsablePanel = 1 "
	'response.write "<br>220 sql:=<br>" & sql & "<br><br>"
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br> Linea 223 " &
	'response.end
	iExiste = 0
	'
	if not rsx1.EOF then
		arrBloque1 = rsx1.GetRows()  ' Convert recordset to 2D Array
	end if
		'
	rsx1.Close
	Set rsPanelConsumo = Nothing
	sTabla=vbnullstring
	if IsArray(arrBloque1) then
			For i = 0 to ubound(arrBloque1, 2)
				sTabla = chr(123)&  chr(34) & "Nombre1"				& chr(34)& ":" & chr(34) & arrBloque1(0,i)  & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Nombre2" 			& chr(34)& ":" & chr(34) & arrBloque1(1,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Apellido1" 			& chr(34)& ":" & chr(34) & arrBloque1(2,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Apellido2" 			& chr(34)& ":" & chr(34) & arrBloque1(3,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Nacionalidad"		& chr(34)& ":" & chr(34) & arrBloque1(4,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Cedula" 				& chr(34)& ":" & chr(34) & arrBloque1(5,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Celular"	 			& chr(34)& ":" & chr(34) & arrBloque1(6,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "CelularAdicional"	& chr(34)& ":" & chr(34) & arrBloque1(7,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "NumeroCortesia"		& chr(34)& ":" & chr(34) & arrBloque1(8,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Correo" 				& chr(34)& ":" & chr(34) & arrBloque1(9,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "CorreoAlterno" 		& chr(34)& ":" & chr(34) & arrBloque1(10,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Parentesco" 		& chr(34)& ":" & chr(34) & arrBloque1(11,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_EstadoCivil"		& chr(34)& ":" & chr(34) & arrBloque1(12,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Fec_Nacimiento"		& chr(34)& ":" & chr(34) & arrBloque1(13,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Sexo"				& chr(34)& ":" & chr(34) & arrBloque1(14,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Educacion"		& chr(34)& ":" & chr(34) & arrBloque1(15,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_TipoIngreso"		& chr(34)& ":" & chr(34) & arrBloque1(16,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "CantidadPersonas"	& chr(34)& ":" & chr(34) & arrBloque1(17,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_FrecuenciaCompra"	& chr(34)& ":" & chr(34) & arrBloque1(18,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Titular"				& chr(34)& ":" & chr(34) & arrBloque1(19,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Banco"			& chr(34)& ":" & chr(34) & arrBloque1(20,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "NumeroCuenta"		& chr(34)& ":" & chr(34) & arrBloque1(21,i)	& chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_PagoRapido"		& chr(34)& ":" & chr(34) & arrBloque1(22,i)	& chr(34) & chr(44)
				'
				sTabla = sTabla  &  chr(34) & "CedulaTitular"       & chr(34)& ":" & chr(34) & arrBloque1(23,i) & chr(34) & chr(125)&chr(44)
				sTablaJson = sTablaJson & sTabla
				sTabla=vbnullstring
			next
			sTabla = Left(sTablaJson, Len(sTablaJson) - 1) 'Devuelve "Cadena"
			JsonData= chr(91) & sTabla & chr(93) '& chr(125)
		else
			'Eof()
			sTablaJson = sTablaJson & sTabla
			sTabla=vbnullstring
			JsonData=chr(123) & chr(34)& "data" & chr(34)& ":" & chr(91) & sTabla & chr(93) & chr(125)
		end if
		Response.Write(JsonData)
		conexion.close
		set conexion = nothing
%>