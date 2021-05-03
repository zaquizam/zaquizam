<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	ynum=Request.Form("id")

	dim gDatosSol
	dim rsx1
	set rsx1 = CreateObject("ADODB.Recordset")
	rsx1.CursorType = adOpenKeyset 
	rsx1.LockType = 2 'adLockOptimistic 

	sql = ""
	sql = sql & " SELECT "
	sql = sql & " Id_DomesticaFija, "		'0
	sql = sql & " Id_PersonalLabores, "		'1
	sql = sql & " Id_DomesticaDia, "		'2
	sql = sql & " id_ConexionInternet1, "	'3
	sql = sql & " id_ConexionInternet2, "	'4
	sql = sql & " id_ConexionInternet3, "	'5
	sql = sql & " id_CelularJefe, "			'6
	sql = sql & " id_SeguroHCMParticular, "	'7
	sql = sql & " id_SeguroHCMColectivo, "	'8
	sql = sql & " id_SeguroHCMSS, "			'9
	sql = sql & " Id_AireAcondicionado, "	'10
	sql = sql & " Id_Calentador1, "			'11
	sql = sql & " Id_Calentador2, "			'12
	sql = sql & " Id_Computador1, "			'13
	sql = sql & " Id_Computador2, "			'14
	sql = sql & " Id_DVD, "					'15
	sql = sql & " Id_HomeTheater, "			'16
	sql = sql & " Id_JuegosVodeo, "			'17
	sql = sql & " Id_HornoMicro, "			'18
	sql = sql & " Id_Secadora, "			'19
	sql = sql & " Id_Lavadora1, "			'20
	sql = sql & " Id_Lavadora2, "			'21
	sql = sql & " Id_Lavadora3, "			'22
	sql = sql & " Id_Nevera, "				'23
	sql = sql & " Id_Freezer, "				'24
	sql = sql & " Id_Cocina1, "				'25
	sql = sql & " Id_Cocina2, "				'26
	sql = sql & " Id_Cocina3, "				'27
	sql = sql & " Id_Cocina4, "				'28
	sql = sql & " Id_LavaPlato "			'29
	sql = sql & " FROM "
	sql = sql & " PH_PanelHogar "
	sql = sql & " WHERE "
	sql = sql & " Id_PanelHogar = " & ynum
	'response.write "<br>220 sql:=" & sql
	'response.end
    rsx1.Open sql ,conexion
	'response.write "<br> Linea 223 " &
	'response.end
	iExiste = 0
	'
	if not rsx1.EOF then
		arrBloque6 = rsx1.GetRows()  ' Convert recordset to 2D Array
	end if
		'
	rsx1.Close
	Set rsPanelConsumo = Nothing
	sTabla=vbnullstring
	if IsArray(arrBloque6) then
			For i = 0 to ubound(arrBloque6, 2)
				sTabla = chr(123)&  chr(34) & "Id_DomesticaFija"		& chr(34)& ":" & chr(34) & arrBloque6(0,i)  & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_PersonalLabores"		& chr(34)& ":" & chr(34) & arrBloque6(1,i)  & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_DomesticaDia"			& chr(34)& ":" & chr(34) & arrBloque6(2,i)  & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "id_ConexionInternet1"	& chr(34)& ":" & chr(34) & arrBloque6(3,i)  & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "id_ConexionInternet2"	& chr(34)& ":" & chr(34) & arrBloque6(4,i)  & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "id_ConexionInternet3"	& chr(34)& ":" & chr(34) & arrBloque6(5,i)  & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "id_CelularJefe"			& chr(34)& ":" & chr(34) & arrBloque6(6,i)  & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "id_SeguroHCMParticular"	& chr(34)& ":" & chr(34) & arrBloque6(7,i)  & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "id_SeguroHCMColectivo"	& chr(34)& ":" & chr(34) & arrBloque6(8,i)  & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "id_SeguroHCMSS"			& chr(34)& ":" & chr(34) & arrBloque6(9,i)  & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_AireAcondicionado"	& chr(34)& ":" & chr(34) & arrBloque6(10,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Calentador1"			& chr(34)& ":" & chr(34) & arrBloque6(11,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Calentador2"			& chr(34)& ":" & chr(34) & arrBloque6(12,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Computador1"			& chr(34)& ":" & chr(34) & arrBloque6(13,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Computador2"			& chr(34)& ":" & chr(34) & arrBloque6(14,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_DVD"					& chr(34)& ":" & chr(34) & arrBloque6(15,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_HomeTheater"			& chr(34)& ":" & chr(34) & arrBloque6(16,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_JuegosVodeo"			& chr(34)& ":" & chr(34) & arrBloque6(17,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_HornoMicro"			& chr(34)& ":" & chr(34) & arrBloque6(18,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Secadora"				& chr(34)& ":" & chr(34) & arrBloque6(19,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Lavadora1"			& chr(34)& ":" & chr(34) & arrBloque6(20,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Lavadora2"			& chr(34)& ":" & chr(34) & arrBloque6(21,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Lavadora3"			& chr(34)& ":" & chr(34) & arrBloque6(22,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Nevera"				& chr(34)& ":" & chr(34) & arrBloque6(23,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Freezer"				& chr(34)& ":" & chr(34) & arrBloque6(24,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Cocina1"				& chr(34)& ":" & chr(34) & arrBloque6(25,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Cocina2"				& chr(34)& ":" & chr(34) & arrBloque6(26,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Cocina3"				& chr(34)& ":" & chr(34) & arrBloque6(27,i) & chr(34) & chr(44)
				sTabla = sTabla  &  chr(34) & "Id_Cocina4"				& chr(34)& ":" & chr(34) & arrBloque6(28,i) & chr(34) & chr(44)
				'
				sTabla = sTabla  &  chr(34) & "Id_LavaPlato"        	& chr(34)& ":" & chr(34) & arrBloque6(29,i) & chr(34) & chr(125)&chr(44)
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