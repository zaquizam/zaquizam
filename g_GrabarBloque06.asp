
<%@language=vbscript%>

<!--#include file="Conexion.asp"-->

<%
Session.LCID = 8202 
'==========================================================================================
' Variables y Constantes
'==========================================================================================
	'response.write "<br>LLEGO25" 
	'response.end
	ynum=Request.QueryString("num") 
	ysx01=Request.QueryString("sx01")
	ysx02=Request.QueryString("sx02")
	ysx03=Request.QueryString("sx03")
	ysx04=Request.QueryString("sx04")
	ysx05=Request.QueryString("sx05")
	ysx06=Request.QueryString("sx06")
	ysx07=Request.QueryString("sx07")
	ysx08=Request.QueryString("sx08")
	ysx09=Request.QueryString("sx09")
	ysx10=Request.QueryString("sx10")
	ysx11=Request.QueryString("sx11")
	ysx12=Request.QueryString("sx12")
	ysx13=Request.QueryString("sx13")
	ysx14=Request.QueryString("sx14")
	ysx15=Request.QueryString("sx15")
	ysx16=Request.QueryString("sx16")
	ysx17=Request.QueryString("sx17")
	ysx18=Request.QueryString("sx18")
	ysx19=Request.QueryString("sx19")
	ysx20=Request.QueryString("sx20")
	ysx21=Request.QueryString("sx21")
	ysx22=Request.QueryString("sx22")
	ysx23=Request.QueryString("sx23")
	ysx24=Request.QueryString("sx24")
	ysx25=Request.QueryString("sx25")
	ysx26=Request.QueryString("sx26")
	ysx27=Request.QueryString("sx27")
	ysx28=Request.QueryString("sx28")
	ysx29=Request.QueryString("sx29")
	ysx30=Request.QueryString("sx30")
	
	dim rsx3
	set rsx3 = CreateObject("ADODB.Recordset")
	rsx3.CursorType = 0
	rsx3.LockType = 3

	sql = ""
	sql = sql & " Select * from PH_PanelHogar "
	sql = sql & " Where Id_PanelHogar = " & cint(ynum)
	'response.write "<br>220 sql:=" & sql
	'response.end
	rsx3.Open sql ,conexion
	rsx3("Id_DomesticaFija") = ysx01
	rsx3("Id_PersonalLabores") = ysx02
	rsx3("Id_DomesticaDia") = ysx03
	rsx3("id_ConexionInternet1") = ysx04
	rsx3("id_ConexionInternet2") = ysx05
	rsx3("id_ConexionInternet3") = ysx06
	rsx3("id_CelularJefe") = ysx07
	rsx3("id_SeguroHCMParticular") = ysx08
	rsx3("id_SeguroHCMColectivo") = ysx09
	rsx3("id_SeguroHCMSS") = ysx10
	rsx3("Id_AireAcondicionado") = ysx11
	rsx3("Id_Calentador1") = ysx12
	rsx3("Id_Calentador2") = ysx13
	rsx3("Id_Computador1") = ysx14
	rsx3("Id_Computador2") = ysx15
	rsx3("Id_DVD") = ysx16
	rsx3("Id_HomeTheater") = ysx17
	rsx3("Id_JuegosVodeo") = ysx18
	rsx3("Id_HornoMicro") = ysx19
	rsx3("Id_Secadora") = ysx20
	rsx3("Id_Lavadora1") = ysx21
	rsx3("Id_Lavadora2") = ysx22
	rsx3("Id_Lavadora3") = ysx23
	rsx3("Id_Nevera") = ysx24
	rsx3("Id_Freezer") = ysx25
	rsx3("Id_Cocina1") = ysx26
	rsx3("Id_Cocina2") = ysx27
	rsx3("Id_Cocina3") = ysx28
	rsx3("Id_Cocina4") = ysx29
	rsx3("Id_LavaPlato") = ysx30
	rsx3.Update
	rsx3.Close 
	'set rsx3 = nothing 
	
%>