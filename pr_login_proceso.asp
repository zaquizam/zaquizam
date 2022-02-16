<%
'-- Creado: 18abr17 
'
Session.TimeOut = 60 
Dim iUsu
Dim sEma
Dim sPas
Dim conexion
Dim iPerf
Dim sPro
Dim usr
Dim pwd
Dim sExiste
Dim sExisteUsu
Dim sExisteLid
Dim sExisteAud
Dim rs
Dim mySession
'
Dim regEx
Set regEx = New RegExp
regEx.Global = true
'regEx.Pattern = "[^0-9a-zA-Z@.-_]"
'
usr = GetSecureVal(Trim(Request.form("email")))
pwd = GetSecureVal(Trim(Request.form("password")))
'
sEma = regEx.Replace(SQLInject(usr), "")
sPas = regEx.Replace(SQLInject(pwd), "")
'
sEma = Trim(Request.form("email"))
sPas = Trim(Request.form("password"))

Dim borra_cookies			
For Each borra_cookies In Request.Cookies
	Response.Cookies(borra_cookies).Expires =#May 25, 2019#						
Next
'
if mySession="" then
    mySession=Session.SessionID
    response.cookies("idsession")=mySession
    response.cookies("idsession").Expires=Date+365
end if
'
Ingresar
'response.write "<br><br>ERROR DE CONEXION A LA BASE DE DATOS (tiempo de espera)<br>"
'response.write "<br><br>CONTACTE AL ADIMINISTRADOR DEL SITIO WEB PARA MANTENIMIENTO PREVENTIVO<br>"
'response.end
'
Sub Ingresar
    if sEma<>"" and sPas<>"" then
		Response.Cookies("emailusu")=sEma
		'response.cookies("emailusu").Expires=date+365		
        Apertura		
		Periodo
		VerificarUsuario
		'Response.write "<br>Finalizó"
		'Response.end
    end if
End Sub
'
Sub Apertura	
	'Response.write "<br>Entro Apertura"
	'Response.end
	
	Set conexion = Server.CreateObject("ADODB.Connection")
	
	'conexion.ConnectionString= "Provider=SQLOLEDB.1;Password='PHaa11..**';Persist Security Info=True;User ID=Atenas_PH;Initial Catalog=Atenas_PH;Data Source=199.79.62.22"
	'conexion.ConnectionString= "Provider=SQLOLEDB.1;Password='profit';Persist Security Info=True;User ID=profit;Initial Catalog=cacevedo_atenas;Data Source=REYESHUERTA-PC\SQL2014"
	conexion.ConnectionString= "Provider=SQLOLEDB.1;Password='PHaa11..**';Persist Security Info=True;User ID=cacevedo_atenas;Initial Catalog=cacevedo_atenas;Data Source=192.185.6.37"

	'Response.write "<br>64 LLEGO"
	'Response.end

	conexion.mode = 3
	'Response.write "<br>68 LLEGO"
	'Response.end

	conexion.Open
	'Response.write "<br>72 LLEGO"
	'Response.end

	'Response.write "<br>Salio Apertura"
	'Response.end
End Sub
'
Sub Periodo
	'Response.write "<br><br><br>Entro Periodo"
	'Response.end
	'Session("periodo")  = 24218
	'Response.Cookies("periodo")= 24218
	'exit sub
	Dim Primer 
    Dim Ultimo 
	Dim iPeriodo
	Dim Fecha
	Dim sql	
	Dim nMes
	'	
	Fecha=Date()
	iAno = cInt(year(Fecha))
	iMes = cInt(month(Fecha))
	'
	if iMes=1 then nMes="Enero"
	if iMes=2 then nMes="Febrero"
	if iMes=3 then nMes="Marzo"
	if iMes=4 then nMes="Abril"
	if iMes=5 then nMes="Mayo"
	if iMes=6 then nMes="Junio"
	if iMes=7 then nMes="Julio"
	if iMes=8 then nMes="Agosto"
	if iMes=9 then nMes="Septiembre"
	if iMes=10 then nMes="Octubre"
	if iMes=11 then nMes="Noviembre"
	if iMes=12 then nMes="Diciembre"
	'
	iPeriodo = ((cInt(iAno) * 12)) + cInt(iMes)
	'	
    Primer = DateSerial(Year(Fecha), Month(Fecha) + 0, 1)  
    Ultimo = DateSerial(Year(Fecha), Month(Fecha) + 1, 0)  
	'
	sql = ""
    sql = sql & " SELECT * FROM ss_Periodo"
	sql = sql & " WHERE"
	sql = sql & " (((IdAno)=" & year(date()) & ") AND ((IdMes)=" & Month(date()) & "));"
	'
    set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = 1 'adOpenKeyset 
	rs.LockType = 3 'adLockOptimistic 
	'	
	'Response.write "<br>123 LLEGO:= " & sql
	'Response.end
	rs.Open sql,conexion
	'Response.write "<br>126 LLEGO"
	'Response.end
    '
	if rs.eof then 			
		set rs4 = CreateObject("ADODB.Recordset")
		rs4.CursorType = 1 'adOpenKeyset 
		rs4.LockType = 3 'adLockOptimistic 
		sql4="ss_Periodo"
		rs4.open sql4,conexion
		'
		rs4.addnew
		rs4("IdPeriodo") = iPeriodo
		rs4("Frecuencia") ="Mensual"
		rs4("Per_Desde") =cdate(Primer)
		rs4("Per_Hasta") =cdate(Ultimo)
		rs4("Periodo") = nMes & " " & iAno		
		rs4("idano") =iAno		
		rs4("idmes") =iMes		
		rs4("IP")=Request.ServerVariables("REMOTE_ADDR")
		rs4("Usr") = "Automatico"
		rs4("Fec_Ult_Mod") =DateAdd("n",+30,Now())
		'
		rs4.update 
		rs4.close
		set rs4=nothing		    
		'
		Session("periodo")  = iPeriodo    
		Response.Cookies("periodo")= iPeriodo
		'	
	else
		Session("periodo")  = rs("IdPeriodo")    
		Response.Cookies("periodo")= rs("IdPeriodo")
	end if
	'
	rs.close
	set rs=nothing	
	'   
	'Response.write "<br>Salio Periodo"
	'Response.end

End Sub
''
Sub VerificarUsuario	
	'Response.write "<br>Entro VerificarUsuario"
	'Response.end 
	Dim sql	
	sql = ""
    sql = sql & " Select "
    sql = sql & " Id_Usuario, "
    sql = sql & " Usuario, "
    sql = sql & " Nombres, "
    sql = sql & " Apellidos, "
    sql = sql & " pass_Clave, "
    sql = sql & " Id_PerfilUsuario, "
    sql = sql & " Id_Cliente, "
    sql = sql & " Filtro1, "
    sql = sql & " Filtro2, "
    sql = sql & " Filtro3, "
    sql = sql & " Filtro4, "
    sql = sql & " Fec_Login, "
    sql = sql & " Fec_Ult_Mod, "
    sql = sql & " IP, "	
    sql = sql & " Ind_Activo "
    sql = sql & " FROM "
    sql = sql & " ss_Usuarios "
    sql = sql & " WHERE "
    sql = sql & " Usuario = '" & sEma & "'"
    sql = sql & " AND pass_clave = '" & sPas & "'"
	sql = sql & " AND Ind_Activo = 1"
	'
    set rs = CreateObject("ADODB.Recordset")
	rs.CursorType = 1 	'adOpenKeyset 
	rs.LockType = 3 	'adLockOptimistic 
	'
	'	
	rs.Open sql, conexion
    '
	if rs.recordcount= 0 then 
	    rs.close
		sExiste=0		
		sExisteUsu=0
		'response.write "<br>" & Sql & "<br>"
		Response.write "USUARIO NO REGISTRADO, VERIFIQUE LOS DATOS INGRESADOS...!!"		
	    exit sub
	end if
	'response.write "<br>" & Sql & "<br>"
	'Response.end 
	'
	iUsu 				= rs("Id_Usuario")
	iPerf 				= rs("Id_PerfilUsuario")	'
	iPerUsu             = rs("Id_PerfilUsuario")	'
    Session("idUsu")    = rs("Id_Usuario")
    Session("Nombre")   = rs("Nombres")
    Session("Apellido") = rs("Apellidos")
    Session("Usuario")  = rs("Usuario")
    Session("email")    = rs("Usuario")
	Session("perusu")   = rs("Id_PerfilUsuario")
	Session("idCliente")= rs("Id_Cliente")
	'Response.write "<br>227 LLEGO"
	'Response.end
	
	Session("filtro1") 	= rs("Filtro1")
    'Session("filtro2") 	= rs("Filtro2")
    'Session("filtro3") 	= rs("Filtro3")
    'Session("filtro4") 	= rs("Filtro4")
	Session("NomApe") 	= rs("Nombres") & " " & rs("Apellidos")
	'
	Session("TituloApp") = "|Panel Hogares|"
	'
	'Response.write "<br>238 LLEGO"
	'Response.end
	'Response.write "<br>238 LLEGO"
	'Response.end
	Response.Cookies("cliente")=rs("Id_Cliente")
	'Response.write "<br>243 LLEGO"
	'Response.end
	'Response.write rs("Id_Cliente")
	'Response.end
	'Response.write "<br>241 LLEGO"
	'Response.end
	'Response.write "<br>249 LLEGO"
	'Response.end
	
	'Response.Cookies("filtro1")=rs("Filtro1")
	'Response.Cookies("filtro2")=rs("Filtro2")
	'Response.Cookies("filtro3")=rs("Filtro3")
	'Response.Cookies("filtro4")=rs("Filtro4")
	'Response.write "<br>256 LLEGO"
	'Response.end

	Response.Cookies("Idusu")=rs("Id_Usuario") 
	Response.Cookies("nombre")=rs("Nombres") 
	Response.Cookies("apellido")=rs("Apellidos") 
	Response.Cookies("perusu")=rs("Id_PerfilUsuario")
	'Response.Cookies("emailusu")=rs("usuario")
	'response.cookies("emailusu").Expires=date+365
	'
	if isnull(rs("Usuario")) then
	    Response.Cookies("usuario")="Not User"
	else
	    Response.Cookies("usuario")=rs("Usuario")
	end if    
	'
	sNom = rs("Nombres") & " " & rs("Apellidos")
	rs("Fec_Login")=DateAdd("n",-30,Now())
	rs("Fec_Ult_Mod")=DateAdd("n",-30,Now())
	rs("IP")=Request.ServerVariables("REMOTE_ADDR")
	rs.update
	rs.close
	BuscarEmpresa
	VerificarPerfil
	'Response.write "<br>241"
	'response.end 
	GrabarEntrada 	
	'Response.write "<br>244"
	Response.write "usuario"
	''
End Sub


Sub BuscarEmpresa
	Session("idEmpresa") = 1
	exit sub
end sub

Sub VerificarPerfil
	'Response.write "<br>446"
		sql = ""
		sql = sql & " SELECT "
		sql = sql & " id_PerfilUsuario, "
		sql = sql & " Link_Acceso "
		sql = sql & " FROM "
		sql = sql & " ss_PerfilUsuario "
		sql = sql & " WHERE "
		sql = sql & " id_PerfilUsuario = " & iPerf

		set rs = CreateObject("ADODB.Recordset")
		rs.CursorType = 1 'adOpenKeyset 
		rs.LockType = 3 'adLockOptimistic 
		rs.Open sql,conexion
		
		if rs.recordcount= 0 then 
			rs.close
			exit sub
		end if
		sPro = rs("Link_Acceso")
		rs.close		
		'Response.write "<br>sPro:=" & sPro
End Sub
'
'
Sub GrabarEntrada
	'exit sub
	'Response.write "<br>501"
	' Grabar Entrada
	set rs4 = CreateObject("ADODB.Recordset")
	rs4.CursorType = 1 'adOpenKeyset 
	rs4.LockType = 3 'adLockOptimistic 
	sql4="ss_AccesosUsuarios"
	'	
	rs4.open sql4,conexion
	rs4.addnew
	rs4("Id_Usuario")=iUsu
	rs4("Fec_Acc")=DateAdd("n",-30,Now())
	rs4("Pagina")=sPro1	
	rs4("Num_Accesos")=1
	if sReferencia <> "" then
		'rs4("Referer")=nro
		rs4("Referer")=sReferencia
	end if
	rs4("idSession")=mySession
	rs4("IP")=Request.ServerVariables("REMOTE_ADDR")

	rs4.update
	rs4.close
	'	
	'Response.write "<br>524"
	'response.end 
End sub	

Function GetSecureVal(param)
	If IsEmpty(param) Or param = "" Then
		GetSecureVal = param
		Exit Function
	End If
	
	If IsNumeric(param) Then
		GetSecureVal = CLng(param)
	Else
		GetSecureVal = Replace(CStr(param), "'", "''")
	End If
End Function

Function SQLInject(strWords) 
	dim badChars, newChars, i
	badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "=", "update") 
	newChars = strWords 
	for i = 0 to uBound(badChars) 
		newChars = replace(newChars, badChars(i), "") 
	next 
	newChars = newChars 
	newChars= replace(newChars, "'", "''")
	newChars= replace(newChars, " ", "")
	newChars= replace(newChars, "'", "|")
	newChars= replace(newChars, "|", "''")
	newChars= replace(newChars, "\""", "|")
	newChars= replace(newChars, "|", "''")
	SQLInject=newChars
End function 
%>