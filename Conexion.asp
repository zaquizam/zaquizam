<% 
Dim conexion
Set conexion = Server.CreateObject("ADODB.Connection")

'conexion.ConnectionString= "Provider=SQLOLEDB.1;Password='PHaa11..**';Persist Security Info=True;User ID=Atenas_PH;Initial Catalog=Atenas_PH;Data Source=199.79.62.22"
'conexion.ConnectionString= "Provider=SQLOLEDB.1;Password='profit';Persist Security Info=True;User ID=profit;Initial Catalog=cacevedo_atenas;Data Source=REYESHUERTA-PC\SQL2014"
conexion.ConnectionString= "Provider=SQLOLEDB.1;Password='PHaa11..**';Persist Security Info=True;User ID=cacevedo_atenas;Initial Catalog=cacevedo_atenas;Data Source=192.185.6.37"
conexion.mode = 3
conexion.Open
	
FUNCTION RemoverSaltodeLinea(byval str)

	IF isNull(str) THEN str = "" END IF
	str = REPLACE(str,vbCr," ")			'Chr(13)
	str = REPLACE(str,vbLf," ")			'Chr(10)
	str = REPLACE(str,VbCrlf," ")		'Chr(13)+Chr(10)
	str = REPLACE(str,vbNewLine," ")	'vbNewLine
	str = REPLACE(str,vbFormFeed," ")	'Chr(12)
	str = REPLACE(str,vbTab," ")		'Chr(9)
	str = REPLACE(str,vbTab," ")		'Chr(11)
	str = REPLACE(str,"'","´")			'Comillas simples
	str = REPLACE(str,"""", "`") 		'Comillas dobles		
	str = REPLACE(str,",", " ") 		'Comillas dobles
	'
	RemoverSaltodeLinea = TRIM(str)
	'
END FUNCTION

%>