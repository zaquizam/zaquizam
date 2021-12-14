<%
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