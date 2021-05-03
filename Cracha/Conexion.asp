<% 
	dim Conexion
	Set conexion = Server.CreateObject("ADODB.Connection")
	
	conexion.ConnectionString= "Provider=SQLOLEDB.1;Password='Adm1npru3b4';Persist Security Info=True;User ID=ADPRUEBA;Initial Catalog=XPRUEBAZ;Data Source=184.168.194.78"
	conexion.mode = 3
	conexion.Open

%>