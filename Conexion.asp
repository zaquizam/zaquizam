<% 
	dim conexion
	Set conexion = Server.CreateObject("ADODB.Connection")
	
	'conexion.ConnectionString= "Provider=SQLOLEDB.1;Password='PHaa11..**';Persist Security Info=True;User ID=Atenas_PH;Initial Catalog=Atenas_PH;Data Source=199.79.62.22"
	'conexion.ConnectionString= "Provider=SQLOLEDB.1;Password='profit';Persist Security Info=True;User ID=profit;Initial Catalog=cacevedo_atenas;Data Source=REYESHUERTA-PC\SQL2014"
	conexion.ConnectionString= "Provider=SQLOLEDB.1;Password='PHaa11..**';Persist Security Info=True;User ID=cacevedo_atenas;Initial Catalog=cacevedo_atenas;Data Source=192.185.6.37"
	conexion.mode = 3
	conexion.Open

%>