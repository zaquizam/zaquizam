<% 
	dim conexionRS
	Set conexionRS = Server.CreateObject("ADODB.Connection")
	
	conexionRS.ConnectionString= "Provider=SQLOLEDB.1;Password='PHaa11..**';Persist Security Info=True;User ID=cacevedo_atenas;Initial Catalog=RetailScannig;Data Source=216.198.73.34"
	conexionRS.mode = 3
	conexionRS.Open

%>