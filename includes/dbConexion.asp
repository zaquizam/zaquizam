<% 
Dim conexionBD
Set conexionBD = Server.CreateObject("ADODB.Connection")
'conexionBD.ConnectionString= "Provider=SQLOLEDB.1;Password='profit';Persist Security Info=True;User ID=profit;Initial Catalog=Venicom;Data Source=WIN-SGOTTCKL2NB"
conexionBD.ConnectionString= "Provider=SQLOLEDB.1;Password='PHaa11..**';Persist Security Info=True;User ID=Atenas_PH;Initial Catalog=Atenas_PH;Data Source=199.79.62.22"

conexionBD.mode = 3	
conexionBD.Open	
Function timeStamp()
	Dim t 
	t = Now
	timeStamp = Month(t) & "/" & Day(t) & "/" & Year(t)    
End Function
%>