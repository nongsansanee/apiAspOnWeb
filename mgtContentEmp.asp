<body>
<!--
	$emprank=$_POST["emprank"];
	$emprankname=$_POST["emprankname"];
	$empengrankname=$_POST["empengrankname"];
	$empengrank=$_POST["empengrank"];
	$empname=$_POST["empname"];
	$empsname=$_POST["empsname"];
	$empengname=$_POST["empengname"];
	$empengsname=$_POST["empengsname"];
	$empflag=$_POST["empflag"];
	$empleave=be2bc($_POST["empleave"]);
	$datein=date("Y-m-d H:i:s");
	$userin=$myauth["userin"];
	$emptype=$_POST["emptype"];
	$worktype=$_POST["worktype"];
	$positiontype=$_POST["positiontype"];
	$workposition=$_POST["workposition"];
	$workunit=$_POST["workunit"];
	$follow=$_POST["follow"];
	$txtsearch=$_POST["txtsearch"];
	-->
<%
action = request.form("action")
sapid = request.form("sapid")
title = request.form("title")
position_desc = request.form("position")
fname = request.form("fname")
mname = request.form("mname")
lname = request.form("lname")
title_eng = request.form("title_eng")
fname_eng = request.form("fname_eng")
mname_eng = request.form("mname_eng")
lname_eng = request.form("lname_eng")
userin = request.form("userin")
//app_id_req = request.form("app_id")


response.write(action)
response.write("<br>")

response.write(sapid)
response.write("<br>")

response.write(title)
response.write("<br>")



 

//iMode = 8

//set oFs = server.createobject("Scripting.FileSystemObject")
//set oTextFile = oFs.OpenTextFile("Test.txt", iMode, True)
//oTextFile.Write "Test Content"
//oTextFile.Close
//set oTextFile = nothing
//set oFS = nothing
 
 

//connect DB
dim con
set con=Server.CreateObject("ADODB.Connection")
con.open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=WebMed_User;Password=Warm104Pro;Initial Catalog=WebMed;Data Source=10.7.2.129"

If con.errors.count <> 0 Then   
 
   Response.Write "Error Connect Database "
   response.write("<br>")
else
      Response.Write "Connected OK"
	  response.write("<br>")





	Dim objFSO, objStream,i
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	//Set objStream = objFSO.CreateTextFile(Server.MapPath("../log/testlog.txt"),true)	
	
	Set objStream =objFSO.OpenTextFile(Server.MapPath("../log/testlog.txt",8,true))
	
	
	
	objStream.WriteLine("action=" & action)
	objStream.WriteLine("sapid=" & sapid)
	objStream.WriteLine("title=" & title)
	objStream.WriteLine("position_desc=" & position_desc)
	objStream.WriteLine("fname=" & fname)
	objStream.WriteLine("mname=" & mname)
	objStream.WriteLine("lname=" & lname)
	objStream.WriteLine("title_eng=" & title_eng)
	objStream.WriteLine("fname_eng=" & fname_eng)
	objStream.WriteLine("mname_eng=" & mname_eng)
	objStream.WriteLine("lname_eng=" & lname_eng)
	objStream.WriteLine("userin=" & userin)
	objStream.WriteLine("-----------------------------------")
	Response.write("Writing")
	objStream.Close
	Set objStream = Nothing
	Set objFSO = Nothing

		
		 
		// sql_sel = " SELECT * FROM app_access  WHERE app_key like '" & key_login & "' " 
		 
		//  response.write(sql_sel)
		// response.write("<br>")
		
		///Set rec = con.Execute(sql_sel)

	//	If Not rec.EOF Then 
			//Response.write(rec.fields("app_name").value)
			//Response.write(rec.fields("app_count").value)
			//response.write("<br>")
			
				
				
				//update app_count
				
					
					//sqlup = "UPDATE app_access SET app_count = '" & count_access & "' WHERE app_key like '" & key_login & "' " 
					//Response.write(sqlup)
					//response.write("<br>")
					
					//Set recup = con.Execute(sqlup)
					//if err<>0 then
					//	response.write("Error update permissions!")
					//end if
				
			
	//	Else 
			//Response.Write("<script>alert('++Not Found application ++');window.history.back();</script>")
	//	End If 

	


		//payload = "{ TotalAmount: 1, CurrencyCode: 1, TransactionId: 36985223, Business: { ApiKey: '" & key & "',Description: '' }, CreditTypes: [ {CreditTypeCode: 1, MaxNumberOfPayments: 1}] }"

		 
	
End If
%>
</body>