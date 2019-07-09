<body>

<%
Function format_zeros(value, zeros)
  If Len(value) < zeros Then
    Do Until Len(value) = zeros
      value = "0" & value
    Loop
  End If
  format_zeros = value
End Function

action = request.form("action")
sapid = request.form("sapid")
empflag = request.form("empflag")
userin = request.form("userin")
unitid = request.form("unitid")
workposition = request.form("workposition")
positiontype = request.form("positiontype")
technicid = request.form("technicid")
manager = request.form("manager")
//app_id_req = request.form("app_id")


//response.write(action)
//response.write("<br>")

//response.write(sapid)
//response.write("<br>")

//response.write(title)
//response.write("<br>")


		//get date time
		'Build the date string in the format yyyy-mm-dd
		somedate = Year(Date())+543 & format_zeros(Month(Date()), 2) & format_zeros(Day(Date()), 2)
		'Build the time string in the format hh:mm:ss
		sometime = format_zeros(Hour(Time()), 2) & format_zeros(Minute(Time()), 2) & format_zeros(Second(Time()), 2) 
		'Concatenate both together to build the timestamp yyyy-mm-dd hh:mm:ss
		datetimestamp = somedate & sometime
		
		//response.write(datetimestamp)
		//response.write("<br>")

	
		
	//create file log
	file_name = "../log/" & datetimestamp &"_updempwork_api_" & sapid & ".log"
	
	//file_err = "../log/" & datetimestamp &"_updempwork_api_err.log"
	  
	  	//response.write(file_name)
		//response.write("<br>")
	
	
	
	Dim objFSO, objStream,i
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objStream = objFSO.CreateTextFile(Server.MapPath(file_name),true)	
	
	//Dim objFSOErr,objStreamErr
	//Set objFSOErr = Server.CreateObject("Scripting.FileSystemObject")
	//Set objStreamErr = objFSOErr.CreateTextFile(Server.MapPath(file_err),true)	
 

//connect DB
dim con
set con=Server.CreateObject("ADODB.Connection")
con.open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=WebMed_User;Password=Warm104Pro;Initial Catalog=WebMed;Data Source=10.7.2.129"

If con.errors.count <> 0 Then   
 
  // Response.Write "Error Connect Database "
  // response.write("<br>")
    objStream.WriteLine("Error Connect Database")
else
     // Response.Write "Connected OK"
	 // response.write("<br>")



	//write file log test

	
	objStream.WriteLine("action=" & action)
	objStream.WriteLine("sapid=" & sapid)
	objStream.WriteLine("empflag=" & empflag)
	objStream.WriteLine("unitid=" & unitid)
	objStream.WriteLine("workposition=" & workposition)
	objStream.WriteLine("positiontype=" & positiontype)
	objStream.WriteLine("technicid=" & technicid)
	objStream.WriteLine("userin=" & userin)
	
	
	//Response.write("Writing")
	

		
	// update data to table employee_content
	
	if Len(technicid)=0 then
	   technicid = 0 
	end if
	
	if Len(unitid)=0 then
	   unitid = 0 
	end if
	
	if Len(workposition)=0 then
	   workposition = 0 
	end if
	
	if Len(positiontype)=0 then
	   positiontype = 0 
	end if
	
	if Len(manager) = 0 then					
		manager = 0
	end if
		   file_image = "emppic/u" & Request.Form("sapid") & ".jpg"
	
		
		 sqlup = "UPDATE employee_content SET unitid = " & unitid & " , workposition=" & workposition & ", positiontype=" & positiontype & ", upddatetime= ' " & datetimestamp & "' , technicid=" & technicid & ",image= '" & file_image & "' , manager_flag = " &  manager & "  WHERE sapid = " & sapid   
		 //sqlup = "UPDATE employee_content SET unitid = " & unitid & " WHERE sapid = " & sapid  
		
		// response.write(sqlup)
		// response.write("<br>")
		
	
		 
		 objStream.WriteLine("sqlup=" & sqlup)
		
			
			Set recup = con.Execute(sqlup)
			if err<>0 then
			//	response.write("Error update permissions!")
				 objStream.WriteLine("sqlup=" & sqlup)
				 objStream.WriteLine("Error update permissions!")
			end if


	


		//payload = "{ TotalAmount: 1, CurrencyCode: 1, TransactionId: 36985223, Business: { ApiKey: '" & key & "',Description: '' }, CreditTypes: [ {CreditTypeCode: 1, MaxNumberOfPayments: 1}] }"

		 
	
End If
	con.close
	objStream.Close
	Set objStream = Nothing
	Set objFSO = Nothing
	
	//objStreamErr.Close
	//Set objStreamErr = Nothing
	//Set objFSOErr = Nothing
	
%>
</body>