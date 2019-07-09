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
title = request.form("title")
fname = request.form("fname")
mname = request.form("mname")
lname = request.form("lname")
title_eng = request.form("title_eng")
fname_eng = request.form("fname_eng")
mname_eng = request.form("mname_eng")
lname_eng = request.form("lname_eng")
empflag = request.form("empflag")
userin = request.form("userin")
positiondesc = request.form("positiondesc")


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
	  
	  	//response.write(file_name)
		//response.write("<br>")
	Dim objFSO, objStream,i
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objStream = objFSO.CreateTextFile(Server.MapPath(file_name),true)	
 

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
	objStream.WriteLine("title=" & title)
	objStream.WriteLine("positiondesc=" & positiondesc)
	objStream.WriteLine("fname=" & fname)
	objStream.WriteLine("mname=" & mname)
	objStream.WriteLine("lname=" & lname)
	objStream.WriteLine("title_eng=" & title_eng)
	objStream.WriteLine("fname_eng=" & fname_eng)
	objStream.WriteLine("mname_eng=" & mname_eng)
	objStream.WriteLine("lname_eng=" & lname_eng)
	objStream.WriteLine("userin=" & userin)
	
	
	//Response.write("Writing")
	

		
	// update data to table employee_content
	
	if Len(techicid)=0 then
	   techicid = 0 
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
	
	
	
		
		 //sqlupemp = "UPDATE employee_content SET unitid = " & unitid & " , workposition=" & workposition & ", positiontype=" & positiontype & ", upddatetime= ' " & datetimestamp & "' , technicid=" & techicid & " WHERE sapid = " & sapid   
		
		  sqlupemp="UPDATE employee_content SET "
		  sqlupemp=sqlupemp & "positiondesc='" & Request.Form("positiondesc") & "',"
		  sqlupemp=sqlupemp & "title='" & Request.Form("title") & "',"
		  sqlupemp=sqlupemp & "fname='" & Request.Form("fname") & "',"
		  sqlupemp=sqlupemp & "lname='" & Request.Form("lname") & "',"
		  sqlupemp=sqlupemp & "title_eng='" & Request.Form("title_eng") & "',"
		  sqlupemp=sqlupemp & "fname_eng='" & Request.Form("fname_eng") & "',"
		  sqlupemp=sqlupemp & "lname_eng='" & Request.Form("lname_eng") & "',"
		  sqlupemp=sqlupemp & "upddatetime='" & datetimestamp & "',"
		  sqlupemp=sqlupemp & "empflag=" & empflag 
		  sqlupemp=sqlupemp & " WHERE sapid=" & sapid 
		
		// response.write(sqlup)
		// response.write("<br>")
		 
		 objStream.WriteLine("sqlupemp=" & sqlupemp)
		
			Set recup = con.Execute(sqlupemp)
			if err<>0 then
			//	response.write("Error update permissions!")
				 objStream.WriteLine("Error update permissions!")
			end if

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
	con.close
	objStream.Close
	Set objStream = Nothing
	Set objFSO = Nothing
%>
</body>