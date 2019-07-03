<body>
<!--
create by: sansanee
date create :01072562
1.API for insert data employee  
2.receive data from application HR infoMed when active button insert/update
3.show data on web content menu บุคลากร  automatic
-->
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
position_desc = request.form("position")
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

//sapid = 12345682

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
	file_name = "../log/" & datetimestamp &"_addemp_api.log"
	  
	  	//response.write(file_name)
		//response.write("<br>")
	Dim objFSO, objStream,i
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objStream = objFSO.CreateTextFile(Server.MapPath(file_name),true)	

dim con
set con=Server.CreateObject("ADODB.Connection")
con.open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=WebMed_User;Password=Warm104Pro;Initial Catalog=WebMed;Data Source=10.7.2.129"

If con.errors.count <> 0 Then   
 
   //Response.Write "Error Connect Database "
   //response.write("<br>")
   
   objStream.WriteLine("Error Connect Database")
else
      //Response.Write "Connected OK"
	  //response.write("<br>")
	 
	  
	  //write file log test
	  

	
	
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
	
	//Response.write("Writing")

	  
	 
	 //insert data to table employee_content
	 
	    file_image = "../emppic/u" & Request.Form("sapid") & ".jpg"
		//response.write(file_image)
		//response.write("<br>")
		
		
		
		
	 
		sql_insert="INSERT INTO employee_content (sapid,positiondesc,"
		sql_insert=sql_insert & "title,fname,lname,title_eng,fname_eng,lname_eng,adddatetime,image,empflag,sapid_mgt)"
		sql_insert=sql_insert & " VALUES "
		//sql_insert=sql_insert & "('" & Request.Form("sapid") & "',"
		sql_insert=sql_insert & "(" & sapid & ","
		sql_insert=sql_insert & "'" & Request.Form("positiondesc") & "',"
		sql_insert=sql_insert & "'" & Request.Form("title") & "',"
		sql_insert=sql_insert & "'" & Request.Form("fname") & "',"
		sql_insert=sql_insert & "'" & Request.Form("lname") & "',"
		sql_insert=sql_insert & "'" & Request.Form("title_eng") & "',"
		sql_insert=sql_insert & "'" & Request.Form("fname_eng") & "',"
		sql_insert=sql_insert & "'" & Request.Form("lname_eng") & "',"
		sql_insert=sql_insert & "'" & datetimestamp & "',"
		sql_insert=sql_insert & "'" & file_image & "',"
		//sql_insert=sql_insert & "'" & Request.Form("empflag") & "',"	
		sql_insert=sql_insert & Request.Form("empflag") & ","	
		//sql_insert=sql_insert & "'" & Request.Form("userin") & "')"
		sql_insert=sql_insert & Request.Form("userin") & ")"
		 
		// response.write(sql_insert)
		// response.write("<br>")
		
		//Set rec = con.Execute(sql_insert)
		
		//Set recup = con.Execute(sql_insert)
		//if err<>0 then
		//		response.write("Error update permissions! "  & "  error=" & err)
		//end if
		
		 objStream.WriteLine("sql_insert=" & sql_insert)
		
		on error resume next
		con.Execute sql_insert,recaffected
		if err<>0 then
		  Response.Write("No update permissions! " & "error=" & err)
		  objStream.WriteLine("sql insert  error=" & err)
		else
		  Response.Write("<h3>" & recaffected & " record added</h3>")
		  objStream.WriteLine("sql insert =" & recaffected & " record added" )
		end if
		
		con.close

		//If Not rec.EOF Then 
			//Response.write(rec.fields("app_name").value)
			//Response.write(rec.fields("app_count").value)
			//response.write("<br>")
			
				
				
				//update app_count
				
					
					//sql_insertup = "UPDATE app_access SET app_count = '" & count_access & "' WHERE app_key like '" & key_login & "' " 
					//Response.write(sql_insertup)
					//response.write("<br>")
					
					//Set recup = con.Execute(sql_insertup)
					//if err<>0 then
					//	response.write("Error update permissions!")
					//end if
				
			
	//	Else 
		//	Response.Write("<script>alert('++Not Found application ++');window.history.back();</script>")
	//	End If 

	


		//payload = "{ TotalAmount: 1, CurrencyCode: 1, TransactionId: 36985223, Business: { ApiKey: '" & key & "',Description: '' }, CreditTypes: [ {CreditTypeCode: 1, MaxNumberOfPayments: 1}] }"

	
End If
	con.close
 	objStream.Close
	Set objStream = Nothing
	Set objFSO = Nothing


%>
</body>