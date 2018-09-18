<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Login_Check.asp"-->
<!--#include file="fpdf.asp"-->
<!--#include file="DBconn.asp"-->
<%

AdmitTerm=Request("term")

LastUpdatedTime = Time()
LastUpdatedDt = date()

Set pdf=CreateJsObject("FPDF")
pdf.CreatePDF()
pdf.SetPath("fpdf/")
pdf.LoadExtension("table")
pdf.SetFont "Arial","",16
pdf.Open()
pdf.LoadModels("TestModels") 
pdf.AddPage()

pdf.ChapterTitle2("                                    Current MPH Students  ""      "  &LastUpdatedDt&" "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(3)




'////// MPH Students ////////
set rs=Server.CreateObject("ADODB.recordset")
mph_query="SELECT Count(distinct UIN) mph_students FROM CurrentStudents where ProgramType='MPH' "
rs.Open mph_query,conn
'pdf.ChapterBody(Total_Student)

pdf.OrangeTitle("Program Option MPH")

'pdf.FancyTable()

'//////// MPH ////////////
mph_rows = rs("mph_students")
mph_cols = 5
Dim mph_col(5)
mph_col(1) = "Banner # "
mph_col(2) = "First Name"
mph_col(3) = "Last Name"
mph_col(4) = "Admit Term"
mph_col(5) = "Current Year"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admit Term", "Current Year"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_query="SELECT UIN, Firstname, LastName, AdmitTerm, CurrentYear FROM CurrentStudents where ProgramType='MPH' order by LastName"
					rs.Open mph_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("AdmitTerm"),"|",",")
                    e = Replace(rs("CurrentYear"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPH_Student = "Total number of students in MPH : "&mph_rows
    pdf.ChapterBody(Total_MPH_Student)
    pdf.Ln(5)

'//////// MPH-FT students  ////////////

set rs=Server.CreateObject("ADODB.recordset")
mphft_query="SELECT Count(distinct UIN) mphft_students FROM CurrentStudents where ProgramType='MPH-FT'  "
rs.Open mphft_query,conn
'pdf.ChapterBody(Total_Student)

pdf.OrangeTitle("Program Option MPH-FT")

'pdf.FancyTable()

'//////// MPH-FT ////////////
mphft_rows = rs("mphft_students")
mphft_cols = 5
Dim mphft_col(5)
mphft_col(1) = "Banner # "
mphft_col(2) = "First Name"
mphft_col(3) = "Last Name"
mphft_col(4) = "Admit Term"
mphft_col(5) = "Current Year"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admit Term", "Current Year"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_ft_query="SELECT UIN, Firstname, LastName, AdmitTerm, CurrentYear FROM CurrentStudents where ProgramType='MPH-FT' order by LastName"
					rs.Open mph_ft_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("AdmitTerm"),"|",",")
                    e = Replace(rs("CurrentYear"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPH_FT_Student = "Total number of students in MPH-FT : "&mphft_rows
    pdf.ChapterBody(Total_MPH_FT_Student)
    pdf.Ln(5)

'//////// MPH-PM students ////////////

set rs=Server.CreateObject("ADODB.recordset")
mphpm_query="SELECT Count(distinct UIN) mphpm_students FROM CurrentStudents where ProgramType='MPH-PM'  "
rs.Open mphpm_query,conn
'pdf.ChapterBody(Total_Student)

pdf.OrangeTitle("Program Option MPH-PM")

'pdf.FancyTable()

'//////// MPH-PM ////////////
mphpm_rows = rs("mphpm_students")
mphpm_cols = 5
Dim mphpm_col(5)
mphpm_col(1) = "Banner # "
mphpm_col(2) = "First Name"
mphpm_col(3) = "Last Name"
mphpm_col(4) = "Admit Term"
mphpm_col(5) = "Current Year"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admit Term", "Current Year"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_pm_query="SELECT UIN, Firstname, LastName, AdmitTerm, CurrentYear FROM CurrentStudents where ProgramType='MPH-PM' order by LastName"
					rs.Open mph_pm_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("AdmitTerm"),"|",",")
                    e = Replace(rs("CurrentYear"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPH_PM_Student = "Total number of students in MPH-PM : "&mphpm_rows
    pdf.ChapterBody(Total_MPH_PM_Student)
    pdf.Ln(5)


    
'////// MPH-Adv Students ////////
pdf.OrangeTitle("Program Option MPH-ADV")

set rs=Server.CreateObject("ADODB.recordset")
mphadv_query="SELECT Count(distinct UIN) mphadv_students FROM CurrentStudents where ProgramType='MPH-ADV'  "
rs.Open mphadv_query,conn
'pdf.ChapterBody(Total_Student)

'pdf.FancyTable()

'//////// MPH-ADV ////////////
mphadv_rows = rs("mphadv_students")
mphadv_cols = 5
Dim mphadv_col(5)
mphadv_col(1) = "Banner # "
mphadv_col(2) = "First Name"
mphadv_col(3) = "Last Name"
mphadv_col(4) = "Admit Term"
mphadv_col(5) = "Current Year"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admit Term", "Current Year"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_adv_query="SELECT UIN, Firstname, LastName, AdmitTerm, CurrentYear FROM CurrentStudents where ProgramType='MPH-ADV' order by LastName"
					rs.Open mph_adv_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("AdmitTerm"),"|",",")
                    e = Replace(rs("CurrentYear"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPH_ADV_Student = "Total number of students in MPH-ADV : "&mphadv_rows
    pdf.ChapterBody(Total_MPH_ADV_Student)
    pdf.Ln(5)



    total_accepted=mphadv_rows+mphpm_rows+mphft_rows+mph_rows
pdf.GreyTitle("")
accepted_students = "Total Students : "&total_accepted
pdf.ChapterBody(accepted_students)
pdf.Ln(5)



pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
