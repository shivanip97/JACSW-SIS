<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Login_Check.asp"-->
<!--#include file="fpdf.asp"-->
<!--#include file="DBconn.asp"-->
<%

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

pdf.ChapterTitle2("                  Current Students Concentration Change Report  "&LastUpdatedDt&" "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(10)



'////// Adv Students ////////
pdf.OrangeTitle("Program Option ADV")

set rs=Server.CreateObject("ADODB.recordset")
adv_query="SELECT Count(distinct UIN) adv_students FROM CurrentStudents where ProgramType='AdV' and (Graduated != 'Y'  or Graduated is null)"
rs.Open adv_query,conn

adv_rows = rs("adv_students")
adv_cols = 5
Dim adv_col(5)
adv_col(1) = "Banner # "
adv_col(2) = "First Name"
adv_col(3) = "Last Name"
adv_col(4) = "Current Concentration"
adv_col(5) = "Applicant Concentration"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,40,40
'pdf.SetAligns "C", "L", "L", "C","C"
pdf.SetFont "Arial","B",10
pdf.Row "UIN","First name","Last Name", "Applicant Concentration","Current Concentration"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					adv_query="SELECT c.UIN, c.Firstname, c.LastName, isnull( c.Concentration,'') As CurrentConcentration,  isnull( a.Concentration,'') As ApplicantConcentration FROM CurrentStudents c left outer join Applicants a on c.UIN = a.UIN and a.Application_Status is not null and a.Application_Status != 'IN-Incomplete' and c.ProgramType <> '' and c.AdmitTerm <> '' and (c.Graduated != 'Y'  or c.Graduated is null) and c.Status not in ('LOA', 'TRANS','WDN') and c.Decision <> 'DF' and c.ProgramType = 'Adv' Group By c.UIN, c.Concentration,a.Concentration, c.LastName, c.FirstName order by Lastname"
					rs.Open adv_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("ApplicantConcentration"),"|",",")
                    e = Replace(rs("CurrentConcentration"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_Adv_Student = "Total number of students in AdV : "&adv_rows
    pdf.ChapterBody(Total_Adv_Student)
    pdf.Ln(1)

'//////// FT students  ////////////
pdf.OrangeTitle("Program Option FT")

set rs=Server.CreateObject("ADODB.recordset")
FT_query="SELECT Count(distinct UIN) FT_students FROM CurrentStudents where ProgramType='FT' and (Graduated != 'Y'  or Graduated is null)"
rs.Open FT_query,conn

FT_rows = rs("FT_students")
FT_cols = 5
Dim FT_col(5)
FT_col(1) = "Banner # "
FT_col(2) = "First Name"
FT_col(3) = "Last Name"
FT_col(4) = "Current Concentration"
FT_col(5) = "Applicant Concentration"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,40,40
'pdf.SetAligns "C", "L", "L", "C","C"
pdf.SetFont "Arial","B",10
pdf.Row "UIN","First name","Last Name", "Applicant Concentration","Current Concentration"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					FT_query="SELECT c.UIN, c.Firstname, c.LastName, isnull( c.Concentration,'') As CurrentConcentration,  isnull( a.Concentration,'') As ApplicantConcentration FROM CurrentStudents c left outer join Applicants a on c.UIN = a.UIN and a.Application_Status is not null and a.Application_Status != 'IN-Incomplete' and c.ProgramType <> '' and c.AdmitTerm <> '' and (c.Graduated != 'Y'  or c.Graduated is null) and c.Status not in ('LOA', 'TRANS','WDN') and c.Decision <> 'DF' and c.ProgramType = 'FT' Group By c.UIN, c.Concentration,a.Concentration, c.LastName, c.FirstName order by Lastname"
					rs.Open FT_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("ApplicantConcentration"),"|",",")
                    e = Replace(rs("CurrentConcentration"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_FT_Student = "Total number of students in FT : "&FT_rows
    pdf.ChapterBody(Total_FT_Student)
    pdf.Ln(1)

'//////// MPH-FT students  ////////////

pdf.OrangeTitle("Program Option MPH-FT")

set rs=Server.CreateObject("ADODB.recordset")
MPHFT_query="SELECT Count(distinct UIN) MPHFT_students FROM CurrentStudents where ProgramType='MPH-FT' and (Graduated != 'Y'  or Graduated is null)"
rs.Open MPHFT_query,conn

MPHFT_rows = rs("MPHFT_students")
MPHFT_cols = 5
Dim MPHFT_col(5)
MPHFT_col(1) = "Banner # "
MPHFT_col(2) = "First Name"
MPHFT_col(3) = "Last Name"
MPHFT_col(4) = "Current Concentration"
MPHFT_col(5) = "Applicant Concentration"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,40,40
'pdf.SetAligns "C", "L", "L", "C","C"
pdf.SetFont "Arial","B",10
pdf.Row "UIN","First name","Last Name", "Applicant Concentration","Current Concentration"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					MPHFT_query="SELECT c.UIN, c.Firstname, c.LastName, isnull( c.Concentration,'') As CurrentConcentration,  isnull( a.Concentration,'') As ApplicantConcentration FROM CurrentStudents c left outer join Applicants a on c.UIN = a.UIN and a.Application_Status is not null and a.Application_Status != 'IN-Incomplete' and c.ProgramType <> '' and c.AdmitTerm <> '' and (c.Graduated != 'Y'  or c.Graduated is null) and c.Status not in ('LOA', 'TRANS','WDN') and c.Decision <> 'DF' and c.ProgramType = 'MPH-FT' Group By c.UIN, c.Concentration,a.Concentration, c.LastName, c.FirstName order by Lastname"
					rs.Open MPHFT_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("ApplicantConcentration"),"|",",")
                    e = Replace(rs("CurrentConcentration"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPHFT_Student = "Total number of students in MPH-FT : "&MPHFT_rows
    pdf.ChapterBody(Total_MPHFT_Student)
    pdf.Ln(1)

'////// MPH-PM Students ////////
pdf.OrangeTitle("Program Option MPH-PM")

set rs=Server.CreateObject("ADODB.recordset")
MPHPM_query="SELECT Count(distinct UIN) MPHPM_students FROM CurrentStudents where ProgramType='MPH-PM' and (Graduated != 'Y'  or Graduated is null)"
rs.Open MPHPM_query,conn

MPHPM_rows = rs("MPHPM_students")
MPHPM_cols = 5
Dim MPHPM_col(5)
MPHPM_col(1) = "Banner # "
MPHPM_col(2) = "First Name"
MPHPM_col(3) = "Last Name"
MPHPM_col(4) = "Current Concentration"
MPHPM_col(5) = "Applicant Concentration"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,40,40
'pdf.SetAligns "C", "L", "L", "C","C"
pdf.SetFont "Arial","B",10
pdf.Row "UIN","First name","Last Name", "Applicant Concentration","Current Concentration"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					MPHPM_query="SELECT c.UIN, c.Firstname, c.LastName, isnull( c.Concentration,'') As CurrentConcentration,  isnull( a.Concentration,'') As ApplicantConcentration FROM CurrentStudents c left outer join Applicants a on c.UIN = a.UIN and a.Application_Status is not null and a.Application_Status != 'IN-Incomplete' and c.ProgramType <> '' and c.AdmitTerm <> '' and (c.Graduated != 'Y'  or c.Graduated is null) and c.Status not in ('LOA', 'TRANS','WDN') and c.Decision <> 'DF' and c.ProgramType = 'MPH-PM' Group By c.UIN, c.Concentration,a.Concentration, c.LastName, c.FirstName order by Lastname"
					rs.Open MPHPM_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("ApplicantConcentration"),"|",",")
                    e = Replace(rs("CurrentConcentration"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPHPM_Student = "Total number of students in MPH-PM : "&MPHPM_rows
    pdf.ChapterBody(Total_MPHPM_Student)
    pdf.Ln(1)

'////// PM Students ////////
pdf.OrangeTitle("Program Option PM")

set rs=Server.CreateObject("ADODB.recordset")
PM_query="SELECT Count(distinct UIN) PM_students FROM CurrentStudents where ProgramType='PM' and (Graduated != 'Y'  or Graduated is null)"
rs.Open PM_query,conn

PM_rows = rs("PM_students")
PM_cols = 5
Dim PM_col(5)
PM_col(1) = "Banner # "
PM_col(2) = "First Name"
PM_col(3) = "Last Name"
PM_col(4) = "Current Concentration"
PM_col(5) = "Applicant Concentration"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,40,40
'pdf.SetAligns "C", "L", "L", "C","C"
pdf.SetFont "Arial","B",10
pdf.Row "UIN","First name","Last Name", "Applicant Concentration","Current Concentration"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					PM_query="SELECT c.UIN, c.Firstname, c.LastName, isnull( c.Concentration,'') As CurrentConcentration,  isnull( a.Concentration,'') As ApplicantConcentration FROM CurrentStudents c left outer join Applicants a on c.UIN = a.UIN and a.Application_Status is not null and a.Application_Status != 'IN-Incomplete' and c.ProgramType <> '' and c.AdmitTerm <> '' and (c.Graduated != 'Y'  or c.Graduated is null) and c.Status not in ('LOA', 'TRANS','WDN') and c.Decision <> 'DF' and c.ProgramType = 'PM' Group By c.UIN, c.Concentration,a.Concentration, c.LastName, c.FirstName order by Lastname"
					rs.Open PM_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("ApplicantConcentration"),"|",",")
                    e = Replace(rs("CurrentConcentration"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_PM_Student = "Total number of students in PM : "&PM_rows
    pdf.ChapterBody(Total_PM_Student)
    pdf.Ln(1)

'////// TR-FT Students ////////
pdf.OrangeTitle("Program Option TR-FT")

set rs=Server.CreateObject("ADODB.recordset")
TRFT_query="SELECT Count(distinct UIN) TRFT_students FROM CurrentStudents where ProgramType='TR-FT' and (Graduated != 'Y'  or Graduated is null)"
rs.Open TRFT_query,conn

TRFT_rows = rs("TRFT_students")
TRFT_cols = 5
Dim TRFT_col(5)
TRFT_col(1) = "Banner # "
TRFT_col(2) = "First Name"
TRFT_col(3) = "Last Name"
TRFT_col(4) = "Current Concentration"
TRFT_col(5) = "Applicant Concentration"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,40,40
'pdf.SetAligns "C", "L", "L", "C","C"
pdf.SetFont "Arial","B",10
pdf.Row "UIN","First name","Last Name", "Applicant Concentration","Current Concentration"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					TRFT_query="SELECT c.UIN, c.Firstname, c.LastName, isnull( c.Concentration,'') As CurrentConcentration,  isnull( a.Concentration,'') As ApplicantConcentration FROM CurrentStudents c left outer join Applicants a on c.UIN = a.UIN and a.Application_Status is not null and a.Application_Status != 'IN-Incomplete' and c.ProgramType <> '' and c.AdmitTerm <> '' and (c.Graduated != 'Y'  or c.Graduated is null) and c.Status not in ('LOA', 'TRANS','WDN') and c.Decision <> 'DF' and c.ProgramType = 'TR-FT' Group By c.UIN, c.Concentration,a.Concentration, c.LastName, c.FirstName order by Lastname"
					rs.Open TRFT_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("ApplicantConcentration"),"|",",")
                    e = Replace(rs("CurrentConcentration"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_TRFT_Student = "Total number of students in TR-FT : "&TRFT_rows
    pdf.ChapterBody(Total_TRFT_Student)
    pdf.Ln(1)

'////// TR-PM Students ////////
pdf.OrangeTitle("Program Option TR-PM")

set rs=Server.CreateObject("ADODB.recordset")
TRPM_query="SELECT Count(distinct UIN) TRPM_students FROM CurrentStudents where ProgramType='TR-PM' and (Graduated != 'Y'  or Graduated is null)"
rs.Open TRPM_query,conn

TRPM_rows = rs("TRPM_students")
TRPM_cols = 5
Dim TRPM_col(5)
TRPM_col(1) = "Banner # "
TRPM_col(2) = "First Name"
TRPM_col(3) = "Last Name"
TRPM_col(4) = "Current Concentration"
TRPM_col(5) = "Applicant Concentration"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,40,40
'pdf.SetAligns "C", "L", "L", "C","C"
pdf.SetFont "Arial","B",10
pdf.Row "UIN","First name","Last Name", "Applicant Concentration","Current Concentration"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					TRPM_query="SELECT c.UIN, c.Firstname, c.LastName, isnull( c.Concentration,'') As CurrentConcentration,  isnull( a.Concentration,'') As ApplicantConcentration FROM CurrentStudents c left outer join Applicants a on c.UIN = a.UIN and a.Application_Status is not null and a.Application_Status != 'IN-Incomplete' and c.ProgramType <> '' and c.AdmitTerm <> '' and (c.Graduated != 'Y'  or c.Graduated is null) and c.Status not in ('LOA', 'TRANS','WDN') and c.Decision <> 'DF' and c.ProgramType = 'TR-PM' Group By c.UIN, c.Concentration,a.Concentration, c.LastName, c.FirstName order by Lastname"
					rs.Open TRPM_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("ApplicantConcentration"),"|",",")
                    e = Replace(rs("CurrentConcentration"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_TRPM_Student = "Total number of students in TRPM : "&TRPM_rows
    pdf.ChapterBody(Total_TRPM_Student)
    pdf.Ln(1)



pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
