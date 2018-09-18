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

pdf.ChapterTitle2("  Jane Addams College of Social Work Current Students All Advisors  "&LastUpdatedDt&" "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(3)

'////// All Advisor report ////////

    '//////// Blank ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_blank_query="SELECT Count(distinct UIN) blank_cnf_students FROM CurrentStudents where Advisor = '' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_blank_query,conn
    If rs("blank_cnf_students") <> 0 Then
    pdf.OrangeTitle("")
'pdf.FancyTable()

'//////// Students ////////////

blank_cnf_rows = rs("blank_cnf_students")
blank_cnf_cols = 6
Dim blank_cnf_col(6)
blank_cnf_col(1) = "Degree Program"
blank_cnf_col(2) = "Last Name"
blank_cnf_col(3) = "First Name"
blank_cnf_col(4) = "Email"
blank_cnf_col(5) = "Concentration"
blank_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					blankcnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = '' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open blankcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  

set rs=Server.CreateObject("ADODB.recordset")
totalqueryblank="SELECT Count(distinct UIN) totalblank_students FROM CurrentStudents where Advisor = '' and (Graduated != 'Y'  or Graduated is null) "
rs.Open totalqueryblank,conn
 Total_StudentBlank = "Total number of students with no Advisor: "&rs("totalblank_students")
    pdf.ChapterBody(Total_StudentBlank)
pdf.Ln(10)
    rs.close
'//////// All ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_All_query="SELECT Count(distinct UIN) All_cnf_students FROM CurrentStudents where Advisor = 'All' and (Graduated != 'Y'  or Graduated is null) "
rs.Open cnf_students_All_query,conn
    If rs("All_cnf_students") <> 0 Then
    pdf.OrangeTitle("All")
'pdf.FancyTable()

'//////// Students ////////////

All_cnf_rows = rs("All_cnf_students")
All_cnf_cols = 6
Dim All_cnf_col(6)
All_cnf_col(1) = "Degree Program"
All_cnf_col(2) = "Last Name"
All_cnf_col(3) = "First Name"
All_cnf_col(4) = "Email"
All_cnf_col(5) = "Concentration"
All_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Allcnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'All' and (Graduated != 'Y'  or Graduated is null)  order by LastName"
					rs.Open Allcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  

set rs=Server.CreateObject("ADODB.recordset")
totalqueryAll="SELECT Count(distinct UIN) totalAll_students FROM CurrentStudents where Advisor = 'All' and (Graduated != 'Y'  or Graduated is null) "
rs.Open totalqueryAll,conn
 Total_StudentAll = "Total number of students for Advisor All: "&rs("totalAll_students")
    pdf.ChapterBody(Total_StudentAll)
pdf.Ln(10)
    rs.close

'//////// Aaron Gottlieb ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_AaronGottlieb_query="SELECT Count(distinct UIN) AaronGottlieb_cnf_students FROM CurrentStudents where Advisor = 'Aaron Gottlieb' and (Graduated != 'Y'  or Graduated is null) "
rs.Open cnf_students_AaronGottlieb_query,conn
    If rs("AaronGottlieb_cnf_students") <> 0 Then
    pdf.OrangeTitle("Aaron Gottlieb")
'pdf.FancyTable()

'//////// Students ////////////

AaronGottlieb_cnf_rows = rs("AaronGottlieb_cnf_students")
AaronGottlieb_cnf_cols = 6
Dim AaronGottlieb_cnf_col(6)
AaronGottlieb_cnf_col(1) = "Degree Program"
AaronGottlieb_cnf_col(2) = "Last Name"
AaronGottlieb_cnf_col(3) = "First Name"
AaronGottlieb_cnf_col(4) = "Email"
AaronGottlieb_cnf_col(5) = "Concentration"
AaronGottlieb_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					AaronGottliebcnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Aaron Gottlieb' and (Graduated != 'Y'  or Graduated is null)  order by LastName"
					rs.Open AaronGottliebcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  

set rs=Server.CreateObject("ADODB.recordset")
totalqueryAaronGottlieb="SELECT Count(distinct UIN) totalAaronGottlieb_students FROM CurrentStudents where Advisor = 'Aaron Gottlieb' and (Graduated != 'Y'  or Graduated is null) "
rs.Open totalqueryAaronGottlieb,conn
 Total_StudentAaronGottlieb = "Total number of students for Advisor Aaron Gottlieb: "&rs("totalAaronGottlieb_students")
    pdf.ChapterBody(Total_StudentAaronGottlieb)
pdf.Ln(10)
    rs.close

'//////// Bonecutter ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Bonecutter_query="SELECT Count(distinct UIN) Bonecutter_cnf_students FROM CurrentStudents where Advisor = 'Bonecutter' and (Graduated != 'Y'  or Graduated is null) "
rs.Open cnf_students_Bonecutter_query,conn
    If rs("Bonecutter_cnf_students") <> 0 Then
    pdf.OrangeTitle("Bonecutter")
'pdf.FancyTable()

'//////// Students ////////////

Bonecutter_cnf_rows = rs("Bonecutter_cnf_students")
Bonecutter_cnf_cols = 6
Dim Bonecutter_cnf_col(6)
Bonecutter_cnf_col(1) = "Degree Program"
Bonecutter_cnf_col(2) = "Last Name"
Bonecutter_cnf_col(3) = "First Name"
Bonecutter_cnf_col(4) = "Email"
Bonecutter_cnf_col(5) = "Concentration"
Bonecutter_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Bonecuttercnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Bonecutter' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Bonecuttercnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If                        

    set rs=Server.CreateObject("ADODB.recordset")
totalqueryBonecutter="SELECT Count(distinct UIN) totalBonecutter_students FROM CurrentStudents where Advisor = 'Bonecutter' and (Graduated != 'Y'  or Graduated is null) "
rs.Open totalqueryBonecutter,conn
 Total_StudentBonecutter = "Total number of students for Advisor Bonecutter: "&rs("totalBonecutter_students")
    pdf.ChapterBody(Total_StudentBonecutter)
pdf.Ln(10)
    rs.close

    '//////// Branden McLeod ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_BrandenMcLeod_query="SELECT Count(distinct UIN) BrandenMcLeod_cnf_students FROM CurrentStudents where Advisor = 'Branden McLeod' and (Graduated != 'Y'  or Graduated is null) "
rs.Open cnf_students_BrandenMcLeod_query,conn
    If rs("BrandenMcLeod_cnf_students") <> 0 Then
    pdf.OrangeTitle("Branden McLeod")
'pdf.FancyTable()

'//////// Students ////////////

BrandenMcLeod_cnf_rows = rs("BrandenMcLeod_cnf_students")
BrandenMcLeod_cnf_cols = 6
Dim BrandenMcLeod_cnf_col(6)
BrandenMcLeod_cnf_col(1) = "Degree Program"
BrandenMcLeod_cnf_col(2) = "Last Name"
BrandenMcLeod_cnf_col(3) = "First Name"
BrandenMcLeod_cnf_col(4) = "Email"
BrandenMcLeod_cnf_col(5) = "Concentration"
BrandenMcLeod_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					BrandenMcLeodcnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Branden McLeod' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open BrandenMcLeodcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If                        

    set rs=Server.CreateObject("ADODB.recordset")
totalqueryBrandenMcLeod="SELECT Count(distinct UIN) totalBrandenMcLeod_students FROM CurrentStudents where Advisor = 'Branden McLeod' and (Graduated != 'Y'  or Graduated is null) "
rs.Open totalqueryBrandenMcLeod,conn
 Total_StudentBrandenMcLeod = "Total number of students for Advisor Branden McLeod: "&rs("totalBrandenMcLeod_students")
    pdf.ChapterBody(Total_StudentBrandenMcLeod)
pdf.Ln(10)
    rs.close

    '//////// Johnson Butterfield ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Butterfield_query="SELECT Count(distinct UIN) Butterfield_cnf_students FROM CurrentStudents where Advisor = 'Butterfield' and (Graduated != 'Y'  or Graduated is null) "
rs.Open cnf_students_Butterfield_query,conn
    If rs("Butterfield_cnf_students") <> 0 Then
    pdf.OrangeTitle("Butterfield")
'pdf.FancyTable()

'//////// Students ////////////

Butterfield_cnf_rows = rs("Butterfield_cnf_students")
Butterfield_cnf_cols = 6
Dim Butterfield_cnf_col(6)
Butterfield_cnf_col(1) = "Degree Program"
Butterfield_cnf_col(2) = "Last Name"
Butterfield_cnf_col(3) = "First Name"
Butterfield_cnf_col(4) = "Email"
Butterfield_cnf_col(5) = "Concentration"
Butterfield_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Butterfieldcnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Butterfield' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Butterfieldcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If            

        set rs=Server.CreateObject("ADODB.recordset")
totalqueryButterfield ="SELECT Count(distinct UIN) totalButterfield_students FROM CurrentStudents where Advisor = 'Butterfield' and (Graduated != 'Y'  or Graduated is null) "
rs.Open totalqueryButterfield,conn
 Total_StudentButterfield = "Total number of students for Advisor Butterfield: "&rs("totalButterfield_students")
    pdf.ChapterBody(Total_StudentButterfield)
pdf.Ln(10)
    rs.close
    '//////// Dettlaff ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Dettlaff_query="SELECT Count(distinct UIN) Dettlaff_cnf_students FROM CurrentStudents where Advisor = 'Dettlaff' and (Graduated != 'Y'  or Graduated is null) "
rs.Open cnf_students_Dettlaff_query,conn
    If rs("Dettlaff_cnf_students") <> 0 Then
    pdf.OrangeTitle("Dettlaff")
'pdf.FancyTable()

'//////// Students ////////////

Dettlaff_cnf_rows = rs("Dettlaff_cnf_students")
Dettlaff_cnf_cols = 6
Dim Dettlaff_cnf_col(6)
Dettlaff_cnf_col(1) = "Degree Program"
Dettlaff_cnf_col(2) = "Last Name"
Dettlaff_cnf_col(3) = "First Name"
Dettlaff_cnf_col(4) = "Email"
Dettlaff_cnf_col(5) = "Concentration"
Dettlaff_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Dettlaffcnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Dettlaff' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Dettlaffcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If                       

    set rs=Server.CreateObject("ADODB.recordset")
totalqueryDettlaff="SELECT Count(distinct UIN) totalDettlaff_students FROM CurrentStudents where Advisor = 'Dettlaff' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryDettlaff,conn
 Total_StudentDettlaff = "Total number of students for Advisor Dettlaff: "&rs("totalDettlaff_students")
    pdf.ChapterBody(Total_StudentDettlaff)
pdf.Ln(10)
    rs.close
'//////// Doyle ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Doyle_query="SELECT Count(distinct UIN) Doyle_cnf_students FROM CurrentStudents where Advisor = 'Doyle' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_Doyle_query,conn
    If rs("Doyle_cnf_students") <> 0 Then
    pdf.OrangeTitle("Doyle")
'pdf.FancyTable()

'//////// Students ////////////

Doyle_cnf_rows = rs("Doyle_cnf_students")
Doyle_cnf_cols = 6
Dim Doyle_cnf_col(6)
Doyle_cnf_col(1) = "Degree Program"
Doyle_cnf_col(2) = "Last Name"
Doyle_cnf_col(3) = "First Name"
Doyle_cnf_col(4) = "Email"
Doyle_cnf_col(5) = "Concentration"
Doyle_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Doylecnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Doyle' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Doylecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  

    set rs=Server.CreateObject("ADODB.recordset")
totalqueryDoyle="SELECT Count(distinct UIN) totalDoyle_students FROM CurrentStudents where Advisor = 'Doyle' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryDoyle,conn
 Total_StudentDoyle = "Total number of students for Advisor Doyle: "&rs("totalDoyle_students")
    pdf.ChapterBody(Total_StudentDoyle)
pdf.Ln(10)
    rs.close
    '//////// Gaston ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Gaston_query="SELECT Count(distinct UIN) Gaston_cnf_students FROM CurrentStudents where Advisor = 'Gaston' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_Gaston_query,conn
    If rs("Gaston_cnf_students") <> 0 Then
    pdf.OrangeTitle("Gaston")
'pdf.FancyTable()

'//////// Students ////////////

Gaston_cnf_rows = rs("Gaston_cnf_students")
Gaston_cnf_cols = 6
Dim Gaston_cnf_col(6)
Gaston_cnf_col(1) = "Degree Program"
Gaston_cnf_col(2) = "Last Name"
Gaston_cnf_col(3) = "First Name"
Gaston_cnf_col(4) = "Email"
Gaston_cnf_col(5) = "Concentration"
Gaston_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Gastoncnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Gaston' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Gastoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If                       

set rs=Server.CreateObject("ADODB.recordset")
totalqueryGaston="SELECT Count(distinct UIN) totalGaston_students FROM CurrentStudents where Advisor = 'Gaston' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryGaston,conn
 Total_StudentGaston = "Total number of students for Advisor Gaston: "&rs("totalGaston_students")
    pdf.ChapterBody(Total_StudentGaston)
pdf.Ln(10)
    rs.close
'//////// Gleeson ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Gleeson_query="SELECT Count(distinct UIN) Gleeson_cnf_students FROM CurrentStudents where Advisor = 'Gleeson' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_Gleeson_query,conn
    If rs("Gleeson_cnf_students") <> 0 Then
    pdf.OrangeTitle("Gleeson")
'pdf.FancyTable()

'//////// Students ////////////

Gleeson_cnf_rows = rs("Gleeson_cnf_students")
Gleeson_cnf_cols = 6
Dim Gleeson_cnf_col(6)
Gleeson_cnf_col(1) = "Degree Program"
Gleeson_cnf_col(2) = "Last Name"
Gleeson_cnf_col(3) = "First Name"
Gleeson_cnf_col(4) = "Email"
Gleeson_cnf_col(5) = "Concentration"
Gleeson_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Gleesoncnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Gleeson' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Gleesoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If           

    set rs=Server.CreateObject("ADODB.recordset")
totalqueryGleeson="SELECT Count(distinct UIN) totalGleeson_students FROM CurrentStudents where Advisor = 'Gleeson' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryGleeson,conn
 Total_StudentGleeson = "Total number of students for Advisor Gleeson: "&rs("totalGleeson_students")
    pdf.ChapterBody(Total_StudentGleeson)
pdf.Ln(10)
    rs.close
    '//////// Hairston ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Hairston_query="SELECT Count(distinct UIN) Hairston_cnf_students FROM CurrentStudents where Advisor = 'Hairston' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_Hairston_query,conn
    If rs("Hairston_cnf_students") <> 0 Then
    pdf.OrangeTitle("Hairston")
'pdf.FancyTable()

'//////// Students ////////////

Hairston_cnf_rows = rs("Hairston_cnf_students")
Hairston_cnf_cols = 6
Dim Hairston_cnf_col(6)
Hairston_cnf_col(1) = "Degree Program"
Hairston_cnf_col(2) = "Last Name"
Hairston_cnf_col(3) = "First Name"
Hairston_cnf_col(4) = "Email"
Hairston_cnf_col(5) = "Concentration"
Hairston_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Hairstoncnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Hairston' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Hairstoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  

    set rs=Server.CreateObject("ADODB.recordset")
totalqueryHairston="SELECT Count(distinct UIN) totalHairston_students FROM CurrentStudents where Advisor = 'Hairston' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryHairston,conn
 Total_StudentHairston = "Total number of students for Advisor Hairston: "&rs("totalHairston_students")
    pdf.ChapterBody(Total_StudentHairston)
pdf.Ln(10)
    rs.close
'//////// Hsieh ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Hsieh_query="SELECT Count(distinct UIN) Hsieh_cnf_students FROM CurrentStudents where Advisor = 'Hsieh' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_Hsieh_query,conn
    If rs("Hsieh_cnf_students") <> 0 Then
    pdf.OrangeTitle("Hsieh")
'pdf.FancyTable()

'//////// Students ////////////

Hsieh_cnf_rows = rs("Hsieh_cnf_students")
Hsieh_cnf_cols = 6
Dim Hsieh_cnf_col(6)
Hsieh_cnf_col(1) = "Degree Program"
Hsieh_cnf_col(2) = "Last Name"
Hsieh_cnf_col(3) = "First Name"
Hsieh_cnf_col(4) = "Email"
Hsieh_cnf_col(5) = "Concentration"
Hsieh_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Hsiehcnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Hsieh' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Hsiehcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If            

    set rs=Server.CreateObject("ADODB.recordset")
totalqueryHsieh="SELECT Count(distinct UIN) totalHsieh_students FROM CurrentStudents where Advisor = 'Hsieh' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryHsieh,conn
 Total_StudentHsieh = "Total number of students for Advisor Hsieh: "&rs("totalHsieh_students")
    pdf.ChapterBody(Total_StudentHsieh)
pdf.Ln(10)
    rs.close

    
'//////// Jack Lu ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_JackLu_query="SELECT Count(distinct UIN) JackLu_cnf_students FROM CurrentStudents where Advisor = 'Jack Lu' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_JackLu_query,conn
    If rs("JackLu_cnf_students") <> 0 Then
    pdf.OrangeTitle("Jack Lu")
'pdf.FancyTable()

'//////// Students ////////////

JackLu_cnf_rows = rs("JackLu_cnf_students")
JackLu_cnf_cols = 6
Dim JackLu_cnf_col(6)
JackLu_cnf_col(1) = "Degree Program"
JackLu_cnf_col(2) = "Last Name"
JackLu_cnf_col(3) = "First Name"
JackLu_cnf_col(4) = "Email"
JackLu_cnf_col(5) = "Concentration"
JackLu_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					JackLucnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Jack Lu' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open JackLucnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
set rs=Server.CreateObject("ADODB.recordset")
totalqueryJackLu="SELECT Count(distinct UIN) totalJackLu_students FROM CurrentStudents where Advisor = 'Jack Lu' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryJackLu,conn
 Total_StudentJackLu = "Total number of students for Advisor Jack Lu: "&rs("totalJackLu_students")
    pdf.ChapterBody(Total_StudentJackLu)
pdf.Ln(10)
    rs.close

'//////// A. Johnson ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Johnson_query="SELECT Count(distinct UIN) Johnson_cnf_students FROM CurrentStudents where Advisor = 'Johnson' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_Johnson_query,conn
    If rs("Johnson_cnf_students") <> 0 Then
    pdf.OrangeTitle("Johnson")
'pdf.FancyTable()

'//////// Students ////////////

Johnson_cnf_rows = rs("Johnson_cnf_students")
Johnson_cnf_cols = 6
Dim Johnson_cnf_col(6)
Johnson_cnf_col(1) = "Degree Program"
Johnson_cnf_col(2) = "Last Name"
Johnson_cnf_col(3) = "First Name"
Johnson_cnf_col(4) = "Email"
Johnson_cnf_col(5) = "Concentration"
Johnson_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Johnsoncnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Johnson' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Johnsoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
set rs=Server.CreateObject("ADODB.recordset")
totalqueryJohnson="SELECT Count(distinct UIN) totalJohnson_students FROM CurrentStudents where Advisor = 'Johnson' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryJohnson,conn
 Total_StudentJohnson = "Total number of students for Advisor Johnson: "&rs("totalJohnson_students")
    pdf.ChapterBody(Total_StudentJohnson)
pdf.Ln(10)
    rs.close

'//////// Karen D'Angelo ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Karen_query="SELECT Count(distinct UIN) Karen_cnf_students FROM CurrentStudents where Advisor = 'Karen D''Angelo' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_Karen_query,conn
    If rs("Karen_cnf_students") <> 0 Then
    pdf.OrangeTitle("Karen D'Angelo")
'pdf.FancyTable()

'//////// Students ////////////

Karen_cnf_rows = rs("Karen_cnf_students")
Karen_cnf_cols = 6
Dim Karen_cnf_col(6)
Karen_cnf_col(1) = "Degree Program"
Karen_cnf_col(2) = "Last Name"
Karen_cnf_col(3) = "First Name"
Karen_cnf_col(4) = "Email"
Karen_cnf_col(5) = "Concentration"
Karen_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Karencnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Karen D''Angelo' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Karencnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
set rs=Server.CreateObject("ADODB.recordset")
totalqueryKaren="SELECT Count(distinct UIN) totalKaren_students FROM CurrentStudents where Advisor = 'Karen D''Angelo' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryKaren,conn
 Total_StudentKaren = "Total number of students for Advisor Karen D'Angelo: "&rs("totalKaren_students")
    pdf.ChapterBody(Total_StudentKaren)
pdf.Ln(10)
    rs.close

'//////// Leathers ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Leathers_query="SELECT Count(distinct UIN) Leathers_cnf_students FROM CurrentStudents where Advisor = 'Leathers' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_Leathers_query,conn
    If rs("Leathers_cnf_students") <> 0 Then
    pdf.OrangeTitle("Leathers")
'pdf.FancyTable()

'//////// Students ////////////

Leathers_cnf_rows = rs("Leathers_cnf_students")
Leathers_cnf_cols = 6
Dim Leathers_cnf_col(6)
Leathers_cnf_col(1) = "Degree Program"
Leathers_cnf_col(2) = "Last Name"
Leathers_cnf_col(3) = "First Name"
Leathers_cnf_col(4) = "Email"
Leathers_cnf_col(5) = "Concentration"
Leathers_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Leatherscnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Leathers' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Leatherscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If            

        set rs=Server.CreateObject("ADODB.recordset")
totalqueryLeathers ="SELECT Count(distinct UIN) totalLeathers_students FROM CurrentStudents where Advisor = 'Leathers' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryLeathers,conn
 Total_StudentLeathers = "Total number of students for Advisor Leathers: "&rs("totalLeathers_students")
    pdf.ChapterBody(Total_StudentLeathers)
pdf.Ln(10)
    rs.close
    '//////// McCoy ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_McCoy_query="SELECT Count(distinct UIN) McCoy_cnf_students FROM CurrentStudents where Advisor = 'McCoy' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_McCoy_query,conn
    If rs("McCoy_cnf_students") <> 0 Then
    pdf.OrangeTitle("McCoy")
'pdf.FancyTable()

'//////// Students ////////////

McCoy_cnf_rows = rs("McCoy_cnf_students")
McCoy_cnf_cols = 6
Dim McCoy_cnf_col(6)
McCoy_cnf_col(1) = "Degree Program"
McCoy_cnf_col(2) = "Last Name"
McCoy_cnf_col(3) = "First Name"
McCoy_cnf_col(4) = "Email"
McCoy_cnf_col(5) = "Concentration"
McCoy_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					McCoycnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'McCoy' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open McCoycnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If            

 set rs=Server.CreateObject("ADODB.recordset")
totalqueryMcCoy ="SELECT Count(distinct UIN) totalMcCoy_students FROM CurrentStudents where Advisor = 'McCoy' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryMcCoy,conn
 Total_StudentMcCoy = "Total number of students for Advisor McCoy: "&rs("totalMcCoy_students")
    pdf.ChapterBody(Total_StudentMcCoy)
pdf.Ln(10)
    rs.close
'//////// McKay-Jackson ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_McKayJackson_query="SELECT Count(distinct UIN) McKayJackson_cnf_students FROM CurrentStudents where Advisor = 'McKay-Jackson' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_McKayJackson_query,conn
    If rs("McKayJackson_cnf_students") <> 0 Then
    pdf.OrangeTitle("McKay-Jackson")
'pdf.FancyTable()

'//////// Students ////////////

McKayJackson_cnf_rows = rs("McKayJackson_cnf_students")
McKayJackson_cnf_cols = 6
Dim McKayJackson_cnf_col(6)
McKayJackson_cnf_col(1) = "Degree Program"
McKayJackson_cnf_col(2) = "Last Name"
McKayJackson_cnf_col(3) = "First Name"
McKayJackson_cnf_col(4) = "Email"
McKayJackson_cnf_col(5) = "Concentration"
McKayJackson_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					McKayJacksoncnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'McKay-Jackson' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open McKayJacksoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If           

 set rs=Server.CreateObject("ADODB.recordset")
totalqueryMcKayJackson ="SELECT Count(distinct UIN) totalMcKayJackson_students FROM CurrentStudents where Advisor = 'McKay-Jackson' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryMcKayJackson,conn
 Total_StudentMcKayJackson = "Total number of students for Advisor McKay-Jackson: "&rs("totalMcKayJackson_students")
    pdf.ChapterBody(Total_StudentMcKayJackson)
pdf.Ln(10)
    rs.close
    '//////// Mitchell ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Mitchell_query="SELECT Count(distinct UIN) Mitchell_cnf_students FROM CurrentStudents where Advisor = 'Mitchell' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_Mitchell_query,conn
    If rs("Mitchell_cnf_students") <> 0 Then
    pdf.OrangeTitle("Mitchell")
'pdf.FancyTable()

'//////// Students ////////////

Mitchell_cnf_rows = rs("Mitchell_cnf_students")
Mitchell_cnf_cols = 6
Dim Mitchell_cnf_col(6)
Mitchell_cnf_col(1) = "Degree Program"
Mitchell_cnf_col(2) = "Last Name"
Mitchell_cnf_col(3) = "First Name"
Mitchell_cnf_col(4) = "Email"
Mitchell_cnf_col(5) = "Concentration"
Mitchell_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Mitchellcnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Mitchell' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Mitchellcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If           

     set rs=Server.CreateObject("ADODB.recordset")
totalqueryMitchell ="SELECT Count(distinct UIN) totalMitchell_students FROM CurrentStudents where Advisor = 'Mitchell' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryMitchell,conn
 Total_StudentMitchell = "Total number of students for Advisor Mitchell: "&rs("totalMitchell_students")
    pdf.ChapterBody(Total_StudentMitchell)
pdf.Ln(10)
    rs.close
'//////// Nebbitt ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Nebbitt_query="SELECT Count(distinct UIN) Nebbitt_cnf_students FROM CurrentStudents where Advisor = 'Nebbitt' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_Nebbitt_query,conn
    If rs("Nebbitt_cnf_students") <> 0 Then
    pdf.OrangeTitle("Nebbitt")
'pdf.FancyTable()

'//////// Students ////////////

Nebbitt_cnf_rows = rs("Nebbitt_cnf_students")
Nebbitt_cnf_cols = 6
Dim Nebbitt_cnf_col(6)
Nebbitt_cnf_col(1) = "Degree Program"
Nebbitt_cnf_col(2) = "Last Name"
Nebbitt_cnf_col(3) = "First Name"
Nebbitt_cnf_col(4) = "Email"
Nebbitt_cnf_col(5) = "Concentration"
Nebbitt_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Nebbittcnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Nebbitt' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Nebbittcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If           

     set rs=Server.CreateObject("ADODB.recordset")
totalqueryNebbitt ="SELECT Count(distinct UIN) totalNebbitt_students FROM CurrentStudents where Advisor = 'Nebbitt' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryNebbitt,conn
 Total_StudentNebbitt = "Total number of students for Advisor Nebbitt: "&rs("totalNebbitt_students")
    pdf.ChapterBody(Total_StudentNebbitt)
pdf.Ln(10)
    rs.close
    '//////// OBrien ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_OBrien_query="SELECT Count(distinct UIN) OBrien_cnf_students FROM CurrentStudents where Advisor = 'O''Brien' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_OBrien_query,conn
    If rs("OBrien_cnf_students") <> 0 Then
    pdf.OrangeTitle("O'Brien")
'pdf.FancyTable()

'//////// Students ////////////

OBrien_cnf_rows = rs("OBrien_cnf_students")
OBrien_cnf_cols = 6
Dim OBrien_cnf_col(6)
OBrien_cnf_col(1) = "Degree Program"
OBrien_cnf_col(2) = "Last Name"
OBrien_cnf_col(3) = "First Name"
OBrien_cnf_col(4) = "Email"
OBrien_cnf_col(5) = "Concentration"
OBrien_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					OBriencnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'O''Brien' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open OBriencnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If            

set rs=Server.CreateObject("ADODB.recordset")
totalqueryOBrien ="SELECT Count(distinct UIN) totalOBrien_students FROM CurrentStudents where Advisor = 'O''Brien' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryOBrien,conn
 Total_StudentOBrien = "Total number of students for Advisor O'Brien: "&rs("totalOBrien_students")
    pdf.ChapterBody(Total_StudentOBrien)
pdf.Ln(10)
    rs.close

    '//////// Robert Wilson ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_RobertWilson_query="SELECT Count(distinct UIN) RobertWilson_cnf_students FROM CurrentStudents where Advisor = 'Robert Wilson' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_RobertWilson_query,conn
    If rs("RobertWilson_cnf_students") <> 0 Then
    pdf.OrangeTitle("Robert Wilson")
'pdf.FancyTable()

'//////// Students ////////////

RobertWilson_cnf_rows = rs("RobertWilson_cnf_students")
RobertWilson_cnf_cols = 6
Dim RobertWilson_cnf_col(6)
RobertWilson_cnf_col(1) = "Degree Program"
RobertWilson_cnf_col(2) = "Last Name"
RobertWilson_cnf_col(3) = "First Name"
RobertWilson_cnf_col(4) = "Email"
RobertWilson_cnf_col(5) = "Concentration"
RobertWilson_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					RobertWilsoncnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Robert Wilson' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open RobertWilsoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If            

set rs=Server.CreateObject("ADODB.recordset")
totalqueryRobertWilson ="SELECT Count(distinct UIN) totalRobertWilson_students FROM CurrentStudents where Advisor = 'Robert Wilson' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryRobertWilson,conn
 Total_StudentRobertWilson = "Total number of students for Advisor Robert Wilson: "&rs("totalRobertWilson_students")
    pdf.ChapterBody(Total_StudentRobertWilson)
pdf.Ln(10)
    rs.close

'//////// Swartz ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Swarz_query="SELECT Count(distinct UIN) Swarz_cnf_students FROM CurrentStudents where Advisor = 'Swartz' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_Swarz_query,conn
    If rs("Swarz_cnf_students") <> 0 Then
    pdf.OrangeTitle("Swartz")
'pdf.FancyTable()

'//////// Students ////////////

Swarz_cnf_rows = rs("Swarz_cnf_students")
Swarz_cnf_cols = 6
Dim Swarz_cnf_col(6)
Swarz_cnf_col(1) = "Degree Program"
Swarz_cnf_col(2) = "Last Name"
Swarz_cnf_col(3) = "First Name"
Swarz_cnf_col(4) = "Email"
Swarz_cnf_col(5) = "Concentration"
Swarz_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Swarzcnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Swartz' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Swarzcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If            

set rs=Server.CreateObject("ADODB.recordset")
totalquerySwartz ="SELECT Count(distinct UIN) totalSwartz_students FROM CurrentStudents where Advisor = 'Swartz' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalquerySwartz,conn
 Total_StudentSwartz = "Total number of students for Advisor Swartz: "&rs("totalSwartz_students")
    pdf.ChapterBody(Total_StudentSwartz)
pdf.Ln(10)
    rs.close
    '//////// Watson ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Watson_query="SELECT Count(distinct UIN) Watson_cnf_students FROM CurrentStudents where Advisor = 'Watson' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_Watson_query,conn
    If rs("Watson_cnf_students") <> 0 Then
    pdf.OrangeTitle("Watson")
'pdf.FancyTable()

'//////// Students ////////////

Watson_cnf_rows = rs("Watson_cnf_students")
Watson_cnf_cols = 6
Dim Watson_cnf_col(6)
Watson_cnf_col(1) = "Degree Program"
Watson_cnf_col(2) = "Last Name"
Watson_cnf_col(3) = "First Name"
Watson_cnf_col(4) = "Email"
Watson_cnf_col(5) = "Concentration"
Watson_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Watsoncnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Watson' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open Watsoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If            

    set rs=Server.CreateObject("ADODB.recordset")
totalqueryWatson ="SELECT Count(distinct UIN) totalWatson_students FROM CurrentStudents where Advisor = 'Watson' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryWatson,conn
 Total_StudentWatson = "Total number of students for Advisor Watson: "&rs("totalWatson_students")
    pdf.ChapterBody(Total_StudentWatson)
pdf.Ln(10)
    rs.close
'//////// Wheeler-Brooks ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_WheelerBrooks_query="SELECT Count(distinct UIN) WheelerBrooks_cnf_students FROM CurrentStudents where Advisor = 'Wheeler-Brooks' and (Graduated != 'Y'  or Graduated is null)"
rs.Open cnf_students_WheelerBrooks_query,conn
    If rs("WheelerBrooks_cnf_students") <> 0 Then
    pdf.OrangeTitle("Wheeler-Brooks")
'pdf.FancyTable()

'//////// Students ////////////

WheelerBrooks_cnf_rows = rs("WheelerBrooks_cnf_students")
WheelerBrooks_cnf_cols = 6
Dim WheelerBrooks_cnf_col(6)
WheelerBrooks_cnf_col(1) = "Degree Program"
WheelerBrooks_cnf_col(2) = "Last Name"
WheelerBrooks_cnf_col(3) = "First Name"
WheelerBrooks_cnf_col(4) = "Email"
WheelerBrooks_cnf_col(5) = "Concentration"
WheelerBrooks_cnf_col(6) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,40,40,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					WheelerBrookscnf_query="SELECT DegreeProgram,LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where Advisor = 'Wheeler-Brooks' and (Graduated != 'Y'  or Graduated is null) order by LastName"
					rs.Open WheelerBrookscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("EMail"),"|",",")
                    e = Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If            

set rs=Server.CreateObject("ADODB.recordset")
totalqueryWheelerBrooks ="SELECT Count(distinct UIN) totalWheelerBrooks_students FROM CurrentStudents where Advisor = 'Wheeler-Brooks' and (Graduated != 'Y'  or Graduated is null)"
rs.Open totalqueryWheelerBrooks,conn
 Total_StudentWheelerBrooks = "Total number of students for Advisor Wheeler-Brooks: "&rs("totalWheelerBrooks_students")
    pdf.ChapterBody(Total_StudentWheelerBrooks)
pdf.Ln(10)
    rs.close



pdf.Ln(10)
pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
