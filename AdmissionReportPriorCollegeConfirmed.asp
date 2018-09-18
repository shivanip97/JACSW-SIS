<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Login_Check.asp"-->
<!--#include file="fpdf.asp"-->
<!--#include file="DBconn.asp"-->
<%

AdmitTerm=Request("term")
set rs=Server.CreateObject("ADODB.recordset")
termquery = "select distinct Admit_Term from Applicants where Term_CD like '"&AdmitTerm&"'"


Dim i,pdf
rs.Open termquery,conn
if rs.eof = false then 
Termsel=rs("Admit_Term")
 end if
LastUpdatedTime = Time()
LastUpdatedDt = date()

Set pdf=CreateJsObject("FPDF")
pdf.CreatePDF()
pdf.SetPath("fpdf/")
pdf.LoadExtension("table")
pdf.SetFont "Arial","",18
pdf.Open()
pdf.LoadModels("TestModels") 
pdf.AddPage()

pdf.ChapterTitle2("Report 5-Admissions Report -"&Termsel& "-Applied-Confirmed-Prior College "&LastUpdatedDt&" "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(10)

rs.close


'////// Undergraduate College Students ////////
pdf.OrangeTitle("Undergraduate College")
pdf.Ln(5)

    '//////// Albion College Students ////////////

set rs=Server.CreateObject("ADODB.recordset")
cnf_students_albion_query="SELECT Count(distinct UIN) albion_cnf_students FROM Applicants where UGCollege = 'Albion College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_albion_query,conn
    If rs("albion_cnf_students") <> 0 Then
    pdf.GreyTitle("Albion College")
'pdf.FancyTable()

'//////// Students ////////////

albion_cnf_rows = rs("albion_cnf_students")
albion_cnf_cols = 5
Dim albion_cnf_col(5)
albion_cnf_col(1) = "Banner # "
albion_cnf_col(2) = "First Name"
albion_cnf_col(3) = "Last Name"
albion_cnf_col(4) = "Applied"
albion_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					albioncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Albion College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open albioncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
    
    '//////// American Military University Students ////////////

set rs=Server.CreateObject("ADODB.recordset")
cnf_students_americanMilitary_query="SELECT Count(distinct UIN) americanMilitary_cnf_students FROM Applicants where UGCollege = 'American Military University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_americanMilitary_query,conn
    If rs("americanMilitary_cnf_students") <> 0 Then
    pdf.GreyTitle("American Military University")
'pdf.FancyTable()

'//////// Students ////////////

americanMilitary_cnf_rows = rs("americanMilitary_cnf_students")
americanMilitary_cnf_cols = 5
Dim americanMilitary_cnf_col(5)
americanMilitary_cnf_col(1) = "Banner # "
americanMilitary_cnf_col(2) = "First Name"
americanMilitary_cnf_col(3) = "Last Name"
americanMilitary_cnf_col(4) = "Applied"
americanMilitary_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					americanMilitarycnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'American Military University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open americanMilitarycnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
    
    '//////// American University Students ////////////

set rs=Server.CreateObject("ADODB.recordset")
cnf_students_american_query="SELECT Count(distinct UIN) american_cnf_students FROM Applicants where UGCollege = 'American University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_american_query,conn
    If rs("american_cnf_students") <> 0 Then
    pdf.GreyTitle("American University")
'pdf.FancyTable()

'//////// Students ////////////

american_cnf_rows = rs("american_cnf_students")
american_cnf_cols = 5
Dim american_cnf_col(5)
american_cnf_col(1) = "Banner # "
american_cnf_col(2) = "First Name"
american_cnf_col(3) = "Last Name"
american_cnf_col(4) = "Applied"
american_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					americancnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'American University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open americancnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  

'//////// Andrew Students ////////////
    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_andrew_query="SELECT Count(distinct UIN) andrew_cnf_students FROM Applicants where UGCollege = 'Andrews University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_andrew_query,conn
If rs("andrew_cnf_students") <> 0 Then
pdf.GreyTitle("Andrews University")
'pdf.FancyTable()


andrew_cnf_rows = rs("andrew_cnf_students")
andrew_cnf_cols = 5
Dim andrew_cnf_col(5)
andrew_cnf_col(1) = "Banner # "
andrew_cnf_col(2) = "First Name"
andrew_cnf_col(3) = "Last Name"
andrew_cnf_col(4) = "Applied"
andrew_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					andrewcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Andrews University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open andrewcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn

    '//////// Auburn University Students ////////////

set rs=Server.CreateObject("ADODB.recordset")
cnf_students_auburn_query="SELECT Count(distinct UIN) auburn_cnf_students FROM Applicants where UGCollege = 'Auburn University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_auburn_query,conn
    If rs("auburn_cnf_students") <> 0 Then
    pdf.GreyTitle("Auburn University")
'pdf.FancyTable()

'//////// Students ////////////

auburn_cnf_rows = rs("auburn_cnf_students")
auburn_cnf_cols = 5
Dim auburn_cnf_col(5)
auburn_cnf_col(1) = "Banner # "
auburn_cnf_col(2) = "First Name"
auburn_cnf_col(3) = "Last Name"
auburn_cnf_col(4) = "Applied"
auburn_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					auburncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Auburn University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open auburncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  

    '//////// Augustana College Students ////////////

set rs=Server.CreateObject("ADODB.recordset")
cnf_students_augustana_query="SELECT Count(distinct UIN) augustana_cnf_students FROM Applicants where UGCollege = 'Augustana College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_augustana_query,conn
    If rs("augustana_cnf_students") <> 0 Then
    pdf.GreyTitle("Augustana College")
'pdf.FancyTable()

'//////// Students ////////////

augustana_cnf_rows = rs("augustana_cnf_students")
augustana_cnf_cols = 5
Dim augustana_cnf_col(5)
augustana_cnf_col(1) = "Banner # "
augustana_cnf_col(2) = "First Name"
augustana_cnf_col(3) = "Last Name"
augustana_cnf_col(4) = "Applied"
augustana_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					augustanacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Augustana College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open augustanacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
 
'//////// Aurora Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_aurora_query="SELECT Count(distinct UIN) aurora_cnf_students FROM Applicants where UGCollege = 'Aurora University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_aurora_query,conn
    If rs("aurora_cnf_students") <> 0 Then
    pdf.GreyTitle("Aurora University")
'pdf.FancyTable()

'//////// Students ////////////

aurora_cnf_rows = rs("aurora_cnf_students")
aurora_cnf_cols = 5
Dim aurora_cnf_col(5)
aurora_cnf_col(1) = "Banner # "
aurora_cnf_col(2) = "First Name"
aurora_cnf_col(3) = "Last Name"
aurora_cnf_col(4) = "Applied"
aurora_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					auroracnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Aurora University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open auroracnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If 
  
'//////// Azusa Pacific University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_azusa_query="SELECT Count(distinct UIN) azusa_cnf_students FROM Applicants where UGCollege = 'Azusa Pacific University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_azusa_query,conn
    If rs("azusa_cnf_students") <> 0 Then
    pdf.GreyTitle("Azusa Pacific University")
'pdf.FancyTable()

'//////// Students ////////////

azusa_cnf_rows = rs("azusa_cnf_students")
azusa_cnf_cols = 5
Dim azusa_cnf_col(5)
azusa_cnf_col(1) = "Banner # "
azusa_cnf_col(2) = "First Name"
azusa_cnf_col(3) = "Last Name"
azusa_cnf_col(4) = "Applied"
azusa_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					azusacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Azusa Pacific University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open azusacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If 
     
'//////// Baylor University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_baylor_query="SELECT Count(distinct UIN) baylor_cnf_students FROM Applicants where UGCollege = 'Baylor University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_baylor_query,conn
    If rs("baylor_cnf_students") <> 0 Then
    pdf.GreyTitle("Baylor University")
'pdf.FancyTable()

'//////// Students ////////////

baylor_cnf_rows = rs("baylor_cnf_students")
baylor_cnf_cols = 5
Dim baylor_cnf_col(5)
baylor_cnf_col(1) = "Banner # "
baylor_cnf_col(2) = "First Name"
baylor_cnf_col(3) = "Last Name"
baylor_cnf_col(4) = "Applied"
baylor_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					baylorcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Baylor University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open baylorcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
    
'//////// Bellevue University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Bellevue_query="SELECT Count(distinct UIN) Bellevue_cnf_students FROM Applicants where UGCollege = 'Bellevue University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_Bellevue_query,conn
    If rs("Bellevue_cnf_students") <> 0 Then
    pdf.GreyTitle("Bellevue University")
'pdf.FancyTable()

'//////// Students ////////////

Bellevue_cnf_rows = rs("Bellevue_cnf_students")
Bellevue_cnf_cols = 5
Dim Bellevue_cnf_col(5)
Bellevue_cnf_col(1) = "Banner # "
Bellevue_cnf_col(2) = "First Name"
Bellevue_cnf_col(3) = "Last Name"
Bellevue_cnf_col(4) = "Applied"
Bellevue_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Bellevuecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Bellevue University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open Bellevuecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
    
'//////// Beloit College Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_beloit_query="SELECT Count(distinct UIN) beloit_cnf_students FROM Applicants where UGCollege = 'Beloit College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_beloit_query,conn
    If rs("beloit_cnf_students") <> 0 Then
    pdf.GreyTitle("Beloit College")
'pdf.FancyTable()

'//////// Students ////////////

beloit_cnf_rows = rs("beloit_cnf_students")
beloit_cnf_cols = 5
Dim beloit_cnf_col(5)
beloit_cnf_col(1) = "Banner # "
beloit_cnf_col(2) = "First Name"
beloit_cnf_col(3) = "Last Name"
beloit_cnf_col(4) = "Applied"
beloit_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					beloitcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Beloit College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open beloitcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                       
    '//////// Benedictine Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Benedictine_query="SELECT Count(distinct UIN) benedictine_cnf_students FROM Applicants where UGCollege = 'Benedictine University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_Benedictine_query,conn
    If rs("benedictine_cnf_students") <> 0 Then 
    pdf.GreyTitle("Benedictine University")
'pdf.FancyTable()

benedictine_cnf_rows = rs("benedictine_cnf_students")
benedictine_cnf_cols = 5
Dim benedictine_cnf_col(5)
benedictine_cnf_col(1) = "Banner # "
benedictine_cnf_col(2) = "First Name"
benedictine_cnf_col(3) = "Last Name"
benedictine_cnf_col(4) = "Applied"
benedictine_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					benedictinecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Benedictine University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open benedictinecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                       
    '//////// Bethel Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_bethel_query="SELECT Count(distinct UIN) bethel_cnf_students FROM Applicants where UGCollege = 'Bethel University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_bethel_query,conn
    If rs("bethel_cnf_students") <> 0 Then 
    pdf.GreyTitle("Bethel University")
'pdf.FancyTable()

bethel_cnf_rows = rs("bethel_cnf_students")
bethel_cnf_cols = 5
Dim bethel_cnf_col(5)
bethel_cnf_col(1) = "Banner # "
bethel_cnf_col(2) = "First Name"
bethel_cnf_col(3) = "Last Name"
bethel_cnf_col(4) = "Applied"
bethel_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					bethelcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Bethel University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open bethelcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

'//////// Bowling Green State Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_bowlinggreenstate_query="SELECT Count(distinct UIN) bowlinggreenstate_cnf_students FROM Applicants where UGCollege = 'Bowling Green State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_bowlinggreenstate_query,conn
    If rs("bowlinggreenstate_cnf_students") <> 0 Then
    pdf.GreyTitle("Bowling Green State University")
'pdf.FancyTable()



bowlinggreenstate_cnf_rows = rs("bowlinggreenstate_cnf_students")
bowlinggreenstate_cnf_cols = 5
Dim bowlinggreenstate_cnf_col(5)
bowlinggreenstate_cnf_col(1) = "Banner # "
bowlinggreenstate_cnf_col(2) = "First Name"
bowlinggreenstate_cnf_col(3) = "Last Name"
bowlinggreenstate_cnf_col(4) = "Applied"
bowlinggreenstate_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					bowlinggreenstatecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Bowling Green State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open bowlinggreenstatecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If                
'////////  Bradley Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_bradley_query="SELECT Count(distinct UIN) bradley_cnf_students FROM Applicants where UGCollege = 'Bradley University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_bradley_query,conn
If rs("bradley_cnf_students") <> 0 Then
pdf.GreyTitle("Bradley University")
'pdf.FancyTable()


bradley_cnf_rows = rs("bradley_cnf_students")
bradley_cnf_cols = 5
Dim bradley_cnf_col(5)
bradley_cnf_col(1) = "Banner # "
bradley_cnf_col(2) = "First Name"
bradley_cnf_col(3) = "Last Name"
bradley_cnf_col(4) = "Applied"
bradley_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					bradleycnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Bradley University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open bradleycnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
'//////// Brandeis Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_brandeis_query="SELECT Count(distinct UIN) brandeis_cnf_students FROM Applicants where UGCollege = 'Brandeis University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_brandeis_query,conn
    If rs("brandeis_cnf_students") <> 0 Then
    pdf.GreyTitle("Brandeis University")
'pdf.FancyTable()

'//////// Students ////////////

brandeis_cnf_rows = rs("brandeis_cnf_students")
brandeis_cnf_cols = 5
Dim brandeis_cnf_col(5)
brandeis_cnf_col(1) = "Banner # "
brandeis_cnf_col(2) = "First Name"
brandeis_cnf_col(3) = "Last Name"
brandeis_cnf_col(4) = "Applied"
brandeis_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					brandeiscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Brandeis University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open brandeiscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  

'//////// Brigham Young University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_brigham_query="SELECT Count(distinct UIN) brigham_cnf_students FROM Applicants where UGCollege = 'Brigham Young University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_brigham_query,conn
    If rs("brigham_cnf_students") <> 0 Then
    pdf.GreyTitle("Brigham Young University")
'pdf.FancyTable()

'//////// Students ////////////

brigham_cnf_rows = rs("brigham_cnf_students")
brigham_cnf_cols = 5
Dim brigham_cnf_col(5)
brigham_cnf_col(1) = "Banner # "
brigham_cnf_col(2) = "First Name"
brigham_cnf_col(3) = "Last Name"
brigham_cnf_col(4) = "Applied"
brigham_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					brighamcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Brigham Young University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open brighamcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                                 
    '//////// Buena Vista Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_buenavista_query="SELECT Count(distinct UIN) buenavista_cnf_students FROM Applicants where UGCollege = 'Buena Vista University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_buenavista_query,conn
    If rs("buenavista_cnf_students") <> 0 Then 
    pdf.GreyTitle("Buena Vista University")
'pdf.FancyTable()

buenavista_cnf_rows = rs("buenavista_cnf_students")
buenavista_cnf_cols = 5
Dim buenavista_cnf_col(5)
buenavista_cnf_col(1) = "Banner # "
buenavista_cnf_col(2) = "First Name"
buenavista_cnf_col(3) = "Last Name"
buenavista_cnf_col(4) = "Applied"
buenavista_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					buenavistacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Buena Vista University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open buenavistacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

     '//////// Butler University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_butler_query="SELECT Count(distinct UIN) butler_cnf_students FROM Applicants where UGCollege = 'Butler University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_butler_query,conn
    If rs("butler_cnf_students") <> 0 Then 
    pdf.GreyTitle("Butler University")
'pdf.FancyTable()

butler_cnf_rows = rs("butler_cnf_students")
butler_cnf_cols = 5
Dim butler_cnf_col(5)
butler_cnf_col(1) = "Banner # "
butler_cnf_col(2) = "First Name"
butler_cnf_col(3) = "Last Name"
butler_cnf_col(4) = "Applied"
butler_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					butlercnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Butler University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open butlercnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
  
     '//////// California State University Fullerton Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_fullerton_query="SELECT Count(distinct UIN) fullerton_cnf_students FROM Applicants where UGCollege = 'California State University Fullerton' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_fullerton_query,conn
    If rs("fullerton_cnf_students") <> 0 Then 
    pdf.GreyTitle("California State University Fullerton")
'pdf.FancyTable()

fullerton_cnf_rows = rs("fullerton_cnf_students")
fullerton_cnf_cols = 5
Dim fullerton_cnf_col(5)
fullerton_cnf_col(1) = "Banner # "
fullerton_cnf_col(2) = "First Name"
fullerton_cnf_col(3) = "Last Name"
fullerton_cnf_col(4) = "Applied"
fullerton_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					fullertoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'California State University Fullerton' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open fullertoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
            
  
     '//////// California State University Los Angeles Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_losangeles_query="SELECT Count(distinct UIN) losangeles_cnf_students FROM Applicants where UGCollege = 'California State University Los Angeles' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_losangeles_query,conn
    If rs("losangeles_cnf_students") <> 0 Then 
    pdf.GreyTitle("California State University Los Angeles")
'pdf.FancyTable()

losangeles_cnf_rows = rs("losangeles_cnf_students")
losangeles_cnf_cols = 5
Dim losangeles_cnf_col(5)
losangeles_cnf_col(1) = "Banner # "
losangeles_cnf_col(2) = "First Name"
losangeles_cnf_col(3) = "Last Name"
losangeles_cnf_col(4) = "Applied"
losangeles_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					losangelescnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'California State University Los Angeles' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open losangelescnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
            
     '//////// California State University San Bernardino Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_bernardino_query="SELECT Count(distinct UIN) bernardino_cnf_students FROM Applicants where UGCollege = 'California State University San Bernardino' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_bernardino_query,conn
    If rs("bernardino_cnf_students") <> 0 Then 
    pdf.GreyTitle("California State University San Bernardino")
'pdf.FancyTable()

bernardino_cnf_rows = rs("bernardino_cnf_students")
bernardino_cnf_cols = 5
Dim bernardino_cnf_col(5)
bernardino_cnf_col(1) = "Banner # "
bernardino_cnf_col(2) = "First Name"
bernardino_cnf_col(3) = "Last Name"
bernardino_cnf_col(4) = "Applied"
bernardino_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					bernardinocnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'California State University San Bernardino' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open bernardinocnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
            
    '//////// Calvin College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_calvin_query="SELECT Count(distinct UIN) calvin_cnf_students FROM Applicants where UGCollege = 'Calvin College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_calvin_query,conn
    If rs("calvin_cnf_students") <> 0 Then 
    pdf.GreyTitle("Calvin College")
'pdf.FancyTable()

calvin_cnf_rows = rs("calvin_cnf_students")
calvin_cnf_cols = 5
Dim calvin_cnf_col(5)
calvin_cnf_col(1) = "Banner # "
calvin_cnf_col(2) = "First Name"
calvin_cnf_col(3) = "Last Name"
calvin_cnf_col(4) = "Applied"
calvin_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					calvincnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Calvin College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open calvincnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
             
    '//////// Carleton College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_carleton_query="SELECT Count(distinct UIN) carleton_cnf_students FROM Applicants where UGCollege = 'Carleton College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_carleton_query,conn
    If rs("carleton_cnf_students") <> 0 Then 
    pdf.GreyTitle("Carleton College")
'pdf.FancyTable()

carleton_cnf_rows = rs("carleton_cnf_students")
carleton_cnf_cols = 5
Dim carleton_cnf_col(5)
carleton_cnf_col(1) = "Banner # "
carleton_cnf_col(2) = "First Name"
carleton_cnf_col(3) = "Last Name"
carleton_cnf_col(4) = "Applied"
carleton_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					carletoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Carleton College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open carletoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
             
    '//////// Carnegie Mellon University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_carnegie_query="SELECT Count(distinct UIN) carnegie_cnf_students FROM Applicants where UGCollege = 'Carnegie Mellon University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_carnegie_query,conn
    If rs("carnegie_cnf_students") <> 0 Then 
    pdf.GreyTitle("Carnegie Mellon University")
'pdf.FancyTable()

carnegie_cnf_rows = rs("carnegie_cnf_students")
carnegie_cnf_cols = 5
Dim carnegie_cnf_col(5)
carnegie_cnf_col(1) = "Banner # "
carnegie_cnf_col(2) = "First Name"
carnegie_cnf_col(3) = "Last Name"
carnegie_cnf_col(4) = "Applied"
carnegie_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					carnegiecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Carnegie Mellon University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open carnegiecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
             
    '//////// Carthage College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_carthage_query="SELECT Count(distinct UIN) carthage_cnf_students FROM Applicants where UGCollege = 'Carthage College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_carthage_query,conn
    If rs("carthage_cnf_students") <> 0 Then 
    pdf.GreyTitle("Carthage College")
'pdf.FancyTable()

carthage_cnf_rows = rs("carthage_cnf_students")
carthage_cnf_cols = 5
Dim carthage_cnf_col(5)
carthage_cnf_col(1) = "Banner # "
carthage_cnf_col(2) = "First Name"
carthage_cnf_col(3) = "Last Name"
carthage_cnf_col(4) = "Applied"
carthage_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					carthagecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Carthage College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open carthagecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
             
    '//////// Case Western Reserve University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_casewestern_query="SELECT Count(distinct UIN) casewestern_cnf_students FROM Applicants where UGCollege = 'Case Western Reserve University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_casewestern_query,conn
    If rs("casewestern_cnf_students") <> 0 Then 
    pdf.GreyTitle("Case Western Reserve University")
'pdf.FancyTable()

casewestern_cnf_rows = rs("casewestern_cnf_students")
casewestern_cnf_cols = 5
Dim casewestern_cnf_col(5)
casewestern_cnf_col(1) = "Banner # "
casewestern_cnf_col(2) = "First Name"
casewestern_cnf_col(3) = "Last Name"
casewestern_cnf_col(4) = "Applied"
casewestern_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					casewesterncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Case Western Reserve University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open casewesterncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
            
    '//////// Chicago State University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_chicagostate_query="SELECT Count(distinct UIN) chicagostate_cnf_students FROM Applicants where UGCollege = 'Chicago State University' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_chicagostate_query,conn
    If rs("chicagostate_cnf_students") <> 0 Then 
    pdf.GreyTitle("Chicago State University")
'pdf.FancyTable()

chicagostate_cnf_rows = rs("chicagostate_cnf_students")
chicagostate_cnf_cols = 5
Dim chicagostate_cnf_col(5)
chicagostate_cnf_col(1) = "Banner # "
chicagostate_cnf_col(2) = "First Name"
chicagostate_cnf_col(3) = "Last Name"
chicagostate_cnf_col(4) = "Applied"
chicagostate_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					chicagostatecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Chicago State University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open chicagostatecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
  
'//////// Chung-Ang University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_chungang_query="SELECT Count(distinct UIN) chungang_cnf_students FROM Applicants where UGCollege = 'Chung-Ang University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_chungang_query,conn
    If rs("chungang_cnf_students") <> 0 Then
    pdf.GreyTitle("Chung-Ang University")
'pdf.FancyTable()



chungang_cnf_rows = rs("chungang_cnf_students")
chungang_cnf_cols = 5
Dim chungang_cnf_col(5)
chungang_cnf_col(1) = "Banner # "
chungang_cnf_col(2) = "First Name"
chungang_cnf_col(3) = "Last Name"
chungang_cnf_col(4) = "Applied"
chungang_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					chungangcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Chung-Ang University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open chungangcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If                

 
  
'//////// Clark Atlanta Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_clarkatlanta_query="SELECT Count(distinct UIN) clarkatlanta_cnf_students FROM Applicants where UGCollege = 'Clark Atlanta University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_clarkatlanta_query,conn
    If rs("clarkatlanta_cnf_students") <> 0 Then
    pdf.GreyTitle("Clark Atlanta University")
'pdf.FancyTable()



clarkatlanta_cnf_rows = rs("clarkatlanta_cnf_students")
clarkatlanta_cnf_cols = 5
Dim clarkatlanta_cnf_col(5)
clarkatlanta_cnf_col(1) = "Banner # "
clarkatlanta_cnf_col(2) = "First Name"
clarkatlanta_cnf_col(3) = "Last Name"
clarkatlanta_cnf_col(4) = "Applied"
clarkatlanta_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					clarkatlantacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Clark Atlanta University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open clarkatlantacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If                

    '//////// Clarke Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_clarke_query="SELECT Count(distinct UIN) clarke_cnf_students FROM Applicants where UGCollege = 'Clarke University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_clarke_query,conn
    If rs("clarke_cnf_students") <> 0 Then
    pdf.GreyTitle("Clarke University")
'pdf.FancyTable()

'//////// Students ////////////

clarke_cnf_rows = rs("clarke_cnf_students")
clarke_cnf_cols = 5
Dim clarke_cnf_col(5)
clarke_cnf_col(1) = "Banner # "
clarke_cnf_col(2) = "First Name"
clarke_cnf_col(3) = "Last Name"
clarke_cnf_col(4) = "Applied"
clarke_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					clarkecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Clarke University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open clarkecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
 
    '//////// Colgate University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_colgate_query="SELECT Count(distinct UIN) colgate_cnf_students FROM Applicants where UGCollege = 'Colgate University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_colgate_query,conn
    If rs("colgate_cnf_students") <> 0 Then
    pdf.GreyTitle("Colgate University")
'pdf.FancyTable()

'//////// Students ////////////

colgate_cnf_rows = rs("colgate_cnf_students")
colgate_cnf_cols = 5
Dim colgate_cnf_col(5)
colgate_cnf_col(1) = "Banner # "
colgate_cnf_col(2) = "First Name"
colgate_cnf_col(3) = "Last Name"
colgate_cnf_col(4) = "Applied"
colgate_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					colgatecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Colgate University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open colgatecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
 
    '//////// College of St Scholastica Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_scholastica_query="SELECT Count(distinct UIN) scholastica_cnf_students FROM Applicants where UGCollege = 'College of St Scholastica' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_scholastica_query,conn
    If rs("scholastica_cnf_students") <> 0 Then
    pdf.GreyTitle("College of St Scholastica")
'pdf.FancyTable()

'//////// Students ////////////

scholastica_cnf_rows = rs("scholastica_cnf_students")
scholastica_cnf_cols = 5
Dim scholastica_cnf_col(5)
scholastica_cnf_col(1) = "Banner # "
scholastica_cnf_col(2) = "First Name"
scholastica_cnf_col(3) = "Last Name"
scholastica_cnf_col(4) = "Applied"
scholastica_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					scholasticacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'College of St Scholastica' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open scholasticacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                       
    '//////// Colorado State Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_coloradostate_query="SELECT Count(distinct UIN) coloradostate_cnf_students FROM Applicants where UGCollege = 'Colorado State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_coloradostate_query,conn
    If rs("coloradostate_cnf_students") <> 0 Then 
    pdf.GreyTitle("Colorado State University")
'pdf.FancyTable()

coloradostate_cnf_rows = rs("coloradostate_cnf_students")
coloradostate_cnf_cols = 5
Dim coloradostate_cnf_col(5)
coloradostate_cnf_col(1) = "Banner # "
coloradostate_cnf_col(2) = "First Name"
coloradostate_cnf_col(3) = "Last Name"
coloradostate_cnf_col(4) = "Applied"
coloradostate_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					coloradostatecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Colorado State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open coloradostatecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

'//////// Columbia College - Missouri Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_columbiamissouri_query="SELECT Count(distinct UIN) columbiamissouri_cnf_students FROM Applicants where UGCollege = 'Columbia College - Missouri' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_columbiamissouri_query,conn
    If rs("columbiamissouri_cnf_students") <> 0 Then
    pdf.GreyTitle("Columbia College - Missouri")
'pdf.FancyTable()



columbiamissouri_cnf_rows = rs("columbiamissouri_cnf_students")
columbiamissouri_cnf_cols = 5
Dim columbiamissouri_cnf_col(5)
columbiamissouri_cnf_col(1) = "Banner # "
columbiamissouri_cnf_col(2) = "First Name"
columbiamissouri_cnf_col(3) = "Last Name"
columbiamissouri_cnf_col(4) = "Applied"
columbiamissouri_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					columbiamissouricnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Columbia College - Missouri' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open columbiamissouricnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 
'//////// Columbia College Chicago Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_columbiacollege_query="SELECT Count(distinct UIN) columbiacollege_cnf_students FROM Applicants where UGCollege = 'Columbia College Chicago' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_columbiacollege_query,conn
    If rs("columbiacollege_cnf_students") <> 0 Then
    pdf.GreyTitle("Columbia College Chicago")
'pdf.FancyTable()



columbiacollege_cnf_rows = rs("columbiacollege_cnf_students")
columbiacollege_cnf_cols = 5
Dim columbiacollege_cnf_col(5)
columbiacollege_cnf_col(1) = "Banner # "
columbiacollege_cnf_col(2) = "First Name"
columbiacollege_cnf_col(3) = "Last Name"
columbiacollege_cnf_col(4) = "Applied"
columbiacollege_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					columbiacollegecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Columbia College Chicago' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open columbiacollegecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If                
'////////  Concordia Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_concordia_query="SELECT Count(distinct UIN) concordia_cnf_students FROM Applicants where UGCollege = 'Concordia University River Forest' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_concordia_query,conn
If rs("concordia_cnf_students") <> 0 Then
pdf.GreyTitle("Concordia University River Forest")
'pdf.FancyTable()


concordia_cnf_rows = rs("concordia_cnf_students")
concordia_cnf_cols = 5
Dim concordia_cnf_col(5)
concordia_cnf_col(1) = "Banner # "
concordia_cnf_col(2) = "First Name"
concordia_cnf_col(3) = "Last Name"
concordia_cnf_col(4) = "Applied"
concordia_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					concordiacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Concordia University River Forest' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open concordiacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn

'////////  Creighton University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_creighton_query="SELECT Count(distinct UIN) creighton_cnf_students FROM Applicants where UGCollege = 'Creighton University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_creighton_query,conn
If rs("creighton_cnf_students") <> 0 Then
pdf.GreyTitle("Creighton University")
'pdf.FancyTable()


creighton_cnf_rows = rs("creighton_cnf_students")
creighton_cnf_cols = 5
Dim creighton_cnf_col(5)
creighton_cnf_col(1) = "Banner # "
creighton_cnf_col(2) = "First Name"
creighton_cnf_col(3) = "Last Name"
creighton_cnf_col(4) = "Applied"
creighton_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					creightoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Creighton University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open creightoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
     
'//////// DePaul Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_depaul_query="SELECT Count(distinct UIN) depaul_cnf_students FROM Applicants where UGCollege = 'De Paul University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_depaul_query,conn
    If rs("depaul_cnf_students") <> 0 Then
    pdf.GreyTitle("De Paul University")
'pdf.FancyTable()

'//////// Students ////////////

depaul_cnf_rows = rs("depaul_cnf_students")
depaul_cnf_cols = 5
Dim depaul_cnf_col(5)
depaul_cnf_col(1) = "Banner # "
depaul_cnf_col(2) = "First Name"
depaul_cnf_col(3) = "Last Name"
depaul_cnf_col(4) = "Applied"
depaul_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					depaulcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'De Paul University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open depaulcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
    
'//////// Denison University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_denison_query="SELECT Count(distinct UIN) denison_cnf_students FROM Applicants where UGCollege = 'Denison University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_denison_query,conn
    If rs("denison_cnf_students") <> 0 Then
    pdf.GreyTitle("Denison University")
'pdf.FancyTable()

'//////// Students ////////////

denison_cnf_rows = rs("denison_cnf_students")
denison_cnf_cols = 5
Dim denison_cnf_col(5)
denison_cnf_col(1) = "Banner # "
denison_cnf_col(2) = "First Name"
denison_cnf_col(3) = "Last Name"
denison_cnf_col(4) = "Applied"
denison_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					denisoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Denison University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open denisoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
      
'//////// DePauw Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_depauw_query="SELECT Count(distinct UIN) depauw_cnf_students FROM Applicants where UGCollege = 'DePauw University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_depauw_query,conn
    If rs("depauw_cnf_students") <> 0 Then
    pdf.GreyTitle("DePauw University")
'pdf.FancyTable()

'//////// Students ////////////

depauw_cnf_rows = rs("depauw_cnf_students")
depauw_cnf_cols = 5
Dim depauw_cnf_col(5)
depauw_cnf_col(1) = "Banner # "
depauw_cnf_col(2) = "First Name"
depauw_cnf_col(3) = "Last Name"
depauw_cnf_col(4) = "Applied"
depauw_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					depauwcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'DePauw University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open depauwcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                        
                       
    '//////// Dominican Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_dominican_query="SELECT Count(distinct UIN) dominican_cnf_students FROM Applicants where UGCollege = 'Dominican University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_dominican_query,conn
    If rs("dominican_cnf_students") <> 0 Then 
    pdf.GreyTitle("Dominican University")
'pdf.FancyTable()

dominican_cnf_rows = rs("dominican_cnf_students")
dominican_cnf_cols = 5
Dim dominican_cnf_col(5)
dominican_cnf_col(1) = "Banner # "
dominican_cnf_col(2) = "First Name"
dominican_cnf_col(3) = "Last Name"
dominican_cnf_col(4) = "Applied"
dominican_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					dominicancnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Dominican University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open dominicancnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

'//////// East West University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_eastwest_query="SELECT Count(distinct UIN) eastwest_cnf_students FROM Applicants where UGCollege = 'East West University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_eastwest_query,conn
    If rs("eastwest_cnf_students") <> 0 Then
    pdf.GreyTitle("East West University")
'pdf.FancyTable()

eastwest_cnf_rows = rs("eastwest_cnf_students")
eastwest_cnf_cols = 5
Dim eastwest_cnf_col(5)
eastwest_cnf_col(1) = "Banner # "
eastwest_cnf_col(2) = "First Name"
eastwest_cnf_col(3) = "Last Name"
eastwest_cnf_col(4) = "Applied"
eastwest_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					eastwestcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'East West University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open eastwestcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If   


'//////// Eastern Illinois Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_easternillinois_query="SELECT Count(distinct UIN) easternillinois_cnf_students FROM Applicants where UGCollege = 'Eastern Illinois University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_easternillinois_query,conn
    If rs("easternillinois_cnf_students") <> 0 Then
    pdf.GreyTitle("Eastern Illinois University")
'pdf.FancyTable()



easternillinois_cnf_rows = rs("easternillinois_cnf_students")
easternillinois_cnf_cols = 5
Dim easternillinois_cnf_col(5)
easternillinois_cnf_col(1) = "Banner # "
easternillinois_cnf_col(2) = "First Name"
easternillinois_cnf_col(3) = "Last Name"
easternillinois_cnf_col(4) = "Applied"
easternillinois_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					easternillinoiscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Eastern Illinois University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open easternillinoiscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If   

    '//////// Elhurst Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_elhurst_query="SELECT Count(distinct UIN) elhurst_cnf_students FROM Applicants where UGCollege = 'Elhurst College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_elhurst_query,conn
    If rs("elhurst_cnf_students") <> 0 Then
    pdf.GreyTitle("Elhurst College")
'pdf.FancyTable()

'//////// Students ////////////

elhurst_cnf_rows = rs("elhurst_cnf_students")
elhurst_cnf_cols = 5
Dim elhurst_cnf_col(5)
elhurst_cnf_col(1) = "Banner # "
elhurst_cnf_col(2) = "First Name"
elhurst_cnf_col(3) = "Last Name"
elhurst_cnf_col(4) = "Applied"
elhurst_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					elhurstcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Elhurst College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open elhurstcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                     
    '//////// Elmhurst Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_elmhurst_query="SELECT Count(distinct UIN) elmhurst_cnf_students FROM Applicants where UGCollege = 'Elmhurst College'  and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_elmhurst_query,conn
    If rs("elmhurst_cnf_students") <> 0 Then
    pdf.GreyTitle("Elmhurst College")
'pdf.FancyTable()

'//////// Students ////////////

elmhurst_cnf_rows = rs("elmhurst_cnf_students")
elmhurst_cnf_cols = 5
Dim elmhurst_cnf_col(5)
elmhurst_cnf_col(1) = "Banner # "
elmhurst_cnf_col(2) = "First Name"
elmhurst_cnf_col(3) = "Last Name"
elmhurst_cnf_col(4) = "Applied"
elmhurst_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					elmhurstcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Elmhurst College' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open elmhurstcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If              
                     
    '//////// Elon University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_elon_query="SELECT Count(distinct UIN) elon_cnf_students FROM Applicants where UGCollege = 'Elon University'  and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_elon_query,conn
    If rs("elon_cnf_students") <> 0 Then
    pdf.GreyTitle("Elon University")
'pdf.FancyTable()

'//////// Students ////////////

elon_cnf_rows = rs("elon_cnf_students")
elon_cnf_cols = 5
Dim elon_cnf_col(5)
elon_cnf_col(1) = "Banner # "
elon_cnf_col(2) = "First Name"
elon_cnf_col(3) = "Last Name"
elon_cnf_col(4) = "Applied"
elon_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					eloncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Elon University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open eloncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If              
                      
    '//////// Emmanuel College Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_emmanuel_query="SELECT Count(distinct UIN) emmanuel_cnf_students FROM Applicants where UGCollege = 'Emmanuel College'  and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_emmanuel_query,conn
    If rs("emmanuel_cnf_students") <> 0 Then
    pdf.GreyTitle("Emmanuel College")
'pdf.FancyTable()

'//////// Students ////////////

emmanuel_cnf_rows = rs("emmanuel_cnf_students")
emmanuel_cnf_cols = 5
Dim emmanuel_cnf_col(5)
emmanuel_cnf_col(1) = "Banner # "
emmanuel_cnf_col(2) = "First Name"
emmanuel_cnf_col(3) = "Last Name"
emmanuel_cnf_col(4) = "Applied"
emmanuel_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					emmanuelcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Emmanuel College' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open emmanuelcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If              
                      
    '//////// Evangel University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_evangel_query="SELECT Count(distinct UIN) evangel_cnf_students FROM Applicants where UGCollege = 'Evangel University'  and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_evangel_query,conn
    If rs("evangel_cnf_students") <> 0 Then
    pdf.GreyTitle("Evangel University")
'pdf.FancyTable()

'//////// Students ////////////

evangel_cnf_rows = rs("evangel_cnf_students")
evangel_cnf_cols = 5
Dim evangel_cnf_col(5)
evangel_cnf_col(1) = "Banner # "
evangel_cnf_col(2) = "First Name"
evangel_cnf_col(3) = "Last Name"
evangel_cnf_col(4) = "Applied"
evangel_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					evangelcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Evangel University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open evangelcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If              
                      
    '//////// Florida State University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_floridastate_query="SELECT Count(distinct UIN) floridastate_cnf_students FROM Applicants where UGCollege = 'Florida State University'  and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_floridastate_query,conn
    If rs("floridastate_cnf_students") <> 0 Then
    pdf.GreyTitle("Florida State University")
'pdf.FancyTable()

'//////// Students ////////////

floridastate_cnf_rows = rs("floridastate_cnf_students")
floridastate_cnf_cols = 5
Dim floridastate_cnf_col(5)
floridastate_cnf_col(1) = "Banner # "
floridastate_cnf_col(2) = "First Name"
floridastate_cnf_col(3) = "Last Name"
floridastate_cnf_col(4) = "Applied"
floridastate_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					floridastatecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Florida State University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open floridastatecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If              
                      
    '//////// Fontbonne University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_fontbonne_query="SELECT Count(distinct UIN) fontbonne_cnf_students FROM Applicants where UGCollege = 'Fontbonne University'  and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_fontbonne_query,conn
    If rs("fontbonne_cnf_students") <> 0 Then
    pdf.GreyTitle("Fontbonne University")
'pdf.FancyTable()

'//////// Students ////////////

fontbonne_cnf_rows = rs("fontbonne_cnf_students")
fontbonne_cnf_cols = 5
Dim fontbonne_cnf_col(5)
fontbonne_cnf_col(1) = "Banner # "
fontbonne_cnf_col(2) = "First Name"
fontbonne_cnf_col(3) = "Last Name"
fontbonne_cnf_col(4) = "Applied"
fontbonne_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					fontbonnecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Fontbonne University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open fontbonnecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If              
        
        '//////// George Fox University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_george_query="SELECT Count(distinct UIN) george_cnf_students FROM Applicants where UGCollege = 'George Fox University' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_george_query,conn
    If rs("george_cnf_students") <> 0 Then 
    pdf.GreyTitle("George Fox University")
'pdf.FancyTable()

george_cnf_rows = rs("george_cnf_students")
george_cnf_cols = 5
Dim george_cnf_col(5)
george_cnf_col(1) = "Banner # "
george_cnf_col(2) = "First Name"
george_cnf_col(3) = "Last Name"
george_cnf_col(4) = "Applied"
george_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					georgecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'George Fox University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open georgecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
               
    '//////// Georgia Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_georgia_query="SELECT Count(distinct UIN) georgia_cnf_students FROM Applicants where UGCollege = 'Georgia State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_georgia_query,conn
    If rs("georgia_cnf_students") <> 0 Then 
    pdf.GreyTitle("Georgia State University")
'pdf.FancyTable()

georgia_cnf_rows = rs("georgia_cnf_students")
georgia_cnf_cols = 5
Dim georgia_cnf_col(5)
georgia_cnf_col(1) = "Banner # "
georgia_cnf_col(2) = "First Name"
georgia_cnf_col(3) = "Last Name"
georgia_cnf_col(4) = "Applied"
georgia_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					georgiacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Georgia State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open georgiacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
               
    '//////// Gordon College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_gordon_query="SELECT Count(distinct UIN) gordon_cnf_students FROM Applicants where UGCollege = 'Gordon College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_gordon_query,conn
    If rs("gordon_cnf_students") <> 0 Then 
    pdf.GreyTitle("Gordon College")
'pdf.FancyTable()

gordon_cnf_rows = rs("gordon_cnf_students")
gordon_cnf_cols = 5
Dim gordon_cnf_col(5)
gordon_cnf_col(1) = "Banner # "
gordon_cnf_col(2) = "First Name"
gordon_cnf_col(3) = "Last Name"
gordon_cnf_col(4) = "Applied"
gordon_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					gordoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Gordon College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open gordoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                
    '//////// Goshen College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_goshen_query="SELECT Count(distinct UIN) goshen_cnf_students FROM Applicants where UGCollege = 'Goshen College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_goshen_query,conn
    If rs("goshen_cnf_students") <> 0 Then 
    pdf.GreyTitle("Goshen College")
'pdf.FancyTable()

goshen_cnf_rows = rs("goshen_cnf_students")
goshen_cnf_cols = 5
Dim goshen_cnf_col(5)
goshen_cnf_col(1) = "Banner # "
goshen_cnf_col(2) = "First Name"
goshen_cnf_col(3) = "Last Name"
goshen_cnf_col(4) = "Applied"
goshen_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					goshencnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Goshen College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open goshencnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                
    '//////// Governors Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_governors_query="SELECT Count(distinct UIN) governors_cnf_students FROM Applicants where UGCollege = 'Governors State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_governors_query,conn
    If rs("governors_cnf_students") <> 0 Then 
    pdf.GreyTitle("Governors State University")
'pdf.FancyTable()

governors_cnf_rows = rs("governors_cnf_students")
governors_cnf_cols = 5
Dim governors_cnf_col(5)
governors_cnf_col(1) = "Banner # "
governors_cnf_col(2) = "First Name"
governors_cnf_col(3) = "Last Name"
governors_cnf_col(4) = "Applied"
governors_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					governorscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Governors State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open governorscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

'//////// Grand Valley State Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_grandvalleystate_query="SELECT Count(distinct UIN) grandvalleystate_cnf_students FROM Applicants where UGCollege = 'Grand Valley State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_grandvalleystate_query,conn
    If rs("grandvalleystate_cnf_students") <> 0 Then
    pdf.GreyTitle("Grand Valley State University")
'pdf.FancyTable()



grandvalleystate_cnf_rows = rs("grandvalleystate_cnf_students")
grandvalleystate_cnf_cols = 5
Dim grandvalleystate_cnf_col(5)
grandvalleystate_cnf_col(1) = "Banner # "
grandvalleystate_cnf_col(2) = "First Name"
grandvalleystate_cnf_col(3) = "Last Name"
grandvalleystate_cnf_col(4) = "Applied"
grandvalleystate_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					grandvalleystatecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Grand Valley State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open grandvalleystatecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 
    

'//////// Greenville College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_greenville_query="SELECT Count(distinct UIN) greenville_cnf_students FROM Applicants where UGCollege = 'Greenville College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_greenville_query,conn
    If rs("greenville_cnf_students") <> 0 Then
    pdf.GreyTitle("Greenville College")
'pdf.FancyTable()



greenville_cnf_rows = rs("greenville_cnf_students")
greenville_cnf_cols = 5
Dim greenville_cnf_col(5)
greenville_cnf_col(1) = "Banner # "
greenville_cnf_col(2) = "First Name"
greenville_cnf_col(3) = "Last Name"
greenville_cnf_col(4) = "Applied"
greenville_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					greenvillecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Greenville College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open greenvillecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 
    
                   
'////////  Grinnell Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_grinnell_query="SELECT Count(distinct UIN) grinnell_cnf_students FROM Applicants where UGCollege = 'Grinnell College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_grinnell_query,conn
If rs("grinnell_cnf_students") <> 0 Then
pdf.GreyTitle("Grinnell College")
'pdf.FancyTable()


grinnell_cnf_rows = rs("grinnell_cnf_students")
grinnell_cnf_cols = 5
Dim grinnell_cnf_col(5)
grinnell_cnf_col(1) = "Banner # "
grinnell_cnf_col(2) = "First Name"
grinnell_cnf_col(3) = "Last Name"
grinnell_cnf_col(4) = "Applied"
grinnell_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					grinnellcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Grinnell College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open grinnellcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
                   
'////////  Guilford College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_guilford_query="SELECT Count(distinct UIN) guilford_cnf_students FROM Applicants where UGCollege = 'Guilford College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_guilford_query,conn
If rs("guilford_cnf_students") <> 0 Then
pdf.GreyTitle("Guilford College")
'pdf.FancyTable()


guilford_cnf_rows = rs("guilford_cnf_students")
guilford_cnf_cols = 5
Dim guilford_cnf_col(5)
guilford_cnf_col(1) = "Banner # "
guilford_cnf_col(2) = "First Name"
guilford_cnf_col(3) = "Last Name"
guilford_cnf_col(4) = "Applied"
guilford_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					guilfordcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Guilford College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open guilfordcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
    
'//////// Hamline University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_hamline_query="SELECT Count(distinct UIN) hamline_cnf_students FROM Applicants where UGCollege = 'Hamline University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_hamline_query,conn
    If rs("hamline_cnf_students") <> 0 Then
    pdf.GreyTitle("Hamline University")
'pdf.FancyTable()

'//////// Students ////////////

hamline_cnf_rows = rs("hamline_cnf_students")
hamline_cnf_cols = 5
Dim hamline_cnf_col(5)
hamline_cnf_col(1) = "Banner # "
hamline_cnf_col(2) = "First Name"
hamline_cnf_col(3) = "Last Name"
hamline_cnf_col(4) = "Applied"
hamline_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					hamlinecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Hamline University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open hamlinecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
  
'//////// Hebrew Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_hebrew_query="SELECT Count(distinct UIN) hebrew_cnf_students FROM Applicants where UGCollege = 'Hebrew University of Jerusalem' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_hebrew_query,conn
    If rs("hebrew_cnf_students") <> 0 Then
    pdf.GreyTitle("Hebrew University of Jerusalem")
'pdf.FancyTable()

'//////// Students ////////////

hebrew_cnf_rows = rs("hebrew_cnf_students")
hebrew_cnf_cols = 5
Dim hebrew_cnf_col(5)
hebrew_cnf_col(1) = "Banner # "
hebrew_cnf_col(2) = "First Name"
hebrew_cnf_col(3) = "Last Name"
hebrew_cnf_col(4) = "Applied"
hebrew_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					hebrewcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Hebrew University of Jerusalem' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open hebrewcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
     
'//////// Hope College Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_hope_query="SELECT Count(distinct UIN) hope_cnf_students FROM Applicants where UGCollege = 'Hope College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_hope_query,conn
    If rs("hope_cnf_students") <> 0 Then
    pdf.GreyTitle("Hope College")
'pdf.FancyTable()

'//////// Students ////////////

hope_cnf_rows = rs("hope_cnf_students")
hope_cnf_cols = 5
Dim hope_cnf_col(5)
hope_cnf_col(1) = "Banner # "
hope_cnf_col(2) = "First Name"
hope_cnf_col(3) = "Last Name"
hope_cnf_col(4) = "Applied"
hope_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					hopecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Hope College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open hopecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                       
    '//////// Houghton College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_houghton_query="SELECT Count(distinct UIN) houghton_cnf_students FROM Applicants where UGCollege = 'Houghton College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_houghton_query,conn
    If rs("houghton_cnf_students") <> 0 Then 
    pdf.GreyTitle("Houghton College")
'pdf.FancyTable()

houghton_cnf_rows = rs("houghton_cnf_students")
houghton_cnf_cols = 5
Dim houghton_cnf_col(5)
houghton_cnf_col(1) = "Banner # "
houghton_cnf_col(2) = "First Name"
houghton_cnf_col(3) = "Last Name"
houghton_cnf_col(4) = "Applied"
houghton_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					houghtoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Houghton College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open houghtoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                       
    '//////// Howard University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_howard_query="SELECT Count(distinct UIN) howard_cnf_students FROM Applicants where UGCollege = 'Howard University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_howard_query,conn
    If rs("howard_cnf_students") <> 0 Then 
    pdf.GreyTitle("Howard University")
'pdf.FancyTable()

howard_cnf_rows = rs("howard_cnf_students")
howard_cnf_cols = 5
Dim howard_cnf_col(5)
howard_cnf_col(1) = "Banner # "
howard_cnf_col(2) = "First Name"
howard_cnf_col(3) = "Last Name"
howard_cnf_col(4) = "Applied"
howard_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					howardcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Howard University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open howardcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn


'//////// Illinois State Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_illinoisstate_query="SELECT Count(distinct UIN) illinoisstate_cnf_students FROM Applicants where UGCollege = 'Illinois State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_illinoisstate_query,conn
    If rs("illinoisstate_cnf_students") <> 0 Then
    pdf.GreyTitle("Illinois State University")
'pdf.FancyTable()



illinoisstate_cnf_rows = rs("illinoisstate_cnf_students")
illinoisstate_cnf_cols = 5
Dim illinoisstate_cnf_col(5)
illinoisstate_cnf_col(1) = "Banner # "
illinoisstate_cnf_col(2) = "First Name"
illinoisstate_cnf_col(3) = "Last Name"
illinoisstate_cnf_col(4) = "Applied"
illinoisstate_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					illinoisstatecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Illinois State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open illinoisstatecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If                

'//////// Illinois Wesleyan University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_illinoiswesleyan_query="SELECT Count(distinct UIN) illinoiswesleyan_cnf_students FROM Applicants where UGCollege = 'Illinois Wesleyan University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_illinoiswesleyan_query,conn
    If rs("illinoiswesleyan_cnf_students") <> 0 Then
    pdf.GreyTitle("Illinois Wesleyan University")
'pdf.FancyTable()

illinoiswesleyan_cnf_rows = rs("illinoiswesleyan_cnf_students")
illinoiswesleyan_cnf_cols = 5
Dim illinoiswesleyan_cnf_col(5)
illinoiswesleyan_cnf_col(1) = "Banner # "
illinoiswesleyan_cnf_col(2) = "First Name"
illinoiswesleyan_cnf_col(3) = "Last Name"
illinoiswesleyan_cnf_col(4) = "Applied"
illinoiswesleyan_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					illinoiswesleyancnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Illinois Wesleyan University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open illinoiswesleyancnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If                


    '//////// Indiana State Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_indianastate_query="SELECT Count(distinct UIN) indianastate_cnf_students FROM Applicants where UGCollege = 'Indiana State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_indianastate_query,conn
    If rs("indianastate_cnf_students") <> 0 Then
    pdf.GreyTitle("Indiana State University")
'pdf.FancyTable()

'//////// Students ////////////

indianastate_cnf_rows = rs("indianastate_cnf_students")
indianastate_cnf_cols = 5
Dim indianastate_cnf_col(5)
indianastate_cnf_col(1) = "Banner # "
indianastate_cnf_col(2) = "First Name"
indianastate_cnf_col(3) = "Last Name"
indianastate_cnf_col(4) = "Applied"
indianastate_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					indianastatecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Indiana State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open indianastatecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                        
    '//////// Indiana Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_indiana_query="SELECT Count(distinct UIN) indiana_cnf_students FROM Applicants where UGCollege = 'Indiana University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_indiana_query,conn
    If rs("indiana_cnf_students") <> 0 Then 
    pdf.GreyTitle("Indiana University")
'pdf.FancyTable()

indiana_cnf_rows = rs("indiana_cnf_students")
indiana_cnf_cols = 5
Dim indiana_cnf_col(5)
indiana_cnf_col(1) = "Banner # "
indiana_cnf_col(2) = "First Name"
indiana_cnf_col(3) = "Last Name"
indiana_cnf_col(4) = "Applied"
indiana_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					indianacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Indiana University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open indianacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
        
    '//////// Indiana Bloomington Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_indianabloomington_query="SELECT Count(distinct UIN) indianabloomington_cnf_students FROM Applicants where UGCollege = 'Indiana University Bloomington' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_indianabloomington_query,conn
    If rs("indianabloomington_cnf_students") <> 0 Then 
    pdf.GreyTitle("Indiana University Bloomington")
'pdf.FancyTable()

indianabloomington_cnf_rows = rs("indianabloomington_cnf_students")
indianabloomington_cnf_cols = 5
Dim indianabloomington_cnf_col(5)
indianabloomington_cnf_col(1) = "Banner # "
indianabloomington_cnf_col(2) = "First Name"
indianabloomington_cnf_col(3) = "Last Name"
indianabloomington_cnf_col(4) = "Applied"
indianabloomington_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					indianabloomingtoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Indiana University Bloomington' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open indianabloomingtoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
        
    '//////// Indiana University South Bend Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_indianasouthbend_query="SELECT Count(distinct UIN) indianasouthbend_cnf_students FROM Applicants where UGCollege = 'Indiana University South Bend' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_indianasouthbend_query,conn
    If rs("indianasouthbend_cnf_students") <> 0 Then 
    pdf.GreyTitle("Indiana University South Bend")
'pdf.FancyTable()

indianasouthbend_cnf_rows = rs("indianasouthbend_cnf_students")
indianasouthbend_cnf_cols = 5
Dim indianasouthbend_cnf_col(5)
indianasouthbend_cnf_col(1) = "Banner # "
indianasouthbend_cnf_col(2) = "First Name"
indianasouthbend_cnf_col(3) = "Last Name"
indianasouthbend_cnf_col(4) = "Applied"
indianasouthbend_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					indianasouthbendcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Indiana University South Bend' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open indianasouthbendcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

         '//////// Indiana Wesleyan Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_indianaWesleyan_query="SELECT Count(distinct UIN) indianaWesleyan_cnf_students FROM Applicants where UGCollege = 'Indiana Wesleyan University' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_indianaWesleyan_query,conn
    If rs("indianaWesleyan_cnf_students") <> 0 Then 
    pdf.GreyTitle("Indiana Wesleyan University")
'pdf.FancyTable()

indianaWesleyan_cnf_rows = rs("indianaWesleyan_cnf_students")
indianaWesleyan_cnf_cols = 5
Dim indianaWesleyan_cnf_col(5)
indianaWesleyan_cnf_col(1) = "Banner # "
indianaWesleyan_cnf_col(2) = "First Name"
indianaWesleyan_cnf_col(3) = "Last Name"
indianaWesleyan_cnf_col(4) = "Applied"
indianaWesleyan_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					indianaWesleyancnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Indiana Wesleyan University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open indianaWesleyancnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

         '//////// Iowa State University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_isu_query="SELECT Count(distinct UIN) isu_cnf_students FROM Applicants where UGCollege = 'Iowa State University' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_isu_query,conn
    If rs("isu_cnf_students") <> 0 Then 
    pdf.GreyTitle("Iowa State University")
'pdf.FancyTable()

isu_cnf_rows = rs("isu_cnf_students")
isu_cnf_cols = 5
Dim isu_cnf_col(5)
isu_cnf_col(1) = "Banner # "
isu_cnf_col(2) = "First Name"
isu_cnf_col(3) = "Last Name"
isu_cnf_col(4) = "Applied"
isu_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					isucnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Iowa State University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open isucnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
     
         '//////// Ithaca College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_ithaca_query="SELECT Count(distinct UIN) ithaca_cnf_students FROM Applicants where UGCollege = 'Ithaca College' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_ithaca_query,conn
    If rs("ithaca_cnf_students") <> 0 Then 
    pdf.GreyTitle("Ithaca College")
'pdf.FancyTable()

ithaca_cnf_rows = rs("ithaca_cnf_students")
ithaca_cnf_cols = 5
Dim ithaca_cnf_col(5)
ithaca_cnf_col(1) = "Banner # "
ithaca_cnf_col(2) = "First Name"
ithaca_cnf_col(3) = "Last Name"
ithaca_cnf_col(4) = "Applied"
ithaca_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					ithacacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Ithaca College' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open ithacacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

'//////// John Carroll Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_johncarroll_query="SELECT Count(distinct UIN) johncarroll_cnf_students FROM Applicants where UGCollege = 'John Carroll University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_johncarroll_query,conn
    If rs("johncarroll_cnf_students") <> 0 Then
    pdf.GreyTitle("John Carroll University")
'pdf.FancyTable()



johncarroll_cnf_rows = rs("johncarroll_cnf_students")
johncarroll_cnf_cols = 5
Dim johncarroll_cnf_col(5)
johncarroll_cnf_col(1) = "Banner # "
johncarroll_cnf_col(2) = "First Name"
johncarroll_cnf_col(3) = "Last Name"
johncarroll_cnf_col(4) = "Applied"
johncarroll_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					johncarrollcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'John Carroll University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open johncarrollcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    
    
      'start from here tomorrow            
'////////  Kalamazoo Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_kalamazoo_query="SELECT Count(distinct UIN) kalamazoo_cnf_students FROM Applicants where UGCollege = 'Kalamazoo College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_kalamazoo_query,conn
If rs("kalamazoo_cnf_students") <> 0 Then
pdf.GreyTitle("Kalamazoo College")
'pdf.FancyTable()


kalamazoo_cnf_rows = rs("kalamazoo_cnf_students")
kalamazoo_cnf_cols = 5
Dim kalamazoo_cnf_col(5)
kalamazoo_cnf_col(1) = "Banner # "
kalamazoo_cnf_col(2) = "First Name"
kalamazoo_cnf_col(3) = "Last Name"
kalamazoo_cnf_col(4) = "Applied"
kalamazoo_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					kalamazoocnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Kalamazoo College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open kalamazoocnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
 
'//////// Kansas Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_kansas_query="SELECT Count(distinct UIN) kansas_cnf_students FROM Applicants where UGCollege = 'Kansas State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_kansas_query,conn
    If rs("kansas_cnf_students") <> 0 Then
    pdf.GreyTitle("Kansas State University")
'pdf.FancyTable()

'//////// Students ////////////

kansas_cnf_rows = rs("kansas_cnf_students")
kansas_cnf_cols = 5
Dim kansas_cnf_col(5)
kansas_cnf_col(1) = "Banner # "
kansas_cnf_col(2) = "First Name"
kansas_cnf_col(3) = "Last Name"
kansas_cnf_col(4) = "Applied"
kansas_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					kansascnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Kansas State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open kansascnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                        
    '//////// Keene State College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Keene_query="SELECT Count(distinct UIN) Keene_cnf_students FROM Applicants where UGCollege = 'Keene State College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_Keene_query,conn
    If rs("Keene_cnf_students") <> 0 Then 
    pdf.GreyTitle("Keene State College")
'pdf.FancyTable()

Keene_cnf_rows = rs("Keene_cnf_students")
Keene_cnf_cols = 5
Dim Keene_cnf_col(5)
Keene_cnf_col(1) = "Banner # "
Keene_cnf_col(2) = "First Name"
Keene_cnf_col(3) = "Last Name"
Keene_cnf_col(4) = "Applied"
Keene_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Keenecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Keene State College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open Keenecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

                         
    '//////// Kendall College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_kendall_query="SELECT Count(distinct UIN) kendall_cnf_students FROM Applicants where UGCollege = 'Kendall College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_kendall_query,conn
    If rs("kendall_cnf_students") <> 0 Then 
    pdf.GreyTitle("Kendall College")
'pdf.FancyTable()

kendall_cnf_rows = rs("kendall_cnf_students")
kendall_cnf_cols = 5
Dim kendall_cnf_col(5)
kendall_cnf_col(1) = "Banner # "
kendall_cnf_col(2) = "First Name"
kendall_cnf_col(3) = "Last Name"
kendall_cnf_col(4) = "Applied"
kendall_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					kendallcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Kendall College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open kendallcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

                        
    '//////// Kentucky Wesleyan College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_kentucky_query="SELECT Count(distinct UIN) kentucky_cnf_students FROM Applicants where UGCollege = 'Kentucky Wesleyan College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_kentucky_query,conn
    If rs("kentucky_cnf_students") <> 0 Then 
    pdf.GreyTitle("Kentucky Wesleyan College")
'pdf.FancyTable()

kentucky_cnf_rows = rs("kentucky_cnf_students")
kentucky_cnf_cols = 5
Dim kentucky_cnf_col(5)
kentucky_cnf_col(1) = "Banner # "
kentucky_cnf_col(2) = "First Name"
kentucky_cnf_col(3) = "Last Name"
kentucky_cnf_col(4) = "Applied"
kentucky_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					kentuckycnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Kentucky Wesleyan College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open kentuckycnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

                        
    '//////// Kenyon Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_kenyon_query="SELECT Count(distinct UIN) kenyon_cnf_students FROM Applicants where UGCollege = 'Kenyon College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_kenyon_query,conn
    If rs("kenyon_cnf_students") <> 0 Then 
    pdf.GreyTitle("Kenyon College")
'pdf.FancyTable()

kenyon_cnf_rows = rs("kenyon_cnf_students")
kenyon_cnf_cols = 5
Dim kenyon_cnf_col(5)
kenyon_cnf_col(1) = "Banner # "
kenyon_cnf_col(2) = "First Name"
kenyon_cnf_col(3) = "Last Name"
kenyon_cnf_col(4) = "Applied"
kenyon_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					kenyoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Kenyon College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open kenyoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

'//////// Lake Forest College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_lakeforest_query="SELECT Count(distinct UIN) lakeforest_cnf_students FROM Applicants where UGCollege = 'Lake Forest College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_lakeforest_query,conn
    If rs("lakeforest_cnf_students") <> 0 Then
    pdf.GreyTitle("Lake Forest College")
'pdf.FancyTable()

lakeforest_cnf_rows = rs("lakeforest_cnf_students")
lakeforest_cnf_cols = 5
Dim lakeforest_cnf_col(5)
lakeforest_cnf_col(1) = "Banner # "
lakeforest_cnf_col(2) = "First Name"
lakeforest_cnf_col(3) = "Last Name"
lakeforest_cnf_col(4) = "Applied"
lakeforest_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					lakeforestcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Lake Forest College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open lakeforestcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 
 
'//////// Lawrence University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_lawrence_query="SELECT Count(distinct UIN) lawrence_cnf_students FROM Applicants where UGCollege = 'Lawrence University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_lawrence_query,conn
    If rs("lawrence_cnf_students") <> 0 Then
    pdf.GreyTitle("Lawrence University")
'pdf.FancyTable()



lawrence_cnf_rows = rs("lawrence_cnf_students")
lawrence_cnf_cols = 5
Dim lawrence_cnf_col(5)
lawrence_cnf_col(1) = "Banner # "
lawrence_cnf_col(2) = "First Name"
lawrence_cnf_col(3) = "Last Name"
lawrence_cnf_col(4) = "Applied"
lawrence_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					lawrencecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Lawrence University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open lawrencecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 
 
'//////// Lewis University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_lewis_query="SELECT Count(distinct UIN) lewis_cnf_students FROM Applicants where UGCollege = 'Lewis University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_lewis_query,conn
    If rs("lewis_cnf_students") <> 0 Then
    pdf.GreyTitle("Lewis University")
'pdf.FancyTable()



lewis_cnf_rows = rs("lewis_cnf_students")
lewis_cnf_cols = 5
Dim lewis_cnf_col(5)
lewis_cnf_col(1) = "Banner # "
lewis_cnf_col(2) = "First Name"
lewis_cnf_col(3) = "Last Name"
lewis_cnf_col(4) = "Applied"
lewis_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					lewiscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Lewis University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open lewiscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 
 
'//////// Lincoln University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_lincolnuniv_query="SELECT Count(distinct UIN) lincolnuniv_cnf_students FROM Applicants where UGCollege = 'Lincoln University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_lincolnuniv_query,conn
    If rs("lincolnuniv_cnf_students") <> 0 Then
    pdf.GreyTitle("Lincoln University")
'pdf.FancyTable()



lincolnuniv_cnf_rows = rs("lincolnuniv_cnf_students")
lincolnuniv_cnf_cols = 5
Dim lincolnuniv_cnf_col(5)
lincolnuniv_cnf_col(1) = "Banner # "
lincolnuniv_cnf_col(2) = "First Name"
lincolnuniv_cnf_col(3) = "Last Name"
lincolnuniv_cnf_col(4) = "Applied"
lincolnuniv_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					lincolnunivcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Lincoln University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open lincolnunivcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 
  
'//////// Lindenwood University-Belleville Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_lindenwood_query="SELECT Count(distinct UIN) lindenwood_cnf_students FROM Applicants where UGCollege = 'Lindenwood University-Belleville' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_lindenwood_query,conn
    If rs("lindenwood_cnf_students") <> 0 Then
    pdf.GreyTitle("Lindenwood University-Belleville")
'pdf.FancyTable()

lindenwood_cnf_rows = rs("lindenwood_cnf_students")
lindenwood_cnf_cols = 5
Dim lindenwood_cnf_col(5)
lindenwood_cnf_col(1) = "Banner # "
lindenwood_cnf_col(2) = "First Name"
lindenwood_cnf_col(3) = "Last Name"
lindenwood_cnf_col(4) = "Applied"
lindenwood_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					lindenwoodcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Lindenwood University-Belleville' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open lindenwoodcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 
   
   
'//////// loras Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_loras_query="SELECT Count(distinct UIN) loras_cnf_students FROM Applicants where UGCollege = 'loras College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_loras_query,conn
    If rs("loras_cnf_students") <> 0 Then
    pdf.GreyTitle("loras College")
'pdf.FancyTable()



loras_cnf_rows = rs("loras_cnf_students")
loras_cnf_cols = 5
Dim loras_cnf_col(5)
loras_cnf_col(1) = "Banner # "
loras_cnf_col(2) = "First Name"
loras_cnf_col(3) = "Last Name"
loras_cnf_col(4) = "Applied"
loras_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					lorascnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'loras College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open lorascnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 
    
    '//////// Loyola Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_loyola_query="SELECT Count(distinct UIN) loyola_cnf_students FROM Applicants where UGCollege = 'Loyola University Chicago' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_loyola_query,conn
    If rs("loyola_cnf_students") <> 0 Then
    pdf.GreyTitle("Loyola University Chicago")
'pdf.FancyTable()

'//////// Students ////////////

loyola_cnf_rows = rs("loyola_cnf_students")
loyola_cnf_cols = 5
Dim loyola_cnf_col(5)
loyola_cnf_col(1) = "Banner # "
loyola_cnf_col(2) = "First Name"
loyola_cnf_col(3) = "Last Name"
loyola_cnf_col(4) = "Applied"
loyola_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					loyolacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Loyola University Chicago' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open loyolacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
   
    '//////// Luther College Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_luther_query="SELECT Count(distinct UIN) luther_cnf_students FROM Applicants where UGCollege = 'Luther College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_luther_query,conn
    If rs("luther_cnf_students") <> 0 Then
    pdf.GreyTitle("Luther College")
'pdf.FancyTable()

'//////// Students ////////////

luther_cnf_rows = rs("luther_cnf_students")
luther_cnf_cols = 5
Dim luther_cnf_col(5)
luther_cnf_col(1) = "Banner # "
luther_cnf_col(2) = "First Name"
luther_cnf_col(3) = "Last Name"
luther_cnf_col(4) = "Applied"
luther_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					luthercnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Luther College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open luthercnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  

     '//////// Macalester College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_macalester_query="SELECT Count(distinct UIN) macalester_cnf_students FROM Applicants where UGCollege = 'Macalester College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_macalester_query,conn
    If rs("macalester_cnf_students") <> 0 Then 
    pdf.GreyTitle("Macalester College")
'pdf.FancyTable()

macalester_cnf_rows = rs("macalester_cnf_students")
macalester_cnf_cols = 5
Dim macalester_cnf_col(5)
macalester_cnf_col(1) = "Banner # "
macalester_cnf_col(2) = "First Name"
macalester_cnf_col(3) = "Last Name"
macalester_cnf_col(4) = "Applied"
macalester_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					macalestercnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Macalester College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open macalestercnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                          
    '//////// Marquette University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_marquette_query="SELECT Count(distinct UIN) marquette_cnf_students FROM Applicants where UGCollege = 'Marquette University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_marquette_query,conn
    If rs("marquette_cnf_students") <> 0 Then 
    pdf.GreyTitle("Marquette University")
'pdf.FancyTable()

marquette_cnf_rows = rs("marquette_cnf_students")
marquette_cnf_cols = 5
Dim marquette_cnf_col(5)
marquette_cnf_col(1) = "Banner # "
marquette_cnf_col(2) = "First Name"
marquette_cnf_col(3) = "Last Name"
marquette_cnf_col(4) = "Applied"
marquette_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					marquettecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Marquette University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open marquettecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                        
    '//////// Miami University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_miami_query="SELECT Count(distinct UIN) miami_cnf_students FROM Applicants where UGCollege = 'Miami University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_miami_query,conn
    If rs("miami_cnf_students") <> 0 Then 
    pdf.GreyTitle("Miami University")
'pdf.FancyTable()

miami_cnf_rows = rs("miami_cnf_students")
miami_cnf_cols = 5
Dim miami_cnf_col(5)
miami_cnf_col(1) = "Banner # "
miami_cnf_col(2) = "First Name"
miami_cnf_col(3) = "Last Name"
miami_cnf_col(4) = "Applied"
miami_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					miamicnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Miami University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open miamicnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                         
    '//////// Miami University-Oxford Ohio Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_miamioxford_query="SELECT Count(distinct UIN) miamioxford_cnf_students FROM Applicants where UGCollege = 'Miami University-Oxford Ohio' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_miamioxford_query,conn
    If rs("miamioxford_cnf_students") <> 0 Then 
    pdf.GreyTitle("Miami University-Oxford Ohio")
'pdf.FancyTable()

miamioxford_cnf_rows = rs("miamioxford_cnf_students")
miamioxford_cnf_cols = 5
Dim miamioxford_cnf_col(5)
miamioxford_cnf_col(1) = "Banner # "
miamioxford_cnf_col(2) = "First Name"
miamioxford_cnf_col(3) = "Last Name"
miamioxford_cnf_col(4) = "Applied"
miamioxford_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					miamioxfordcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Miami University-Oxford Ohio' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open miamioxfordcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                        
    '//////// Michigan State University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_msu_query="SELECT Count(distinct UIN) msu_cnf_students FROM Applicants where UGCollege = 'Michigan State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_msu_query,conn
    If rs("msu_cnf_students") <> 0 Then 
    pdf.GreyTitle("Michigan State University")
'pdf.FancyTable()

msu_cnf_rows = rs("msu_cnf_students")
msu_cnf_cols = 5
Dim msu_cnf_col(5)
msu_cnf_col(1) = "Banner # "
msu_cnf_col(2) = "First Name"
msu_cnf_col(3) = "Last Name"
msu_cnf_col(4) = "Applied"
msu_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					msucnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Michigan State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open msucnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                        
    '//////// Moddy Bible Institute Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_moddybible_query="SELECT Count(distinct UIN) moddybible_cnf_students FROM Applicants where UGCollege = 'Moddy Bible Institute' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_moddybible_query,conn
    If rs("moddybible_cnf_students") <> 0 Then 
    pdf.GreyTitle("Moddy Bible Institute")
'pdf.FancyTable()

moddybible_cnf_rows = rs("moddybible_cnf_students")
moddybible_cnf_cols = 5
Dim moddybible_cnf_col(5)
moddybible_cnf_col(1) = "Banner # "
moddybible_cnf_col(2) = "First Name"
moddybible_cnf_col(3) = "Last Name"
moddybible_cnf_col(4) = "Applied"
moddybible_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					moddybiblecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Moddy Bible Institute' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open moddybiblecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

'//////// National Louis Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_nationallouis_query="SELECT Count(distinct UIN) nationallouis_cnf_students FROM Applicants where UGCollege = 'National Louis University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_nationallouis_query,conn
    If rs("nationallouis_cnf_students") <> 0 Then
    pdf.GreyTitle("National Louis University")
'pdf.FancyTable()



nationallouis_cnf_rows = rs("nationallouis_cnf_students")
nationallouis_cnf_cols = 5
Dim nationallouis_cnf_col(5)
nationallouis_cnf_col(1) = "Banner # "
nationallouis_cnf_col(2) = "First Name"
nationallouis_cnf_col(3) = "Last Name"
nationallouis_cnf_col(4) = "Applied"
nationallouis_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					nationallouiscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'National Louis University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open nationallouiscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If  
    

'//////// New York University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_nyu_query="SELECT Count(distinct UIN) nyu_cnf_students FROM Applicants where UGCollege = 'New York University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_nyu_query,conn
    If rs("nyu_cnf_students") <> 0 Then
    pdf.GreyTitle("New York University")
'pdf.FancyTable()

nyu_cnf_rows = rs("nyu_cnf_students")
nyu_cnf_cols = 5
Dim nyu_cnf_col(5)
nyu_cnf_col(1) = "Banner # "
nyu_cnf_col(2) = "First Name"
nyu_cnf_col(3) = "Last Name"
nyu_cnf_col(4) = "Applied"
nyu_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					nyucnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'New York University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open nyucnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If      
                  

'//////// North Central College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_northcentral_query="SELECT Count(distinct UIN) northcentral_cnf_students FROM Applicants where UGCollege = 'North Central College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_northcentral_query,conn
    If rs("northcentral_cnf_students") <> 0 Then
    pdf.GreyTitle("North Central College")
'pdf.FancyTable()



northcentral_cnf_rows = rs("northcentral_cnf_students")
northcentral_cnf_cols = 5
Dim northcentral_cnf_col(5)
northcentral_cnf_col(1) = "Banner # "
northcentral_cnf_col(2) = "First Name"
northcentral_cnf_col(3) = "Last Name"
northcentral_cnf_col(4) = "Applied"
northcentral_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					northcentralcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'North Central College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open northcentralcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If      

'//////// North Park University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_northpark_query="SELECT Count(distinct UIN) northpark_cnf_students FROM Applicants where UGCollege = 'North Park University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_northpark_query,conn
    If rs("northpark_cnf_students") <> 0 Then
    pdf.GreyTitle("North Park University")
'pdf.FancyTable()



northpark_cnf_rows = rs("northpark_cnf_students")
northpark_cnf_cols = 5
Dim northpark_cnf_col(5)
northpark_cnf_col(1) = "Banner # "
northpark_cnf_col(2) = "First Name"
northpark_cnf_col(3) = "Last Name"
northpark_cnf_col(4) = "Applied"
northpark_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					northparkcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'North Park University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open northparkcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If      
                  
'////////  Northeastern Illinois Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_northeasternillinois_query="SELECT Count(distinct UIN) northeasternillinois_cnf_students FROM Applicants where UGCollege = 'Northeastern Illinois University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_northeasternillinois_query,conn
If rs("northeasternillinois_cnf_students") <> 0 Then
pdf.GreyTitle("Northeastern Illinois University")
'pdf.FancyTable()


northeasternillinois_cnf_rows = rs("northeasternillinois_cnf_students")
northeasternillinois_cnf_cols = 5
Dim northeasternillinois_cnf_col(5)
northeasternillinois_cnf_col(1) = "Banner # "
northeasternillinois_cnf_col(2) = "First Name"
northeasternillinois_cnf_col(3) = "Last Name"
northeasternillinois_cnf_col(4) = "Applied"
northeasternillinois_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					northeasternillinoiscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Northeastern Illinois University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open northeasternillinoiscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
    
'//////// Northern Arizona University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_northernarizona_query="SELECT Count(distinct UIN) northernarizona_cnf_students FROM Applicants where UGCollege = 'Northern Arizona University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_northernarizona_query,conn
    If rs("northernarizona_cnf_students") <> 0 Then
    pdf.GreyTitle("Northern Arizona University")
'pdf.FancyTable()

'//////// Students ////////////

northernarizona_cnf_rows = rs("northernarizona_cnf_students")
northernarizona_cnf_cols = 5
Dim northernarizona_cnf_col(5)
northernarizona_cnf_col(1) = "Banner # "
northernarizona_cnf_col(2) = "First Name"
northernarizona_cnf_col(3) = "Last Name"
northernarizona_cnf_col(4) = "Applied"
northernarizona_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					northernarizonacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Northern Arizona University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open northernarizonacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
  
'//////// Northern Michigan Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_northernmichigan_query="SELECT Count(distinct UIN) northernmichigan_cnf_students FROM Applicants where UGCollege = 'Northern Michigan University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_northernmichigan_query,conn
    If rs("northernmichigan_cnf_students") <> 0 Then
    pdf.GreyTitle("Northern Michigan University")
'pdf.FancyTable()

'//////// Students ////////////

northernmichigan_cnf_rows = rs("northernmichigan_cnf_students")
northernmichigan_cnf_cols = 5
Dim northernmichigan_cnf_col(5)
northernmichigan_cnf_col(1) = "Banner # "
northernmichigan_cnf_col(2) = "First Name"
northernmichigan_cnf_col(3) = "Last Name"
northernmichigan_cnf_col(4) = "Applied"
northernmichigan_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					northernmichigancnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Northern Michigan University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open northernmichigancnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                        
    '//////// Northern Illinois Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_northernillinois_query="SELECT Count(distinct UIN) northernillinois_cnf_students FROM Applicants where UGCollege = 'Northern Illinois University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_northernillinois_query,conn
    If rs("northernillinois_cnf_students") <> 0 Then 
    pdf.GreyTitle("Northern Illinois University")
'pdf.FancyTable()

northernillinois_cnf_rows = rs("northernillinois_cnf_students")
northernillinois_cnf_cols = 5
Dim northernillinois_cnf_col(5)
northernillinois_cnf_col(1) = "Banner # "
northernillinois_cnf_col(2) = "First Name"
northernillinois_cnf_col(3) = "Last Name"
northernillinois_cnf_col(4) = "Applied"
northernillinois_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					northernillinoiscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Northern Illinois University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open northernillinoiscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                         
    '//////// Northwest University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_northwest_query="SELECT Count(distinct UIN) northwest_cnf_students FROM Applicants where UGCollege = 'Northwest University' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_northwest_query,conn
    If rs("northwest_cnf_students") <> 0 Then 
    pdf.GreyTitle("Northwest University")
'pdf.FancyTable()

northwest_cnf_rows = rs("northwest_cnf_students")
northwest_cnf_cols = 5
Dim northwest_cnf_col(5)
northwest_cnf_col(1) = "Banner # "
northwest_cnf_col(2) = "First Name"
northwest_cnf_col(3) = "Last Name"
northwest_cnf_col(4) = "Applied"
northwest_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					northwestcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Northwest University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open northwestcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                        
    '//////// Northwestern University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_northwestern_query="SELECT Count(distinct UIN) northwestern_cnf_students FROM Applicants where UGCollege = 'Northwestern University' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_northwestern_query,conn
    If rs("northwestern_cnf_students") <> 0 Then 
    pdf.GreyTitle("Northwestern University")
'pdf.FancyTable()

northwestern_cnf_rows = rs("northwestern_cnf_students")
northwestern_cnf_cols = 5
Dim northwestern_cnf_col(5)
northwestern_cnf_col(1) = "Banner # "
northwestern_cnf_col(2) = "First Name"
northwestern_cnf_col(3) = "Last Name"
northwestern_cnf_col(4) = "Applied"
northwestern_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					northwesterncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Northwestern University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open northwesterncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn



'//////// Ohio State Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_ohiostate_query="SELECT Count(distinct UIN) ohiostate_cnf_students FROM Applicants where UGCollege = 'Ohio State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_ohiostate_query,conn
    If rs("ohiostate_cnf_students") <> 0 Then
    pdf.GreyTitle("Ohio State University")
'pdf.FancyTable()



ohiostate_cnf_rows = rs("ohiostate_cnf_students")
ohiostate_cnf_cols = 5
Dim ohiostate_cnf_col(5)
ohiostate_cnf_col(1) = "Banner # "
ohiostate_cnf_col(2) = "First Name"
ohiostate_cnf_col(3) = "Last Name"
ohiostate_cnf_col(4) = "Applied"
ohiostate_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					ohiostatecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Ohio State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open ohiostatecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If                
    
'//////// Olivet Nazarene University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_olivet_query="SELECT Count(distinct UIN) olivet_cnf_students FROM Applicants where UGCollege = 'Olivet Nazarene University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_olivet_query,conn
    If rs("olivet_cnf_students") <> 0 Then
    pdf.GreyTitle("Olivet Nazarene University")
'pdf.FancyTable()



olivet_cnf_rows = rs("olivet_cnf_students")
olivet_cnf_cols = 5
Dim olivet_cnf_col(5)
olivet_cnf_col(1) = "Banner # "
olivet_cnf_col(2) = "First Name"
olivet_cnf_col(3) = "Last Name"
olivet_cnf_col(4) = "Applied"
olivet_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					olivetcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Olivet Nazarene University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open olivetcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If                

    '//////// Pennsylvania State Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_pennsylvaniastate_query="SELECT Count(distinct UIN) pennsylvaniastate_cnf_students FROM Applicants where UGCollege = 'Pennsylvania State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_pennsylvaniastate_query,conn
    If rs("pennsylvaniastate_cnf_students") <> 0 Then
    pdf.GreyTitle("Pennsylvania State University")
'pdf.FancyTable()

'//////// Students ////////////

pennsylvaniastate_cnf_rows = rs("pennsylvaniastate_cnf_students")
pennsylvaniastate_cnf_cols = 5
Dim pennsylvaniastate_cnf_col(5)
pennsylvaniastate_cnf_col(1) = "Banner # "
pennsylvaniastate_cnf_col(2) = "First Name"
pennsylvaniastate_cnf_col(3) = "Last Name"
pennsylvaniastate_cnf_col(4) = "Applied"
pennsylvaniastate_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					pennsylvaniastatecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Pennsylvania State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open pennsylvaniastatecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                        
    '//////// Portland State Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_portlandstate_query="SELECT Count(distinct UIN) portlandstate_cnf_students FROM Applicants where UGCollege = 'Portland State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_portlandstate_query,conn
    If rs("portlandstate_cnf_students") <> 0 Then 
    pdf.GreyTitle("Portland State University")
'pdf.FancyTable()

portlandstate_cnf_rows = rs("portlandstate_cnf_students")
portlandstate_cnf_cols = 5
Dim portlandstate_cnf_col(5)
portlandstate_cnf_col(1) = "Banner # "
portlandstate_cnf_col(2) = "First Name"
portlandstate_cnf_col(3) = "Last Name"
portlandstate_cnf_col(4) = "Applied"
portlandstate_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					portlandstatecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Portland State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open portlandstatecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

    '//////// Purdue University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_purdue_query="SELECT Count(distinct UIN) purdue_cnf_students FROM Applicants where UGCollege = 'Purdue University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_purdue_query,conn
    If rs("purdue_cnf_students") <> 0 Then
    pdf.GreyTitle("Purdue University")
'pdf.FancyTable()

purdue_cnf_rows = rs("purdue_cnf_students")
purdue_cnf_cols = 5
Dim purdue_cnf_col(5)
purdue_cnf_col(1) = "Banner # "
purdue_cnf_col(2) = "First Name"
purdue_cnf_col(3) = "Last Name"
purdue_cnf_col(4) = "Applied"
purdue_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					purduecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Purdue University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open purduecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 

    '//////// Quincy Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_quincy_query="SELECT Count(distinct UIN) quincy_cnf_students FROM Applicants where UGCollege = 'Quincy University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_quincy_query,conn
    If rs("quincy_cnf_students") <> 0 Then
    pdf.GreyTitle("Quincy University")
'pdf.FancyTable()

quincy_cnf_rows = rs("quincy_cnf_students")
quincy_cnf_cols = 5
Dim quincy_cnf_col(5)
quincy_cnf_col(1) = "Banner # "
quincy_cnf_col(2) = "First Name"
quincy_cnf_col(3) = "Last Name"
quincy_cnf_col(4) = "Applied"
quincy_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					quincycnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Quincy University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open quincycnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 

    '//////// Rice Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_rice_query="SELECT Count(distinct UIN) rice_cnf_students FROM Applicants where UGCollege = 'Rice University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_rice_query,conn
    If rs("rice_cnf_students") <> 0 Then
    pdf.GreyTitle("Rice University")
'pdf.FancyTable()

rice_cnf_rows = rs("rice_cnf_students")
rice_cnf_cols = 5
Dim rice_cnf_col(5)
rice_cnf_col(1) = "Banner # "
rice_cnf_col(2) = "First Name"
rice_cnf_col(3) = "Last Name"
rice_cnf_col(4) = "Applied"
rice_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					ricecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Rice University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open ricecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 

'//////// Roosevelt Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_roosevelt_query="SELECT Count(distinct UIN) roosevelt_cnf_students FROM Applicants where UGCollege = 'Roosevelt University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_roosevelt_query,conn
    If rs("roosevelt_cnf_students") <> 0 Then
    pdf.GreyTitle("Roosevelt University")
'pdf.FancyTable()



roosevelt_cnf_rows = rs("roosevelt_cnf_students")
roosevelt_cnf_cols = 5
Dim roosevelt_cnf_col(5)
roosevelt_cnf_col(1) = "Banner # "
roosevelt_cnf_col(2) = "First Name"
roosevelt_cnf_col(3) = "Last Name"
roosevelt_cnf_col(4) = "Applied"
roosevelt_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					rooseveltcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Roosevelt University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open rooseveltcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 
             
'////////  Saint Louis Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_saintLouis_query="SELECT Count(distinct UIN) saintLouis_cnf_students FROM Applicants where UGCollege = 'Saint Louis University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_saintLouis_query,conn
If rs("saintLouis_cnf_students") <> 0 Then
pdf.GreyTitle("Saint Louis University")
'pdf.FancyTable()


saintLouis_cnf_rows = rs("saintLouis_cnf_students")
saintLouis_cnf_cols = 5
Dim saintLouis_cnf_col(5)
saintLouis_cnf_col(1) = "Banner # "
saintLouis_cnf_col(2) = "First Name"
saintLouis_cnf_col(3) = "Last Name"
saintLouis_cnf_col(4) = "Applied"
saintLouis_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					saintLouiscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Saint Louis University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open saintLouiscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
                  
'////////  Saint Xavier Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_saintxavier_query="SELECT Count(distinct UIN) saintxavier_cnf_students FROM Applicants where UGCollege = 'Saint Xavier University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_saintxavier_query,conn
If rs("saintxavier_cnf_students") <> 0 Then
pdf.GreyTitle("Saint Xavier University")
'pdf.FancyTable()


saintxavier_cnf_rows = rs("saintxavier_cnf_students")
saintxavier_cnf_cols = 5
Dim saintxavier_cnf_col(5)
saintxavier_cnf_col(1) = "Banner # "
saintxavier_cnf_col(2) = "First Name"
saintxavier_cnf_col(3) = "Last Name"
saintxavier_cnf_col(4) = "Applied"
saintxavier_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					saintxaviercnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Saint Xavier University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open saintxaviercnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
           
'////////  San Francisco State University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_sanfrancisco_query="SELECT Count(distinct UIN) sanfrancisco_cnf_students FROM Applicants where UGCollege = 'San Francisco State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_sanfrancisco_query,conn
If rs("sanfrancisco_cnf_students") <> 0 Then
pdf.GreyTitle("San Francisco State University")
'pdf.FancyTable()


sanfrancisco_cnf_rows = rs("sanfrancisco_cnf_students")
sanfrancisco_cnf_cols = 5
Dim sanfrancisco_cnf_col(5)
sanfrancisco_cnf_col(1) = "Banner # "
sanfrancisco_cnf_col(2) = "First Name"
sanfrancisco_cnf_col(3) = "Last Name"
sanfrancisco_cnf_col(4) = "Applied"
sanfrancisco_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					sanfranciscocnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'San Francisco State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open sanfranciscocnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
    
'//////// Santa Clara University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_santaclara_query="SELECT Count(distinct UIN) santaclara_cnf_students FROM Applicants where UGCollege = 'Santa Clara University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_santaclara_query,conn
    If rs("santaclara_cnf_students") <> 0 Then
    pdf.GreyTitle("Santa Clara University")
'pdf.FancyTable()

'//////// Students ////////////

santaclara_cnf_rows = rs("santaclara_cnf_students")
santaclara_cnf_cols = 5
Dim santaclara_cnf_col(5)
santaclara_cnf_col(1) = "Banner # "
santaclara_cnf_col(2) = "First Name"
santaclara_cnf_col(3) = "Last Name"
santaclara_cnf_col(4) = "Applied"
santaclara_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					santaclaracnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Santa Clara University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open santaclaracnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
     
'//////// Skidmore Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_skidmore_query="SELECT Count(distinct UIN) skidmore_cnf_students FROM Applicants where UGCollege = 'Skidmore College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_skidmore_query,conn
    If rs("skidmore_cnf_students") <> 0 Then
    pdf.GreyTitle("Skidmore College")
'pdf.FancyTable()

'//////// Students ////////////

skidmore_cnf_rows = rs("skidmore_cnf_students")
skidmore_cnf_cols = 5
Dim skidmore_cnf_col(5)
skidmore_cnf_col(1) = "Banner # "
skidmore_cnf_col(2) = "First Name"
skidmore_cnf_col(3) = "Last Name"
skidmore_cnf_col(4) = "Applied"
skidmore_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					skidmorecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Skidmore College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open skidmorecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
    
'//////// Smith Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_smith_query="SELECT Count(distinct UIN) smith_cnf_students FROM Applicants where UGCollege = 'Smith College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_smith_query,conn
    If rs("smith_cnf_students") <> 0 Then
    pdf.GreyTitle("Smith College")
'pdf.FancyTable()

'//////// Students ////////////

smith_cnf_rows = rs("smith_cnf_students")
smith_cnf_cols = 5
Dim smith_cnf_col(5)
smith_cnf_col(1) = "Banner # "
smith_cnf_col(2) = "First Name"
smith_cnf_col(3) = "Last Name"
smith_cnf_col(4) = "Applied"
smith_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					smithcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Smith College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open smithcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
          
'//////// South Dakota State University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_southdakota_query="SELECT Count(distinct UIN) southdakota_cnf_students FROM Applicants where UGCollege = 'South Dakota State University' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_southdakota_query,conn
    If rs("southdakota_cnf_students") <> 0 Then
    pdf.GreyTitle("South Dakota State University")
'pdf.FancyTable()

'//////// Students ////////////

southdakota_cnf_rows = rs("southdakota_cnf_students")
southdakota_cnf_cols = 5
Dim southdakota_cnf_col(5)
southdakota_cnf_col(1) = "Banner # "
southdakota_cnf_col(2) = "First Name"
southdakota_cnf_col(3) = "Last Name"
southdakota_cnf_col(4) = "Applied"
southdakota_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					southdakotacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'South Dakota State University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open southdakotacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                          
    '//////// Southern Illinois University Carbondale Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_southernillinoiscarbondale_query="SELECT Count(distinct UIN) southernillinoiscarbondale_cnf_students FROM Applicants where UGCollege = 'Southern Illinois University Carbondale' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_southernillinoiscarbondale_query,conn
    If rs("southernillinoiscarbondale_cnf_students") <> 0 Then 
    pdf.GreyTitle("Southern Illinois University Carbondale")
'pdf.FancyTable()

southernillinoiscarbondale_cnf_rows = rs("southernillinoiscarbondale_cnf_students")
southernillinoiscarbondale_cnf_cols = 5
Dim southernillinoiscarbondale_cnf_col(5)
southernillinoiscarbondale_cnf_col(1) = "Banner # "
southernillinoiscarbondale_cnf_col(2) = "First Name"
southernillinoiscarbondale_cnf_col(3) = "Last Name"
southernillinoiscarbondale_cnf_col(4) = "Applied"
southernillinoiscarbondale_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					southernillinoiscarbondalecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Southern Illinois University Carbondale' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open southernillinoiscarbondalecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                          
    '//////// Southern Illinois University Edwardsville Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_southernillinoisedwardsville_query="SELECT Count(distinct UIN) southernillinoisedwardsville_cnf_students FROM Applicants where UGCollege = 'Southern Illinois University Edwardsville' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_southernillinoisedwardsville_query,conn
    If rs("southernillinoisedwardsville_cnf_students") <> 0 Then 
    pdf.GreyTitle("Southern Illinois University Edwardsville")
'pdf.FancyTable()

southernillinoisedwardsville_cnf_rows = rs("southernillinoisedwardsville_cnf_students")
southernillinoisedwardsville_cnf_cols = 5
Dim southernillinoisedwardsville_cnf_col(5)
southernillinoisedwardsville_cnf_col(1) = "Banner # "
southernillinoisedwardsville_cnf_col(2) = "First Name"
southernillinoisedwardsville_cnf_col(3) = "Last Name"
southernillinoisedwardsville_cnf_col(4) = "Applied"
southernillinoisedwardsville_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					southernillinoisedwardsvillecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Southern Illinois University Edwardsville' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open southernillinoisedwardsvillecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                          
    '//////// Southern New Hampshire University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_southernnewhampshire_query="SELECT Count(distinct UIN) southernnewhampshire_cnf_students FROM Applicants where UGCollege = 'Southern New Hampshire University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_southernnewhampshire_query,conn
    If rs("southernnewhampshire_cnf_students") <> 0 Then 
    pdf.GreyTitle("Southern New Hampshire University")
'pdf.FancyTable()

southernnewhampshire_cnf_rows = rs("southernnewhampshire_cnf_students")
southernnewhampshire_cnf_cols = 5
Dim southernnewhampshire_cnf_col(5)
southernnewhampshire_cnf_col(1) = "Banner # "
southernnewhampshire_cnf_col(2) = "First Name"
southernnewhampshire_cnf_col(3) = "Last Name"
southernnewhampshire_cnf_col(4) = "Applied"
southernnewhampshire_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					southernnewhampshirecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Southern New Hampshire University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open southernnewhampshirecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                           
    '//////// Southern University Carbondale Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_southerncarbondale_query="SELECT Count(distinct UIN) southerncarbondale_cnf_students FROM Applicants where UGCollege = 'Southern University Carbondale' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_southerncarbondale_query,conn
    If rs("southerncarbondale_cnf_students") <> 0 Then 
    pdf.GreyTitle("Southern University Carbondale")
'pdf.FancyTable()

southerncarbondale_cnf_rows = rs("southerncarbondale_cnf_students")
southerncarbondale_cnf_cols = 5
Dim southerncarbondale_cnf_col(5)
southerncarbondale_cnf_col(1) = "Banner # "
southerncarbondale_cnf_col(2) = "First Name"
southerncarbondale_cnf_col(3) = "Last Name"
southerncarbondale_cnf_col(4) = "Applied"
southerncarbondale_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					southerncarbondalecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Southern University Carbondale' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open southerncarbondalecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
  
'//////// St Augustine College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_staugustine_query="SELECT Count(distinct UIN) staugustine_cnf_students FROM Applicants where UGCollege = 'St Augustine College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_staugustine_query,conn
    If rs("staugustine_cnf_students") <> 0 Then
    pdf.GreyTitle("St Augustine College")
'pdf.FancyTable()



staugustine_cnf_rows = rs("staugustine_cnf_students")
staugustine_cnf_cols = 5
Dim staugustine_cnf_col(5)
staugustine_cnf_col(1) = "Banner # "
staugustine_cnf_col(2) = "First Name"
staugustine_cnf_col(3) = "Last Name"
staugustine_cnf_col(4) = "Applied"
staugustine_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					staugustinecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'St Augustine College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open staugustinecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If   

'//////// St. Marys College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_stmarys_query="SELECT Count(distinct UIN) stmarys_cnf_students FROM Applicants where UGCollege = 'St. Marys College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_stmarys_query,conn
    If rs("stmarys_cnf_students") <> 0 Then
    pdf.GreyTitle("St. Marys College")
'pdf.FancyTable()



stmarys_cnf_rows = rs("stmarys_cnf_students")
stmarys_cnf_cols = 5
Dim stmarys_cnf_col(5)
stmarys_cnf_col(1) = "Banner # "
stmarys_cnf_col(2) = "First Name"
stmarys_cnf_col(3) = "Last Name"
stmarys_cnf_col(4) = "Applied"
stmarys_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					stmaryscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'St. Marys College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open stmaryscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If   

'//////// St. Norbert College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_stnorbert_query="SELECT Count(distinct UIN) stnorbert_cnf_students FROM Applicants where UGCollege = 'St. Norbert College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_stnorbert_query,conn
    If rs("stnorbert_cnf_students") <> 0 Then
    pdf.GreyTitle("St. Norbert College")
'pdf.FancyTable()



stnorbert_cnf_rows = rs("stnorbert_cnf_students")
stnorbert_cnf_cols = 5
Dim stnorbert_cnf_col(5)
stnorbert_cnf_col(1) = "Banner # "
stnorbert_cnf_col(2) = "First Name"
stnorbert_cnf_col(3) = "Last Name"
stnorbert_cnf_col(4) = "Applied"
stnorbert_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					stnorbertcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'St. Norbert College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open stnorbertcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If   

'//////// St. Xavier University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_stxavier_query="SELECT Count(distinct UIN) stxavier_cnf_students FROM Applicants where UGCollege = 'St. Xavier University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_stxavier_query,conn
    If rs("stxavier_cnf_students") <> 0 Then
    pdf.GreyTitle("St. Xavier University")
'pdf.FancyTable()



stxavier_cnf_rows = rs("stxavier_cnf_students")
stxavier_cnf_cols = 5
Dim stxavier_cnf_col(5)
stxavier_cnf_col(1) = "Banner # "
stxavier_cnf_col(2) = "First Name"
stxavier_cnf_col(3) = "Last Name"
stxavier_cnf_col(4) = "Applied"
stxavier_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					stxaviercnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'St. Xavier University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open stxaviercnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If   

    '//////// State University of New York Binghampton Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_SUNYbinghampton_query="SELECT Count(distinct UIN) SUNYbinghampton_cnf_students FROM Applicants where UGCollege = 'State University of New York Binghampton' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_SUNYbinghampton_query,conn
    If rs("SUNYbinghampton_cnf_students") <> 0 Then
    pdf.GreyTitle("State University of New York Binghampton")
'pdf.FancyTable()

'//////// Students ////////////

SUNYbinghampton_cnf_rows = rs("SUNYbinghampton_cnf_students")
SUNYbinghampton_cnf_cols = 5
Dim SUNYbinghampton_cnf_col(5)
SUNYbinghampton_cnf_col(1) = "Banner # "
SUNYbinghampton_cnf_col(2) = "First Name"
SUNYbinghampton_cnf_col(3) = "Last Name"
SUNYbinghampton_cnf_col(4) = "Applied"
SUNYbinghampton_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					SUNYbinghamptoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'State University of New York Binghampton' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open SUNYbinghamptoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                        
    '//////// Taylor Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_taylor_query="SELECT Count(distinct UIN) taylor_cnf_students FROM Applicants where UGCollege = 'Taylor University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_taylor_query,conn
    If rs("taylor_cnf_students") <> 0 Then 
    pdf.GreyTitle("Taylor University")
'pdf.FancyTable()

taylor_cnf_rows = rs("taylor_cnf_students")
taylor_cnf_cols = 5
Dim taylor_cnf_col(5)
taylor_cnf_col(1) = "Banner # "
taylor_cnf_col(2) = "First Name"
taylor_cnf_col(3) = "Last Name"
taylor_cnf_col(4) = "Applied"
taylor_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					taylorcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Taylor University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open taylorcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                        
    '//////// Temple Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_temple_query="SELECT Count(distinct UIN) temple_cnf_students FROM Applicants where UGCollege = 'Temple University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_temple_query,conn
    If rs("temple_cnf_students") <> 0 Then 
    pdf.GreyTitle("Temple University")
'pdf.FancyTable()

temple_cnf_rows = rs("temple_cnf_students")
temple_cnf_cols = 5
Dim temple_cnf_col(5)
temple_cnf_col(1) = "Banner # "
temple_cnf_col(2) = "First Name"
temple_cnf_col(3) = "Last Name"
temple_cnf_col(4) = "Applied"
temple_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					templecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Temple University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open templecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                                                
    '//////// Tennessee State University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_tennessee_query="SELECT Count(distinct UIN) tennessee_cnf_students FROM Applicants where UGCollege = 'Tennessee State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_tennessee_query,conn
    If rs("tennessee_cnf_students") <> 0 Then 
    pdf.GreyTitle("Tennessee State University")
'pdf.FancyTable()

tennessee_cnf_rows = rs("tennessee_cnf_students")
tennessee_cnf_cols = 5
Dim tennessee_cnf_col(5)
tennessee_cnf_col(1) = "Banner # "
tennessee_cnf_col(2) = "First Name"
tennessee_cnf_col(3) = "Last Name"
tennessee_cnf_col(4) = "Applied"
tennessee_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					tennesseecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Tennessee State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open tennesseecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                                                
    '//////// Texas State University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_texasstate_query="SELECT Count(distinct UIN) texasstate_cnf_students FROM Applicants where UGCollege = 'Texas State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_texasstate_query,conn
    If rs("texasstate_cnf_students") <> 0 Then 
    pdf.GreyTitle("Texas State University")
'pdf.FancyTable()

texasstate_cnf_rows = rs("texasstate_cnf_students")
texasstate_cnf_cols = 5
Dim texasstate_cnf_col(5)
texasstate_cnf_col(1) = "Banner # "
texasstate_cnf_col(2) = "First Name"
texasstate_cnf_col(3) = "Last Name"
texasstate_cnf_col(4) = "Applied"
texasstate_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					texasstatecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Texas State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open texasstatecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                         
    '//////// Touro College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_touro_query="SELECT Count(distinct UIN) touro_cnf_students FROM Applicants where UGCollege = 'Touro College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_touro_query,conn
    If rs("touro_cnf_students") <> 0 Then 
    pdf.GreyTitle("Touro College")
'pdf.FancyTable()

touro_cnf_rows = rs("touro_cnf_students")
touro_cnf_cols = 5
Dim touro_cnf_col(5)
touro_cnf_col(1) = "Banner # "
touro_cnf_col(2) = "First Name"
touro_cnf_col(3) = "Last Name"
touro_cnf_col(4) = "Applied"
touro_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					tourocnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Touro College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open tourocnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                        
    '//////// Trinity Christian College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_trinity_query="SELECT Count(distinct UIN) trinity_cnf_students FROM Applicants where UGCollege = 'Trinity Christian College' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_trinity_query,conn
    If rs("trinity_cnf_students") <> 0 Then 
    pdf.GreyTitle("Trinity Christian College")
'pdf.FancyTable()

trinity_cnf_rows = rs("trinity_cnf_students")
trinity_cnf_cols = 5
Dim trinity_cnf_col(5)
trinity_cnf_col(1) = "Banner # "
trinity_cnf_col(2) = "First Name"
trinity_cnf_col(3) = "Last Name"
trinity_cnf_col(4) = "Applied"
trinity_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					trinitycnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Trinity Christian College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open trinitycnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

'//////// Truman State Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_trumanstate_query="SELECT Count(distinct UIN) trumanstate_cnf_students FROM Applicants where UGCollege = 'Truman State University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_trumanstate_query,conn
    If rs("trumanstate_cnf_students") <> 0 Then
    pdf.GreyTitle("Truman State University")
'pdf.FancyTable()



trumanstate_cnf_rows = rs("trumanstate_cnf_students")
trumanstate_cnf_cols = 5
Dim trumanstate_cnf_col(5)
trumanstate_cnf_col(1) = "Banner # "
trumanstate_cnf_col(2) = "First Name"
trumanstate_cnf_col(3) = "Last Name"
trumanstate_cnf_col(4) = "Applied"
trumanstate_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					trumanstatecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Truman State University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open trumanstatecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If   

'//////// Tufts University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_tufts_query="SELECT Count(distinct UIN) tufts_cnf_students FROM Applicants where UGCollege = 'Tufts University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_tufts_query,conn
    If rs("tufts_cnf_students") <> 0 Then
    pdf.GreyTitle("Tufts University")
'pdf.FancyTable()



tufts_cnf_rows = rs("tufts_cnf_students")
tufts_cnf_cols = 5
Dim tufts_cnf_col(5)
tufts_cnf_col(1) = "Banner # "
tufts_cnf_col(2) = "First Name"
tufts_cnf_col(3) = "Last Name"
tufts_cnf_col(4) = "Applied"
tufts_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					tuftscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Tufts University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open tuftscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If   

    '////////  United International College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_united_query="SELECT Count(distinct UIN) united_cnf_students FROM Applicants where UGCollege = 'United International College' and confirmed='Y' and  term_cd='"&AdmitTerm&"'"
rs.Open cnf_students_united_query,conn
If rs("united_cnf_students") <> 0 Then
pdf.GreyTitle("United International College")
'pdf.FancyTable()


united_cnf_rows = rs("united_cnf_students")
united_cnf_cols = 5
Dim united_cnf_col(5)
united_cnf_col(1) = "Banner # "
united_cnf_col(2) = "First Name"
united_cnf_col(3) = "Last Name"
united_cnf_col(4) = "Applied"
united_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					unitedcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'United International College' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open unitedcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
 

    '////////  University of Alabama Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_universityofalabama_query="SELECT Count(distinct UIN) universityofalabama_cnf_students FROM Applicants where UGCollege = 'University of Alabama' and confirmed='Y' and  term_cd='"&AdmitTerm&"'"
rs.Open cnf_students_universityofalabama_query,conn
If rs("universityofalabama_cnf_students") <> 0 Then
pdf.GreyTitle("University of Alabama")
'pdf.FancyTable()


universityofalabama_cnf_rows = rs("universityofalabama_cnf_students")
universityofalabama_cnf_cols = 5
Dim universityofalabama_cnf_col(5)
universityofalabama_cnf_col(1) = "Banner # "
universityofalabama_cnf_col(2) = "First Name"
universityofalabama_cnf_col(3) = "Last Name"
universityofalabama_cnf_col(4) = "Applied"
universityofalabama_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					universityofalabamacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Alabama' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open universityofalabamacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
  
    '////////  University of California Berkeley Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_universityofcaliforniaBerkeley_query="SELECT Count(distinct UIN) universityofcaliforniaBerkeley_cnf_students FROM Applicants where UGCollege = 'University of California Berkeley' and confirmed='Y' and  term_cd='"&AdmitTerm&"'"
rs.Open cnf_students_universityofcaliforniaBerkeley_query,conn
If rs("universityofcaliforniaBerkeley_cnf_students") <> 0 Then
pdf.GreyTitle("University of California Berkeley")
'pdf.FancyTable()


universityofcaliforniaBerkeley_cnf_rows = rs("universityofcaliforniaBerkeley_cnf_students")
universityofcaliforniaBerkeley_cnf_cols = 5
Dim universityofcaliforniaBerkeley_cnf_col(5)
universityofcaliforniaBerkeley_cnf_col(1) = "Banner # "
universityofcaliforniaBerkeley_cnf_col(2) = "First Name"
universityofcaliforniaBerkeley_cnf_col(3) = "Last Name"
universityofcaliforniaBerkeley_cnf_col(4) = "Applied"
universityofcaliforniaBerkeley_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					universityofcaliforniaBerkeleycnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of California Berkeley' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open universityofcaliforniaBerkeleycnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
  
    '////////  University of California Irvine Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_universityofcaliforniairvine_query="SELECT Count(distinct UIN) universityofcaliforniairvine_cnf_students FROM Applicants where UGCollege = 'University of California Irvine' and confirmed='Y' and  term_cd='"&AdmitTerm&"'"
rs.Open cnf_students_universityofcaliforniairvine_query,conn
If rs("universityofcaliforniairvine_cnf_students") <> 0 Then
pdf.GreyTitle("University of California Irvine")
'pdf.FancyTable()


universityofcaliforniairvine_cnf_rows = rs("universityofcaliforniairvine_cnf_students")
universityofcaliforniairvine_cnf_cols = 5
Dim universityofcaliforniairvine_cnf_col(5)
universityofcaliforniairvine_cnf_col(1) = "Banner # "
universityofcaliforniairvine_cnf_col(2) = "First Name"
universityofcaliforniairvine_cnf_col(3) = "Last Name"
universityofcaliforniairvine_cnf_col(4) = "Applied"
universityofcaliforniairvine_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					universityofcaliforniairvinecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of California Irvine' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open universityofcaliforniairvinecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
    
    '////////  University of California San Diego Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_universityofcaliforniasan_query="SELECT Count(distinct UIN) universityofcaliforniasan_cnf_students FROM Applicants where UGCollege = 'University of California San Diego' and confirmed='Y' and  term_cd='"&AdmitTerm&"'"
rs.Open cnf_students_universityofcaliforniasan_query,conn
If rs("universityofcaliforniasan_cnf_students") <> 0 Then
pdf.GreyTitle("University of California San Diego")
'pdf.FancyTable()


universityofcaliforniasan_cnf_rows = rs("universityofcaliforniasan_cnf_students")
universityofcaliforniasan_cnf_cols = 5
Dim universityofcaliforniasan_cnf_col(5)
universityofcaliforniasan_cnf_col(1) = "Banner # "
universityofcaliforniasan_cnf_col(2) = "First Name"
universityofcaliforniasan_cnf_col(3) = "Last Name"
universityofcaliforniasan_cnf_col(4) = "Applied"
universityofcaliforniasan_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					universityofcaliforniasancnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of California San Diego' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open universityofcaliforniasancnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
    
 '////////  University of California Santa Barbara Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_universityofcaliforniabarbara_query="SELECT Count(distinct UIN) universityofcaliforniabarbara_cnf_students FROM Applicants where UGCollege = 'University of California Santa Barbara' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_universityofcaliforniabarbara_query,conn
If rs("universityofcaliforniabarbara_cnf_students") <> 0 Then
pdf.GreyTitle("University of California Santa Barbara")
'pdf.FancyTable()


universityofcaliforniabarbara_cnf_rows = rs("universityofcaliforniabarbara_cnf_students")
universityofcaliforniabarbara_cnf_cols = 5
Dim universityofcaliforniabarbara_cnf_col(5)
universityofcaliforniabarbara_cnf_col(1) = "Banner # "
universityofcaliforniabarbara_cnf_col(2) = "First Name"
universityofcaliforniabarbara_cnf_col(3) = "Last Name"
universityofcaliforniabarbara_cnf_col(4) = "Applied"
universityofcaliforniabarbara_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					universityofcaliforniabarbaracnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of California Santa Barbara' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open universityofcaliforniabarbaracnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
     
 '////////  University of California Santa Cruz Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_universityofcalifornia_query="SELECT Count(distinct UIN) universityofcalifornia_cnf_students FROM Applicants where UGCollege = 'University of California Santa Cruz' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_universityofcalifornia_query,conn
If rs("universityofcalifornia_cnf_students") <> 0 Then
pdf.GreyTitle("University of California Santa Cruz")
'pdf.FancyTable()


universityofcalifornia_cnf_rows = rs("universityofcalifornia_cnf_students")
universityofcalifornia_cnf_cols = 5
Dim universityofcalifornia_cnf_col(5)
universityofcalifornia_cnf_col(1) = "Banner # "
universityofcalifornia_cnf_col(2) = "First Name"
universityofcalifornia_cnf_col(3) = "Last Name"
universityofcalifornia_cnf_col(4) = "Applied"
universityofcalifornia_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					universityofcaliforniacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of California Santa Cruz' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open universityofcaliforniacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
             
'////////  University of Chicago Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_universityofchicago_query="SELECT Count(distinct UIN) universityofchicago_cnf_students FROM Applicants where UGCollege = 'University of Chicago' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_universityofchicago_query,conn
If rs("universityofchicago_cnf_students") <> 0 Then
pdf.GreyTitle("University of Chicago")
'pdf.FancyTable()


universityofchicago_cnf_rows = rs("universityofchicago_cnf_students")
universityofchicago_cnf_cols = 5
Dim universityofchicago_cnf_col(5)
universityofchicago_cnf_col(1) = "Banner # "
universityofchicago_cnf_col(2) = "First Name"
universityofchicago_cnf_col(3) = "Last Name"
universityofchicago_cnf_col(4) = "Applied"
universityofchicago_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					universityofchicagocnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Chicago' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open universityofchicagocnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
 
       
    '////////  University of Cincinnati Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_universityofCincinnati_query="SELECT Count(distinct UIN) universityofCincinnati_cnf_students FROM Applicants where UGCollege = 'University of Cincinnati' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_universityofCincinnati_query,conn
If rs("universityofCincinnati_cnf_students") <> 0 Then
pdf.GreyTitle("University of Cincinnati")
'pdf.FancyTable()


universityofCincinnati_cnf_rows = rs("universityofCincinnati_cnf_students")
universityofCincinnati_cnf_cols = 5
Dim universityofCincinnati_cnf_col(5)
universityofCincinnati_cnf_col(1) = "Banner # "
universityofCincinnati_cnf_col(2) = "First Name"
universityofCincinnati_cnf_col(3) = "Last Name"
universityofCincinnati_cnf_col(4) = "Applied"
universityofCincinnati_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					universityofCincinnaticnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Cincinnati' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open universityofCincinnaticnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn

    '//////// Dayton Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_dayton_query="SELECT Count(distinct UIN) dayton_cnf_students FROM Applicants where UGCollege = 'University of Dayton' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_dayton_query,conn
    If rs("dayton_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Dayton")
'pdf.FancyTable()

dayton_cnf_rows = rs("dayton_cnf_students")
dayton_cnf_cols = 5
Dim dayton_cnf_col(5)
dayton_cnf_col(1) = "Banner # "
dayton_cnf_col(2) = "First Name"
dayton_cnf_col(3) = "Last Name"
dayton_cnf_col(4) = "Applied"
dayton_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					daytoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Dayton' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open daytoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 
 
    '//////// Evansville Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_evansville_query="SELECT Count(distinct UIN) evansville_cnf_students FROM Applicants where UGCollege = 'University of Evansville' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_evansville_query,conn
    If rs("evansville_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Evansville")
'pdf.FancyTable()

evansville_cnf_rows = rs("evansville_cnf_students")
evansville_cnf_cols = 5
Dim evansville_cnf_col(5)
evansville_cnf_col(1) = "Banner # "
evansville_cnf_col(2) = "First Name"
evansville_cnf_col(3) = "Last Name"
evansville_cnf_col(4) = "Applied"
evansville_cnf_col(5) = "Confirmed"

pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					evansvillecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Evansville' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open evansvillecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 

    '//////// University of Georgia Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_univgeorgia_query="SELECT Count(distinct UIN) univgeorgia_cnf_students FROM Applicants where UGCollege = 'University of Georgia' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_univgeorgia_query,conn
    If rs("univgeorgia_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Georgia")
'pdf.FancyTable()

univgeorgia_cnf_rows = rs("univgeorgia_cnf_students")
univgeorgia_cnf_cols = 5
Dim univgeorgia_cnf_col(5)
univgeorgia_cnf_col(1) = "Banner # "
univgeorgia_cnf_col(2) = "First Name"
univgeorgia_cnf_col(3) = "Last Name"
univgeorgia_cnf_col(4) = "Applied"
univgeorgia_cnf_col(5) = "Confirmed"

pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					univgeorgiacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Georgia' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open univgeorgiacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 
                    
'//////// UIC Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_uic_query="SELECT Count(distinct UIN) uic_cnf_students FROM Applicants where UGCollege = 'University of Illinois Chicago' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_uic_query,conn
    If rs("uic_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Illinois Chicago")
'pdf.FancyTable()

uic_cnf_rows = rs("uic_cnf_students")
uic_cnf_cols = 5
Dim uic_cnf_col(5)
uic_cnf_col(1) = "Banner # "
uic_cnf_col(2) = "First Name"
uic_cnf_col(3) = "Last Name"
uic_cnf_col(4) = "Applied"
uic_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					uiccnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Illinois Chicago' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open uiccnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If
                    
'//////// UIS Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_uis_query="SELECT Count(distinct UIN) uis_cnf_students FROM Applicants where UGCollege = 'University of Illinois Springfield' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_uis_query,conn
    If rs("uis_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Illinois Springfield")
'pdf.FancyTable()

uis_cnf_rows = rs("uis_cnf_students")
uis_cnf_cols = 5
Dim uis_cnf_col(5)
uis_cnf_col(1) = "Banner # "
uis_cnf_col(2) = "First Name"
uis_cnf_col(3) = "Last Name"
uis_cnf_col(4) = "Applied"
uis_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					uiscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Illinois Springfield' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open uiscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If

'//////// University of Illinois Urbana Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_universityurbana_query="SELECT Count(distinct UIN) universityurbana_cnf_students FROM Applicants where UGCollege = 'University of Illinois Urbana' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_universityurbana_query,conn
    If rs("universityurbana_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Illinois Urbana")
'pdf.FancyTable()

'//////// Students ////////////

universityurbana_cnf_rows = rs("universityurbana_cnf_students")
universityurbana_cnf_cols = 5
Dim universityurbana_cnf_col(5)
universityurbana_cnf_col(1) = "Banner # "
universityurbana_cnf_col(2) = "First Name"
universityurbana_cnf_col(3) = "Last Name"
universityurbana_cnf_col(4) = "Applied"
universityurbana_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					universityurbanacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Illinois Urbana' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open universityurbanacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                        
    '//////// University of Illinois Urbana Champaign Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_urbanachampaign_query="SELECT Count(distinct UIN) urbanachampaign_cnf_students FROM Applicants where UGCollege = 'University of Illinois Urbana Champaign' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_urbanachampaign_query,conn
    If rs("urbanachampaign_cnf_students") <> 0 Then 
    pdf.GreyTitle("University of Illinois Urbana Champaign")
'pdf.FancyTable()

urbanachampaign_cnf_rows = rs("urbanachampaign_cnf_students")
urbanachampaign_cnf_cols = 5
Dim urbanachampaign_cnf_col(5)
urbanachampaign_cnf_col(1) = "Banner # "
urbanachampaign_cnf_col(2) = "First Name"
urbanachampaign_cnf_col(3) = "Last Name"
urbanachampaign_cnf_col(4) = "Applied"
urbanachampaign_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					urbanachampaigncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Illinois Urbana Champaign' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open urbanachampaigncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
'//////// University of Iowa Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_univiowa_query="SELECT Count(distinct UIN) univiowa_cnf_students FROM Applicants where UGCollege = 'University of Iowa' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_univiowa_query,conn
    If rs("univiowa_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Iowa")
'pdf.FancyTable()



univiowa_cnf_rows = rs("univiowa_cnf_students")
univiowa_cnf_cols = 5
Dim univiowa_cnf_col(5)
univiowa_cnf_col(1) = "Banner # "
univiowa_cnf_col(2) = "First Name"
univiowa_cnf_col(3) = "Last Name"
univiowa_cnf_col(4) = "Applied"
univiowa_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					univiowacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Iowa' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open univiowacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If  
'//////// University of Kansas Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_univkansas_query="SELECT Count(distinct UIN) univkansas_cnf_students FROM Applicants where UGCollege = 'University of Kansas' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_univkansas_query,conn
    If rs("univkansas_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Kansas")
'pdf.FancyTable()



univkansas_cnf_rows = rs("univkansas_cnf_students")
univkansas_cnf_cols = 5
Dim univkansas_cnf_col(5)
univkansas_cnf_col(1) = "Banner # "
univkansas_cnf_col(2) = "First Name"
univkansas_cnf_col(3) = "Last Name"
univkansas_cnf_col(4) = "Applied"
univkansas_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					univkansascnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Kansas' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open univkansascnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If                

    '//////// University of Massachusetts Amherst Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_massachusetts_query="SELECT Count(distinct UIN) massachusetts_cnf_students FROM Applicants where UGCollege = 'University of Massachusetts Amherst' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_massachusetts_query,conn
    If rs("massachusetts_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Massachusetts Amherst")
'pdf.FancyTable()

'//////// Students ////////////

massachusetts_cnf_rows = rs("massachusetts_cnf_students")
massachusetts_cnf_cols = 5
Dim massachusetts_cnf_col(5)
massachusetts_cnf_col(1) = "Banner # "
massachusetts_cnf_col(2) = "First Name"
massachusetts_cnf_col(3) = "Last Name"
massachusetts_cnf_col(4) = "Applied"
massachusetts_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					massachusettscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Massachusetts Amherst' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open massachusettscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                        

    '//////// University of Michigan Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_univmichigan_query="SELECT Count(distinct UIN) univmichigan_cnf_students FROM Applicants where UGCollege = 'University of Michigan' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_univmichigan_query,conn
    If rs("univmichigan_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Michigan")
'pdf.FancyTable()

'//////// Students ////////////

univmichigan_cnf_rows = rs("univmichigan_cnf_students")
univmichigan_cnf_cols = 5
Dim univmichigan_cnf_col(5)
univmichigan_cnf_col(1) = "Banner # "
univmichigan_cnf_col(2) = "First Name"
univmichigan_cnf_col(3) = "Last Name"
univmichigan_cnf_col(4) = "Applied"
univmichigan_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					univmichigancnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Michigan' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open univmichigancnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  

    '//////// University of Michigan Ann Arbor Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_univmichiganann_query="SELECT Count(distinct UIN) univmichiganann_cnf_students FROM Applicants where UGCollege = 'University of Michigan Ann Arbor' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_univmichiganann_query,conn
    If rs("univmichiganann_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Michigan Ann Arbor")
'pdf.FancyTable()

'//////// Students ////////////

univmichiganann_cnf_rows = rs("univmichiganann_cnf_students")
univmichiganann_cnf_cols = 5
Dim univmichiganann_cnf_col(5)
univmichiganann_cnf_col(1) = "Banner # "
univmichiganann_cnf_col(2) = "First Name"
univmichiganann_cnf_col(3) = "Last Name"
univmichiganann_cnf_col(4) = "Applied"
univmichiganann_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					univmichigananncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Michigan Ann Arbor' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open univmichigananncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                        
    '//////// University of Minnesota Twin Cities Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_minnesota_query="SELECT Count(distinct UIN) minnesota_cnf_students FROM Applicants where UGCollege = 'University of Minnesota Twin Cities' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_minnesota_query,conn
    If rs("minnesota_cnf_students") <> 0 Then 
    pdf.GreyTitle("University of Minnesota Twin Cities")
'pdf.FancyTable()

minnesota_cnf_rows = rs("minnesota_cnf_students")
minnesota_cnf_cols = 5
Dim minnesota_cnf_col(5)
minnesota_cnf_col(1) = "Banner # "
minnesota_cnf_col(2) = "First Name"
minnesota_cnf_col(3) = "Last Name"
minnesota_cnf_col(4) = "Applied"
minnesota_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					minnesotacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Minnesota Twin Cities' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open minnesotacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
  
'//////// University of Missouri - St. Louis Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_missouristlouis_query="SELECT Count(distinct UIN) missouristlouis_cnf_students FROM Applicants where UGCollege = 'University of Missouri - St. Louis' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_missouristlouis_query,conn
    If rs("missouristlouis_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Missouri - St. Louis")
'pdf.FancyTable()



missouristlouis_cnf_rows = rs("missouristlouis_cnf_students")
missouristlouis_cnf_cols = 5
Dim missouristlouis_cnf_col(5)
missouristlouis_cnf_col(1) = "Banner # "
missouristlouis_cnf_col(2) = "First Name"
missouristlouis_cnf_col(3) = "Last Name"
missouristlouis_cnf_col(4) = "Applied"
missouristlouis_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					missouristlouiscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Missouri - St. Louis' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open missouristlouiscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    
   
'//////// University of Missouri Columbia Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_missouri_query="SELECT Count(distinct UIN) missouri_cnf_students FROM Applicants where UGCollege = 'University of Missouri Columbia' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_missouri_query,conn
    If rs("missouri_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Missouri Columbia")
'pdf.FancyTable()



missouri_cnf_rows = rs("missouri_cnf_students")
missouri_cnf_cols = 5
Dim missouri_cnf_col(5)
missouri_cnf_col(1) = "Banner # "
missouri_cnf_col(2) = "First Name"
missouri_cnf_col(3) = "Last Name"
missouri_cnf_col(4) = "Applied"
missouri_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					missouricnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Missouri Columbia' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open missouricnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    
  
'//////// University of Nebraska Lincoln Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_nebraskalincoln_query="SELECT Count(distinct UIN) nebraskalincoln_cnf_students FROM Applicants where UGCollege = 'University of Nebraska Lincoln' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_nebraskalincoln_query,conn
    If rs("nebraskalincoln_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Nebraska Lincoln")
'pdf.FancyTable()



nebraskalincoln_cnf_rows = rs("nebraskalincoln_cnf_students")
nebraskalincoln_cnf_cols = 5
Dim nebraskalincoln_cnf_col(5)
nebraskalincoln_cnf_col(1) = "Banner # "
nebraskalincoln_cnf_col(2) = "First Name"
nebraskalincoln_cnf_col(3) = "Last Name"
nebraskalincoln_cnf_col(4) = "Applied"
nebraskalincoln_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					nebraskalincolncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Nebraska Lincoln' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open nebraskalincolncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    

 '//////// University of New Hampshire Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_newhampshire_query="SELECT Count(distinct UIN) newhampshire_cnf_students FROM Applicants where UGCollege = 'University of New Hampshire' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_newhampshire_query,conn
    If rs("newhampshire_cnf_students") <> 0 Then
    pdf.GreyTitle("University of New Hampshire")
'pdf.FancyTable()

newhampshire_cnf_rows = rs("newhampshire_cnf_students")
newhampshire_cnf_cols = 5
Dim newhampshire_cnf_col(5)
newhampshire_cnf_col(1) = "Banner # "
newhampshire_cnf_col(2) = "First Name"
newhampshire_cnf_col(3) = "Last Name"
newhampshire_cnf_col(4) = "Applied"
newhampshire_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					newhampshirecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of New Hampshire' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open newhampshirecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    

 '//////// University of Nigeria Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_nigeria_query="SELECT Count(distinct UIN) nigeria_cnf_students FROM Applicants where UGCollege = 'University of Nigeria' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_nigeria_query,conn
    If rs("nigeria_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Nigeria")
'pdf.FancyTable()

nigeria_cnf_rows = rs("nigeria_cnf_students")
nigeria_cnf_cols = 5
Dim nigeria_cnf_col(5)
nigeria_cnf_col(1) = "Banner # "
nigeria_cnf_col(2) = "First Name"
nigeria_cnf_col(3) = "Last Name"
nigeria_cnf_col(4) = "Applied"
nigeria_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					nigeriacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Nigeria' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open nigeriacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    
  
'//////// University of Norte Dame Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_nortedame_query="SELECT Count(distinct UIN) nortedame_cnf_students FROM Applicants where UGCollege = 'University of Norte Dame' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_nortedame_query,conn
    If rs("nortedame_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Norte Dame")
'pdf.FancyTable()



nortedame_cnf_rows = rs("nortedame_cnf_students")
nortedame_cnf_cols = 5
Dim nortedame_cnf_col(5)
nortedame_cnf_col(1) = "Banner # "
nortedame_cnf_col(2) = "First Name"
nortedame_cnf_col(3) = "Last Name"
nortedame_cnf_col(4) = "Applied"
nortedame_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					nortedamecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Norte Dame' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open nortedamecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    
   
'//////// University of North Carolina Chapel Hill Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_uncch_query="SELECT Count(distinct UIN) uncch_cnf_students FROM Applicants where UGCollege = 'University of North Carolina Chapel Hill' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_uncch_query,conn
    If rs("uncch_cnf_students") <> 0 Then
    pdf.GreyTitle("University of North Carolina Chapel Hill")
'pdf.FancyTable()

uncch_cnf_rows = rs("uncch_cnf_students")
uncch_cnf_cols = 5
Dim uncch_cnf_col(5)
uncch_cnf_col(1) = "Banner # "
uncch_cnf_col(2) = "First Name"
uncch_cnf_col(3) = "Last Name"
uncch_cnf_col(4) = "Applied"
uncch_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					uncchcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of North Carolina Chapel Hill' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open uncchcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    
    
'//////// University of North Carolina Wilmington Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_wilmington_query="SELECT Count(distinct UIN) wilmington_cnf_students FROM Applicants where UGCollege = 'University of North Carolina Wilmington' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_wilmington_query,conn
    If rs("wilmington_cnf_students") <> 0 Then
    pdf.GreyTitle("University of North Carolina Wilmington")
'pdf.FancyTable()

wilmington_cnf_rows = rs("wilmington_cnf_students")
wilmington_cnf_cols = 5
Dim wilmington_cnf_col(5)
wilmington_cnf_col(1) = "Banner # "
wilmington_cnf_col(2) = "First Name"
wilmington_cnf_col(3) = "Last Name"
wilmington_cnf_col(4) = "Applied"
wilmington_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					wilmingtoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of North Carolina Wilmington' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open wilmingtoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    
 
'//////// University of North Dakota Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_northdakota_query="SELECT Count(distinct UIN) northdakota_cnf_students FROM Applicants where UGCollege = 'University of North Dakota' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_northdakota_query,conn
    If rs("northdakota_cnf_students") <> 0 Then
    pdf.GreyTitle("University of North Dakota")
'pdf.FancyTable()
northdakota_cnf_rows = rs("northdakota_cnf_students")
northdakota_cnf_cols = 5
Dim northdakota_cnf_col(5)
northdakota_cnf_col(1) = "Banner # "
northdakota_cnf_col(2) = "First Name"
northdakota_cnf_col(3) = "Last Name"
northdakota_cnf_col(4) = "Applied"
northdakota_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					northdakotacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of North Dakota' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open northdakotacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    
 
'//////// University of Oregon Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_oregon_query="SELECT Count(distinct UIN) oregon_cnf_students FROM Applicants where UGCollege = 'University of Oregon' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_oregon_query,conn
    If rs("oregon_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Oregon")
'pdf.FancyTable()
oregon_cnf_rows = rs("oregon_cnf_students")
oregon_cnf_cols = 5
Dim oregon_cnf_col(5)
oregon_cnf_col(1) = "Banner # "
oregon_cnf_col(2) = "First Name"
oregon_cnf_col(3) = "Last Name"
oregon_cnf_col(4) = "Applied"
oregon_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					oregoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Oregon' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open oregoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    
   
'//////// University of Pittsburgh Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_pittsburgh_query="SELECT Count(distinct UIN) pittsburgh_cnf_students FROM Applicants where UGCollege = 'University of Pittsburgh' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_pittsburgh_query,conn
    If rs("pittsburgh_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Pittsburgh")
'pdf.FancyTable()



pittsburgh_cnf_rows = rs("pittsburgh_cnf_students")
pittsburgh_cnf_cols = 5
Dim pittsburgh_cnf_col(5)
pittsburgh_cnf_col(1) = "Banner # "
pittsburgh_cnf_col(2) = "First Name"
pittsburgh_cnf_col(3) = "Last Name"
pittsburgh_cnf_col(4) = "Applied"
pittsburgh_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					pittsburghcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Pittsburgh' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open pittsburghcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    
  
'//////// University of Redlands Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_redlands_query="SELECT Count(distinct UIN) redlands_cnf_students FROM Applicants where UGCollege = 'University of Redlands' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_redlands_query,conn
    If rs("redlands_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Redlands")
'pdf.FancyTable()
redlands_cnf_rows = rs("redlands_cnf_students")
redlands_cnf_cols = 5
Dim redlands_cnf_col(5)
redlands_cnf_col(1) = "Banner # "
redlands_cnf_col(2) = "First Name"
redlands_cnf_col(3) = "Last Name"
redlands_cnf_col(4) = "Applied"
redlands_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					redlandscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Redlands' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open redlandscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    
 
'//////// University of Rochester Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_rochester_query="SELECT Count(distinct UIN) rochester_cnf_students FROM Applicants where UGCollege = 'University of Rochester' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_rochester_query,conn
    If rs("rochester_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Rochester")
'pdf.FancyTable()



rochester_cnf_rows = rs("rochester_cnf_students")
rochester_cnf_cols = 5
Dim rochester_cnf_col(5)
rochester_cnf_col(1) = "Banner # "
rochester_cnf_col(2) = "First Name"
rochester_cnf_col(3) = "Last Name"
rochester_cnf_col(4) = "Applied"
rochester_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					rochestercnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Rochester' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open rochestercnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    
  
'//////// University of Scranton Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_Scranton_query="SELECT Count(distinct UIN) Scranton_cnf_students FROM Applicants where UGCollege = 'University of Scranton' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_Scranton_query,conn
    If rs("Scranton_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Scranton")
'pdf.FancyTable()



Scranton_cnf_rows = rs("Scranton_cnf_students")
Scranton_cnf_cols = 5
Dim Scranton_cnf_col(5)
Scranton_cnf_col(1) = "Banner # "
Scranton_cnf_col(2) = "First Name"
Scranton_cnf_col(3) = "Last Name"
Scranton_cnf_col(4) = "Applied"
Scranton_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					Scrantoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Scranton' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open Scrantoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If    
            
   
            
'////////  University of South Carolina Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_southcarolina_query="SELECT Count(distinct UIN) southcarolina_cnf_students FROM Applicants where UGCollege = 'University of South Carolina' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_southcarolina_query,conn
If rs("southcarolina_cnf_students") <> 0 Then
pdf.GreyTitle("University of South Carolina")
'pdf.FancyTable()


southcarolina_cnf_rows = rs("southcarolina_cnf_students")
southcarolina_cnf_cols = 5
Dim southcarolina_cnf_col(5)
southcarolina_cnf_col(1) = "Banner # "
southcarolina_cnf_col(2) = "First Name"
southcarolina_cnf_col(3) = "Last Name"
southcarolina_cnf_col(4) = "Applied"
southcarolina_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					southcarolinacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of South Carolina' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open southcarolinacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
    
            
'////////  University of South Carolina Columbia Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_southcarolinaColumbia_query="SELECT Count(distinct UIN) southcarolinaColumbia_cnf_students FROM Applicants where UGCollege = 'University of South Carolina Columbia' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_southcarolinaColumbia_query,conn
If rs("southcarolinaColumbia_cnf_students") <> 0 Then
pdf.GreyTitle("University of South Carolina Columbia")
'pdf.FancyTable()


southcarolinaColumbia_cnf_rows = rs("southcarolinaColumbia_cnf_students")
southcarolinaColumbia_cnf_cols = 5
Dim southcarolinaColumbia_cnf_col(5)
southcarolinaColumbia_cnf_col(1) = "Banner # "
southcarolinaColumbia_cnf_col(2) = "First Name"
southcarolinaColumbia_cnf_col(3) = "Last Name"
southcarolinaColumbia_cnf_col(4) = "Applied"
southcarolinaColumbia_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					southcarolinaColumbiacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of South Carolina Columbia' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open southcarolinaColumbiacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
           
'////////  University of Texas Austin Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_austin_query="SELECT Count(distinct UIN) austin_cnf_students FROM Applicants where UGCollege = 'University of Texas Austin' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_austin_query,conn
If rs("austin_cnf_students") <> 0 Then
pdf.GreyTitle("University of Texas Austin")
'pdf.FancyTable()


austin_cnf_rows = rs("austin_cnf_students")
austin_cnf_cols = 5
Dim austin_cnf_col(5)
austin_cnf_col(1) = "Banner # "
austin_cnf_col(2) = "First Name"
austin_cnf_col(3) = "Last Name"
austin_cnf_col(4) = "Applied"
austin_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					austincnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Texas Austin' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open austincnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
             
'////////  University of Toledo Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_toledo_query="SELECT Count(distinct UIN) toledo_cnf_students FROM Applicants where UGCollege = 'University of Toledo' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_toledo_query,conn
If rs("toledo_cnf_students") <> 0 Then
pdf.GreyTitle("University of Toledo")
'pdf.FancyTable()


toledo_cnf_rows = rs("toledo_cnf_students")
toledo_cnf_cols = 5
Dim toledo_cnf_col(5)
toledo_cnf_col(1) = "Banner # "
toledo_cnf_col(2) = "First Name"
toledo_cnf_col(3) = "Last Name"
toledo_cnf_col(4) = "Applied"
toledo_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					toledocnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Toledo' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open toledocnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
    
            
'////////  University of Vermont Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_vermont_query="SELECT Count(distinct UIN) vermont_cnf_students FROM Applicants where UGCollege = 'University of Vermont' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_vermont_query,conn
If rs("vermont_cnf_students") <> 0 Then
pdf.GreyTitle("University of Vermont")
'pdf.FancyTable()


vermont_cnf_rows = rs("vermont_cnf_students")
vermont_cnf_cols = 5
Dim vermont_cnf_col(5)
vermont_cnf_col(1) = "Banner # "
vermont_cnf_col(2) = "First Name"
vermont_cnf_col(3) = "Last Name"
vermont_cnf_col(4) = "Applied"
vermont_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					vermontcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Vermont' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open vermontcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
 
'//////// University of Washington Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_washington_query="SELECT Count(distinct UIN) washington_cnf_students FROM Applicants where UGCollege = 'University of Washington' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_washington_query,conn
    If rs("washington_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Washington")
'pdf.FancyTable()

'//////// Students ////////////

washington_cnf_rows = rs("washington_cnf_students")
washington_cnf_cols = 5
Dim washington_cnf_col(5)
washington_cnf_col(1) = "Banner # "
washington_cnf_col(2) = "First Name"
washington_cnf_col(3) = "Last Name"
washington_cnf_col(4) = "Applied"
washington_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					washingtoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Washington' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open washingtoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
 
'//////// University of Washington Seattle Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_washingtonseattle_query="SELECT Count(distinct UIN) washingtonseattle_cnf_students FROM Applicants where UGCollege = 'University of Washington Seattle' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_washingtonseattle_query,conn
    If rs("washingtonseattle_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Washington Seattle")
'pdf.FancyTable()

'//////// Students ////////////

washingtonseattle_cnf_rows = rs("washingtonseattle_cnf_students")
washingtonseattle_cnf_cols = 5
Dim washingtonseattle_cnf_col(5)
washingtonseattle_cnf_col(1) = "Banner # "
washingtonseattle_cnf_col(2) = "First Name"
washingtonseattle_cnf_col(3) = "Last Name"
washingtonseattle_cnf_col(4) = "Applied"
washingtonseattle_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					washingtonseattlecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Washington Seattle' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open washingtonseattlecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
  
'//////// University of Wisconsin Eau Claire Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_wisconsineauclaire_query="SELECT Count(distinct UIN) wisconsineauclaire_cnf_students FROM Applicants where UGCollege = 'University of Wisconsin Eau Claire' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_wisconsineauclaire_query,conn
    If rs("wisconsineauclaire_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Wisconsin Eau Claire")
'pdf.FancyTable()

'//////// Students ////////////

wisconsineauclaire_cnf_rows = rs("wisconsineauclaire_cnf_students")
wisconsineauclaire_cnf_cols = 5
Dim wisconsineauclaire_cnf_col(5)
wisconsineauclaire_cnf_col(1) = "Banner # "
wisconsineauclaire_cnf_col(2) = "First Name"
wisconsineauclaire_cnf_col(3) = "Last Name"
wisconsineauclaire_cnf_col(4) = "Applied"
wisconsineauclaire_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					wisconsineauclairecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Wisconsin Eau Claire' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open wisconsineauclairecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
                        
    '//////// University of Wisconsin La Crosse Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_wisconsinlacrosse_query="SELECT Count(distinct UIN) wisconsinlacrosse_cnf_students FROM Applicants where UGCollege = 'University of Wisconsin La Crosse' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_wisconsinlacrosse_query,conn
    If rs("wisconsinlacrosse_cnf_students") <> 0 Then 
    pdf.GreyTitle("University of Wisconsin La Crosse")
'pdf.FancyTable()

wisconsinlacrosse_cnf_rows = rs("wisconsinlacrosse_cnf_students")
wisconsinlacrosse_cnf_cols = 5
Dim wisconsinlacrosse_cnf_col(5)
wisconsinlacrosse_cnf_col(1) = "Banner # "
wisconsinlacrosse_cnf_col(2) = "First Name"
wisconsinlacrosse_cnf_col(3) = "Last Name"
wisconsinlacrosse_cnf_col(4) = "Applied"
wisconsinlacrosse_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					wisconsinlacrossecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Wisconsin La Crosse' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open wisconsinlacrossecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
                              
    '//////// University of Wisconsin Madison Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_wisconsinmadison_query="SELECT Count(distinct UIN) wisconsinmadison_cnf_students FROM Applicants where UGCollege = 'University of Wisconsin Madison' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_wisconsinmadison_query,conn
    If rs("wisconsinmadison_cnf_students") <> 0 Then 
    pdf.GreyTitle("University of Wisconsin Madison")
'pdf.FancyTable()

wisconsinmadison_cnf_rows = rs("wisconsinmadison_cnf_students")
wisconsinmadison_cnf_cols = 5
Dim wisconsinmadison_cnf_col(5)
wisconsinmadison_cnf_col(1) = "Banner # "
wisconsinmadison_cnf_col(2) = "First Name"
wisconsinmadison_cnf_col(3) = "Last Name"
wisconsinmadison_cnf_col(4) = "Applied"
wisconsinmadison_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					wisconsinmadisoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Wisconsin Madison' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open wisconsinmadisoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
  
'//////// University of Wisconsin Milwaukee Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_wisconsinmilwaukee_query="SELECT Count(distinct UIN) wisconsinmilwaukee_cnf_students FROM Applicants where UGCollege = 'University of Wisconsin Milwaukee' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_wisconsinmilwaukee_query,conn
    If rs("wisconsinmilwaukee_cnf_students") <> 0 Then
    pdf.GreyTitle("University of Wisconsin Milwaukee")
'pdf.FancyTable()

'//////// Students ////////////

wisconsinmilwaukee_cnf_rows = rs("wisconsinmilwaukee_cnf_students")
wisconsinmilwaukee_cnf_cols = 5
Dim wisconsinmilwaukee_cnf_col(5)
wisconsinmilwaukee_cnf_col(1) = "Banner # "
wisconsinmilwaukee_cnf_col(2) = "First Name"
wisconsinmilwaukee_cnf_col(3) = "Last Name"
wisconsinmilwaukee_cnf_col(4) = "Applied"
wisconsinmilwaukee_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					wisconsinmilwaukeecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'University of Wisconsin Milwaukee' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open wisconsinmilwaukeecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
   
'//////// Upper Iowa University Students ////////////

    set rs=Server.CreateObject("ADODB.recordset")
cnf_students_upperiowa_query="SELECT Count(distinct UIN) upperiowa_cnf_students FROM Applicants where UGCollege = 'Upper Iowa University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_upperiowa_query,conn
    If rs("upperiowa_cnf_students") <> 0 Then
    pdf.GreyTitle("Upper Iowa University")
'pdf.FancyTable()

'//////// Students ////////////

upperiowa_cnf_rows = rs("upperiowa_cnf_students")
upperiowa_cnf_cols = 5
Dim upperiowa_cnf_col(5)
upperiowa_cnf_col(1) = "Banner # "
upperiowa_cnf_col(2) = "First Name"
upperiowa_cnf_col(3) = "Last Name"
upperiowa_cnf_col(4) = "Applied"
upperiowa_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					upperiowacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Upper Iowa University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open upperiowacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
  Else
    rs.close
    End If  
    
    '//////// Valparaiso University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_valparaiso_query="SELECT Count(distinct UIN) valparaiso_cnf_students FROM Applicants where UGCollege = 'Valparaiso University' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_valparaiso_query,conn
    If rs("valparaiso_cnf_students") <> 0 Then 
    pdf.GreyTitle("Valparaiso University")
'pdf.FancyTable()

valparaiso_cnf_rows = rs("valparaiso_cnf_students")
valparaiso_cnf_cols = 5
Dim valparaiso_cnf_col(5)
valparaiso_cnf_col(1) = "Banner # "
valparaiso_cnf_col(2) = "First Name"
valparaiso_cnf_col(3) = "Last Name"
valparaiso_cnf_col(4) = "Applied"
valparaiso_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					valparaisocnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Valparaiso University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open valparaisocnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
  
    '//////// Villanova University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_villanova_query="SELECT Count(distinct UIN) villanova_cnf_students FROM Applicants where UGCollege = 'Villanova University' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_villanova_query,conn
    If rs("villanova_cnf_students") <> 0 Then 
    pdf.GreyTitle("Villanova University")
'pdf.FancyTable()

villanova_cnf_rows = rs("villanova_cnf_students")
villanova_cnf_cols = 5
Dim villanova_cnf_col(5)
villanova_cnf_col(1) = "Banner # "
villanova_cnf_col(2) = "First Name"
villanova_cnf_col(3) = "Last Name"
villanova_cnf_col(4) = "Applied"
villanova_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					villanovacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Villanova University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open villanovacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn
  
    '//////// Wesleyan University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_wesleyanuniv_query="SELECT Count(distinct UIN) wesleyanuniv_cnf_students FROM Applicants where UGCollege = 'Wesleyan University' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_wesleyanuniv_query,conn
    If rs("wesleyanuniv_cnf_students") <> 0 Then 
    pdf.GreyTitle("Wesleyan University")
'pdf.FancyTable()

wesleyanuniv_cnf_rows = rs("wesleyanuniv_cnf_students")
wesleyanuniv_cnf_cols = 5
Dim wesleyanuniv_cnf_col(5)
wesleyanuniv_cnf_col(1) = "Banner # "
wesleyanuniv_cnf_col(2) = "First Name"
wesleyanuniv_cnf_col(3) = "Last Name"
wesleyanuniv_cnf_col(4) = "Applied"
wesleyanuniv_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					wesleyanunivcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Wesleyan University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open wesleyanunivcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn



    '//////// West Chester University of Pennsylvania Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_westchester_query="SELECT Count(distinct UIN) westchester_cnf_students FROM Applicants where UGCollege = 'West Chester University of Pennsylvania' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_westchester_query,conn
    If rs("westchester_cnf_students") <> 0 Then 
    pdf.GreyTitle("West Chester University of Pennsylvania")
'pdf.FancyTable()

westchester_cnf_rows = rs("westchester_cnf_students")
westchester_cnf_cols = 5
Dim westchester_cnf_col(5)
westchester_cnf_col(1) = "Banner # "
westchester_cnf_col(2) = "First Name"
westchester_cnf_col(3) = "Last Name"
westchester_cnf_col(4) = "Applied"
westchester_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					westchestercnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'West Chester University of Pennsylvania' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open westchestercnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

    '//////// Western Illinois University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_westernillinois_query="SELECT Count(distinct UIN) westernillinois_cnf_students FROM Applicants where UGCollege = 'Western Illinois University' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_westernillinois_query,conn
    If rs("westernillinois_cnf_students") <> 0 Then 
    pdf.GreyTitle("Western Illinois University")
'pdf.FancyTable()

westernillinois_cnf_rows = rs("westernillinois_cnf_students")
westernillinois_cnf_cols = 5
Dim westernillinois_cnf_col(5)
westernillinois_cnf_col(1) = "Banner # "
westernillinois_cnf_col(2) = "First Name"
westernillinois_cnf_col(3) = "Last Name"
westernillinois_cnf_col(4) = "Applied"
westernillinois_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					westernillinoiscnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Western Illinois University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open westernillinoiscnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn


'//////// Western Michigan University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_westernmichigan_query="SELECT Count(distinct UIN) westernmichigan_cnf_students FROM Applicants where UGCollege = 'Western Michigan University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_westernmichigan_query,conn
    If rs("westernmichigan_cnf_students") <> 0 Then
    pdf.GreyTitle("Western Michigan University")
'pdf.FancyTable()



westernmichigan_cnf_rows = rs("westernmichigan_cnf_students")
westernmichigan_cnf_cols = 5
Dim westernmichigan_cnf_col(5)
westernmichigan_cnf_col(1) = "Banner # "
westernmichigan_cnf_col(2) = "First Name"
westernmichigan_cnf_col(3) = "Last Name"
westernmichigan_cnf_col(4) = "Applied"
westernmichigan_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					westernmichigancnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Western Michigan University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open westernmichigancnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 

'//////// Western Washington University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_westernwashington_query="SELECT Count(distinct UIN) westernwashington_cnf_students FROM Applicants where UGCollege = 'Western Washington University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_westernwashington_query,conn
    If rs("westernwashington_cnf_students") <> 0 Then
    pdf.GreyTitle("Western Washington University")
'pdf.FancyTable()



westernwashington_cnf_rows = rs("westernwashington_cnf_students")
westernwashington_cnf_cols = 5
Dim westernwashington_cnf_col(5)
westernwashington_cnf_col(1) = "Banner # "
westernwashington_cnf_col(2) = "First Name"
westernwashington_cnf_col(3) = "Last Name"
westernwashington_cnf_col(4) = "Applied"
westernwashington_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					westernwashingtoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Western Washington University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open westernwashingtoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close    
    pdf.Ln(5)
    Else
    rs.close
    End If 
    
 
    '//////// Westwood College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_westwood_query="SELECT Count(distinct UIN) westwood_cnf_students FROM Applicants where UGCollege = 'Westwood College' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_westwood_query,conn
    If rs("westwood_cnf_students") <> 0 Then 
    pdf.GreyTitle("Westwood College")
'pdf.FancyTable()

westwood_cnf_rows = rs("westwood_cnf_students")
westwood_cnf_cols = 5
Dim westwood_cnf_col(5)
westwood_cnf_col(1) = "Banner # "
westwood_cnf_col(2) = "First Name"
westwood_cnf_col(3) = "Last Name"
westwood_cnf_col(4) = "Applied"
westwood_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					westwoodcnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Westwood College' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open westwoodcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

    '//////// Wheaton College Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_wheaton_query="SELECT Count(distinct UIN) wheaton_cnf_students FROM Applicants where UGCollege = 'Wheaton College' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_wheaton_query,conn
    If rs("wheaton_cnf_students") <> 0 Then 
    pdf.GreyTitle("Wheaton College")
'pdf.FancyTable()

wheaton_cnf_rows = rs("wheaton_cnf_students")
wheaton_cnf_cols = 5
Dim wheaton_cnf_col(5)
wheaton_cnf_col(1) = "Banner # "
wheaton_cnf_col(2) = "First Name"
wheaton_cnf_col(3) = "Last Name"
wheaton_cnf_col(4) = "Applied"
wheaton_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					wheatoncnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Wheaton College' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open wheatoncnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

    '//////// Winona State University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_winona_query="SELECT Count(distinct UIN) winona_cnf_students FROM Applicants where UGCollege = 'Winona State University' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_winona_query,conn
    If rs("winona_cnf_students") <> 0 Then 
    pdf.GreyTitle("Winona State University")
'pdf.FancyTable()

winona_cnf_rows = rs("winona_cnf_students")
winona_cnf_cols = 5
Dim winona_cnf_col(5)
winona_cnf_col(1) = "Banner # "
winona_cnf_col(2) = "First Name"
winona_cnf_col(3) = "Last Name"
winona_cnf_col(4) = "Applied"
winona_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					winonacnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Winona State University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open winonacnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

    '////////   Xavier Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_xavier_query="SELECT Count(distinct UIN) xavier_cnf_students FROM Applicants where UGCollege = 'Xavier University' and confirmed='Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_xavier_query,conn
If rs("xavier_cnf_students") <> 0 Then
pdf.GreyTitle("Xavier University")
'pdf.FancyTable()


xavier_cnf_rows = rs("xavier_cnf_students")
xavier_cnf_cols = 5
Dim xavier_cnf_col(5)
xavier_cnf_col(1) = "Banner # "
xavier_cnf_col(2) = "First Name"
xavier_cnf_col(3) = "Last Name"
xavier_cnf_col(4) = "Applied"
xavier_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					xaviercnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Xavier University' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open xaviercnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
    pdf.Ln(5)
    Else
    rs.close
    End If
'rs.Open query,conn
    '//////// Yale University Students ////////////
set rs=Server.CreateObject("ADODB.recordset")
cnf_students_yale_query="SELECT Count(distinct UIN) yale_cnf_students FROM Applicants where UGCollege = 'Yale University' and confirmed = 'Y' and  Term_CD like '"&AdmitTerm&"' "
rs.Open cnf_students_yale_query,conn
    If rs("yale_cnf_students") <> 0 Then 
    pdf.GreyTitle("Yale University")
'pdf.FancyTable()

yale_cnf_rows = rs("yale_cnf_students")
yale_cnf_cols = 5
Dim yale_cnf_col(5)
yale_cnf_col(1) = "Banner # "
yale_cnf_col(2) = "First Name"
yale_cnf_col(3) = "Last Name"
yale_cnf_col(4) = "Applied"
yale_cnf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,65,60,20,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name","Applied", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					yalecnf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where UGCollege = 'Yale University' and confirmed = 'Y' and term_cd='"&AdmitTerm&"'"
					rs.Open yalecnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace("Yes","|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close 
    pdf.Ln(5)
    Else
    rs.close
    End If  
'rs.Open query,conn

           

    


                      
pdf.Ln(5)



pdf.Ln(10)
pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
