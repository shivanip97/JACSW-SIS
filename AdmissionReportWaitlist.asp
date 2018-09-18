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
pdf.SetFont "Arial","",16
pdf.Open()
pdf.LoadModels("TestModels") 
pdf.AddPage()

pdf.ChapterTitle2("                   Report 12- Admissions Report - "&Termsel& " - WaitList      "  &LastUpdatedDt&" "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(10)

rs.close


'////// Adv Students ////////
pdf.OrangeTitle("Program Option ADV")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// Adv CHF ////////////



set rs=Server.CreateObject("ADODB.recordset")
adv_chf_query="SELECT Count(distinct UIN) adv_chf_students FROM Applicants where Program_type='Adv' and Concentration='CHF' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open adv_chf_query,conn

adv_chf_rows = rs("adv_chf_students")
adv_chf_cols = 3
Dim adv_chf_col(3)
adv_chf_col(1) = "Banner # "
adv_chf_col(2) = "First Name"
adv_chf_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					advchf_query="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='Adv' and Concentration='CHF' and Admission_decision = 'W' and term_cd='"&AdmitTerm&"'"
					rs.Open advchf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_Adv_CHF_Student = "Total number of students in CHF : "&adv_chf_rows
    pdf.ChapterBody(Total_Adv_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
advchudquery="SELECT Count(distinct UIN) adv_chud_students FROM Applicants where Program_type='Adv' and Concentration='CHUD' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open advchudquery,conn
'//////// Courses Table ////////////
advchud_rows = rs("adv_chud_students")
adv_chud_cols = 3
Dim adv_chud_col(3)
adv_chud_col(1) = "Banner # "
adv_chud_col(2) = "First Name"
adv_chud_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					adv_chud="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='Adv' and Concentration='CHUD' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open adv_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    
    pdf.Ln(1)
    rs.close
    Totaladvchud_Student = "Total number of students in CHUD : "&advchud_rows
    pdf.ChapterBody(Totaladvchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
advmhquery="SELECT Count(distinct UIN) adv_mh_students FROM Applicants where Program_type='Adv' and Concentration='MH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open advmhquery,conn
advmh_rows = rs("adv_mh_students")
adv_mh_cols = 3
Dim adv_mh_col(3)
adv_mh_col(1) = "Banner # "
adv_mh_col(2) = "First Name"
adv_mh_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					adv_mh="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='Adv' and Concentration='MH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open adv_mh,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
pdf.ln(1)
    Totaladvmh_Student = "Total number of students in MH : "&advmh_rows
    pdf.ChapterBody(Totaladvmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
advschquery="SELECT Count(distinct UIN) adv_sch_students FROM Applicants where Program_type='Adv' and Concentration='SCH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open advschquery,conn
advsch_rows = rs("adv_sch_students")
adv_sch_cols = 3
Dim adv_sch_col(3)
adv_sch_col(1) = "Banner # "
adv_sch_col(2) = "First Name"
adv_sch_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					adv_chud="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='Adv' and Concentration='SCH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open adv_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
    Totaladvsch_Student = "Total number of students in SCH : "&advsch_rows
    pdf.ChapterBody(Totaladvsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
advquery="SELECT Count(distinct UIN) adv_students FROM Applicants where Program_type='Adv' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open advquery,conn
    
    Totaladv_Student = "Total number of students in ADV : "&rs("adv_students")
    pdf.ChapterBody(Totaladv_Student)
    pdf.Ln(15)
rs.close



'//////// FT students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
ft_chf_query="SELECT Count(distinct UIN) ft_chf_students FROM Applicants where Program_type='FT' and Concentration='CHF' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open ft_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option FT")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// FT CHF ////////////
ft_chf_rows = rs("ft_chf_students")
ft_chf_cols = 3
Dim ft_chf_col(3)
ft_chf_col(1) = "Banner # "
ft_chf_col(2) = "First Name"
ft_chf_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					ftchf_query="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='FT' and Concentration='CHF' and Admission_decision = 'W' and term_cd='"&AdmitTerm&"'"
					rs.Open ftchf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_FT_CHF_Student = "Total number of students in CHF : "&ft_chf_rows
    pdf.ChapterBody(Total_FT_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
ftchudquery="SELECT Count(distinct UIN) ft_chud_students FROM Applicants where Program_type='FT' and Concentration='CHUD' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open ftchudquery,conn
'//////// FT CHUD ////////////
ftchud_rows = rs("ft_chud_students")
ft_chud_cols = 3
Dim ft_chud_col(3)
ft_chud_col(1) = "Banner # "
ft_chud_col(2) = "First Name"
ft_chud_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					ft_chud="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='FT' and Concentration='CHUD' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open ft_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    
    pdf.Ln(1)
    rs.close
    Totalftchud_Student = "Total number of students in CHUD : "&ftchud_rows
    pdf.ChapterBody(Totalftchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
ftmhquery="SELECT Count(distinct UIN) ft_mh_students FROM Applicants where Program_type='FT' and Concentration='MH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open ftmhquery,conn
ftmh_rows = rs("ft_mh_students")
ft_mh_cols = 3
Dim ft_mh_col(3)
ft_mh_col(1) = "Banner # "
ft_mh_col(2) = "First Name"
ft_mh_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					ft_mh="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='FT' and Concentration='MH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open ft_mh,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
pdf.ln(1)
    Totalftmh_Student = "Total number of students in MH : "&ftmh_rows
    pdf.ChapterBody(Totalftmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
ftschquery="SELECT Count(distinct UIN) ft_sch_students FROM Applicants where Program_type='FT' and Concentration='SCH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open ftschquery,conn
ftsch_rows = rs("ft_sch_students")
ft_sch_cols = 3
Dim ft_sch_col(3)
ft_sch_col(1) = "Banner # "
ft_sch_col(2) = "First Name"
ft_sch_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					ft_chud="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='FT' and Concentration='SCH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open ft_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
    Totalftsch_Student = "Total number of students in SCH : "&ftsch_rows
    pdf.ChapterBody(Totalftsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
ftquery="SELECT Count(distinct UIN) ft_students FROM Applicants where Program_type='FT' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open ftquery,conn
    
    Totalft_Student = "Total number of students in FT : "&rs("ft_students")
    pdf.ChapterBody(Totalft_Student)
    pdf.Ln(35)
    rs.close

    

'//////// MPH-ADV students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
mphadv_chf_query="SELECT Count(distinct UIN) mph_adv_chf_students FROM Applicants where Program_type='MPH-ADV' and Concentration='CHF' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadv_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option MPH-ADV")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// MPH-ADV CHF ////////////
mph_adv_chf_rows = rs("mph_adv_chf_students")
mph_adv_chf_cols = 3
Dim mph_adv_chf_col(3)
mph_adv_chf_col(1) = "Banner # "
mph_adv_chf_col(2) = "First Name"
mph_adv_chf_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_advchf_query="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='MPH-ADV' and Concentration='CHF' and Admission_decision = 'W' and term_cd='"&AdmitTerm&"'"
					rs.Open mph_advchf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_mph_adv_CHF_Student = "Total number of students in CHF : "&mph_adv_chf_rows
    pdf.ChapterBody(Total_mph_adv_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphadvchudquery="SELECT Count(distinct UIN) mphadv_chud_students FROM Applicants where Program_type='MPH-ADV' and Concentration='CHUD' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadvchudquery,conn
'//////// FT CHUD ////////////
mphadvchud_rows = rs("mphadv_chud_students")
mph_adv_chud_cols = 3
Dim mph_adv_chud_col(3)
mph_adv_chud_col(1) = "Banner # "
mph_adv_chud_col(2) = "First Name"
mph_adv_chud_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_adv_chud="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='MPH-ADV' and Concentration='CHUD' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mph_adv_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    
    pdf.Ln(1)
    rs.close
    Totalmphadvchud_Student = "Total number of students in CHUD : "&mphadvchud_rows
    pdf.ChapterBody(Totalmphadvchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphadvmhquery="SELECT Count(distinct UIN) mphadv_mh_students FROM Applicants where Program_type='MPH-ADV' and Concentration='MH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadvmhquery,conn
mphadvmh_rows = rs("mphadv_mh_students")
mph_adv_mh_cols = 3
Dim mph_adv_mh_col(3)
mph_adv_mh_col(1) = "Banner # "
mph_adv_mh_col(2) = "First Name"
mph_adv_mh_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_adv_mh="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='MPH-ADV' and Concentration='MH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mph_adv_mh,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
pdf.ln(1)
    Totalmphadvmh_Student = "Total number of students in MH : "&mphadvmh_rows
    pdf.ChapterBody(Totalmphadvmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphadvschquery="SELECT Count(distinct UIN) mph_adv_sch_students FROM Applicants where Program_type='MPH_ADV' and Concentration='SCH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadvschquery,conn
mphadvsch_rows = rs("mph_adv_sch_students")
mph_adv_sch_cols = 3
Dim mph_adv_sch_col(3)
mph_adv_sch_col(1) = "Banner # "
mph_adv_sch_col(2) = "First Name"
mph_adv_sch_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_adv_sch="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='MPH-ADV' and Concentration='SCH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mph_adv_sch,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
    Totalmphadvsch_Student = "Total number of students in SCH : "&mphadvsch_rows
    pdf.ChapterBody(Totalmphadvsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
mphadvquery="SELECT Count(distinct UIN) mph_adv_students FROM Applicants where Program_type='MPH-ADV' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadvquery,conn
    
    Totalmphadv_Student = "Total number of students in MPH-ADV : "&rs("mph_adv_students")
    pdf.ChapterBody(Totalmphadv_Student)
    pdf.Ln(15)
    rs.close

    '//////// MPH-FT students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
mphft_chf_query="SELECT Count(distinct UIN) mph_ft_chf_students FROM Applicants where Program_type='MPH-FT' and Concentration='CHF' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphft_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option MPH-FT")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// MPH-FT CHF ////////////
mph_ft_chf_rows = rs("mph_ft_chf_students")
mph_ft_chf_cols = 3
Dim mph_ft_chf_col(3)
mph_ft_chf_col(1) = "Banner # "
mph_ft_chf_col(2) = "First Name"
mph_ft_chf_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_ftchf_query="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='MPH-FT' and Concentration='CHF' and Admission_decision = 'W' and term_cd='"&AdmitTerm&"'"
					rs.Open mph_ftchf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_mph_ft_CHF_Student = "Total number of students in CHF : "&mph_ft_chf_rows
    pdf.ChapterBody(Total_mph_ft_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphftchudquery="SELECT Count(distinct UIN) mphft_chud_students FROM Applicants where Program_type='MPH-FT' and Concentration='CHUD' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphftchudquery,conn
'//////// MPH FT CHUD ////////////
mphftchud_rows = rs("mphft_chud_students")
mph_ft_chud_cols = 3
Dim mph_ft_chud_col(3)
mph_ft_chud_col(1) = "Banner # "
mph_ft_chud_col(2) = "First Name"
mph_ft_chud_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_ft_chud="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='MPH-FT' and Concentration='CHUD' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mph_ft_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    
    pdf.Ln(1)
    rs.close
    Totalmphftchud_Student = "Total number of students in CHUD : "&mphftchud_rows
    pdf.ChapterBody(Totalmphftchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphftmhquery="SELECT Count(distinct UIN) mphft_mh_students FROM Applicants where Program_type='MPH-FT' and Concentration='MH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphftmhquery,conn
mphftmh_rows = rs("mphft_mh_students")
mph_ft_mh_cols = 3
Dim mph_ft_mh_col(3)
mph_ft_mh_col(1) = "Banner # "
mph_ft_mh_col(2) = "First Name"
mph_ft_mh_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_ft_mh="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='MPH-FT' and Concentration='MH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mph_ft_mh,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
pdf.ln(1)
    Totalmphftmh_Student = "Total number of students in MH : "&mphftmh_rows
    pdf.ChapterBody(Totalmphftmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphftschquery="SELECT Count(distinct UIN) mph_ft_sch_students FROM Applicants where Program_type='MPH_FT' and Concentration='SCH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphftschquery,conn
mphftsch_rows = rs("mph_ft_sch_students")
mph_ft_sch_cols = 3
Dim mph_ft_sch_col(3)
mph_ft_sch_col(1) = "Banner # "
mph_ft_sch_col(2) = "First Name"
mph_ft_sch_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_ft_sch="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='MPH-FT' and Concentration='SCH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mph_ft_sch,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
    Totalmphftsch_Student = "Total number of students in SCH : "&mphftsch_rows
    pdf.ChapterBody(Totalmphftsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
mphftquery="SELECT Count(distinct UIN) mph_ft_students FROM Applicants where Program_type='MPH-FT' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphftquery,conn
    
    Totalmphft_Student = "Total number of students in MPH-FT : "&rs("mph_ft_students")
    pdf.ChapterBody(Totalmphft_Student)
    pdf.Ln(60)
    rs.close


    '//////// MPH-PM students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
mphpm_chf_query="SELECT Count(distinct UIN) mph_pm_chf_students FROM Applicants where Program_type='MPH-PM' and Concentration='CHF' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphpm_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option MPH-PM")
pdf.GreyTitle("A. Concentration CHF")


'//////// MPH-PM CHF ////////////
mph_pm_chf_rows = rs("mph_pm_chf_students")
mph_pm_chf_cols = 4
Dim mph_pm_chf_col(4)
mph_pm_chf_col(1) = "Banner # "
mph_pm_chf_col(2) = "First Name"
mph_pm_chf_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_pmchf_query="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='MPH-PM' and Concentration='CHF' and Admission_decision = 'W' and term_cd='"&AdmitTerm&"'"
					rs.Open mph_pmchf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_mph_pm_CHF_Student = "Total number of students in CHF : "&mph_pm_chf_rows
    pdf.ChapterBody(Total_mph_pm_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphpmchudquery="SELECT Count(distinct UIN) mphpm_chud_students FROM Applicants where Program_type='MPH-PM' and Concentration='CHUD' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphpmchudquery,conn
'//////// MPH PM CHUD ////////////
mphpmchud_rows = rs("mphpm_chud_students")
mph_pm_chud_cols = 3
Dim mph_pm_chud_col(3)
mph_pm_chud_col(1) = "Banner # "
mph_pm_chud_col(2) = "First Name"
mph_pm_chud_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_pm_chud="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='MPH-PM' and Concentration='CHUD' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mph_pm_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    
    pdf.Ln(1)
    rs.close
    Totalmphpmchud_Student = "Total number of students in CHUD : "&mphpmchud_rows
    pdf.ChapterBody(Totalmphpmchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphpmmhquery="SELECT Count(distinct UIN) mphpm_mh_students FROM Applicants where Program_type='MPH-PM' and Concentration='MH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphpmmhquery,conn
mphpmmh_rows = rs("mphpm_mh_students")
mph_pm_mh_cols = 3
Dim mph_pm_mh_col(3)
mph_pm_mh_col(1) = "Banner # "
mph_pm_mh_col(2) = "First Name"
mph_pm_mh_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_pm_mh="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='MPH-PM' and Concentration='MH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mph_pm_mh,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
pdf.ln(1)
    Totalmphpmmh_Student = "Total number of students in MH : "&mphpmmh_rows
    pdf.ChapterBody(Totalmphpmmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphpmschquery="SELECT Count(distinct UIN) mph_pm_sch_students FROM Applicants where Program_type='MPH_PM' and Concentration='SCH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphpmschquery,conn
mphpmsch_rows = rs("mph_pm_sch_students")
mph_pm_sch_cols = 3
Dim mph_pm_sch_col(3)
mph_pm_sch_col(1) = "Banner # "
mph_pm_sch_col(2) = "First Name"
mph_pm_sch_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_pm_sch="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='MPH-PM' and Concentration='SCH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mph_pm_sch,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
    Totalmphpmsch_Student = "Total number of students in SCH : "&mphpmsch_rows
    pdf.ChapterBody(Totalmphpmsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
mphpmquery="SELECT Count(distinct UIN) mph_pm_students FROM Applicants where Program_type='MPH-PM' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphpmquery,conn
    
    Totalmphpm_Student = "Total number of students in MPH-PM : "&rs("mph_pm_students")
    pdf.ChapterBody(Totalmphpm_Student)
    pdf.Ln(15)
    rs.close



'//////// PM students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
pm_chf_query="SELECT Count(distinct UIN) pm_chf_students FROM Applicants where Program_type='PM' and Concentration='CHF' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open pm_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option PM")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// PM CHF ////////////
pm_chf_rows = rs("pm_chf_students")
pm_chf_cols = 3
Dim pm_chf_col(3)
pm_chf_col(1) = "Banner # "
pm_chf_col(2) = "First Name"
pm_chf_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					pmchf_query="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='PM' and Concentration='CHF' and Admission_decision = 'W' and term_cd='"&AdmitTerm&"'"
					rs.Open pmchf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_PM_CHF_Student = "Total number of students in CHF : "&pm_chf_rows
    pdf.ChapterBody(Total_PM_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
pmchudquery="SELECT Count(distinct UIN) pm_chud_students FROM Applicants where Program_type='PM' and Concentration='CHUD' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open pmchudquery,conn
'//////// PM CHUD ////////////
pmchud_rows = rs("pm_chud_students")
pm_chud_cols = 3
Dim pm_chud_col(3)
pm_chud_col(1) = "Banner # "
pm_chud_col(2) = "First Name"
pm_chud_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					pm_chud="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='PM' and Concentration='CHUD' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open pm_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    
    pdf.Ln(1)
    rs.close
    Totalpmchud_Student = "Total number of students in CHUD : "&pmchud_rows
    pdf.ChapterBody(Totalpmchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
pmmhquery="SELECT Count(distinct UIN) pm_mh_students FROM Applicants where Program_type='PM' and Concentration='MH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open pmmhquery,conn
pmmh_rows = rs("pm_mh_students")
pm_mh_cols = 3
Dim pm_mh_col(3)
pm_mh_col(1) = "Banner # "
pm_mh_col(2) = "First Name"
pm_mh_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					pm_mh="SELECT UIN, Firstname, LastName FROM Applicants where Program_type='PM' and Concentration='MH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open pm_mh,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
pdf.ln(1)
    Totalpmmh_Student = "Total number of students in MH : "&pmmh_rows
    pdf.ChapterBody(Totalpmmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
pmschquery="SELECT Count(distinct UIN) pm_sch_students FROM Applicants where Program_type='PM' and Concentration='SCH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open pmschquery,conn
pmsch_rows = rs("pm_sch_students")
pm_sch_cols = 3
Dim pm_sch_col(3)
pm_sch_col(1) = "Banner # "
pm_sch_col(2) = "First Name"
pm_sch_col(3) = "Last Name"



pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					pm_chud="SELECT UIN, Firstname, LastName, Confirmed FROM Applicants where Program_type='PM' and Concentration='SCH' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"'"

					rs.Open pm_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    
                    
                    pdf.Row a,b,c
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
    Totalpmsch_Student = "Total number of students in SCH : "&pmsch_rows
    pdf.ChapterBody(Totalpmsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
pmquery="SELECT Count(distinct UIN) pm_students FROM Applicants where Program_type='PM' and Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open pmquery,conn
    
    Totalpm_Student = "Total number of students in PM : "&rs("pm_students")
    pdf.ChapterBody(Totalpm_Student)
    pdf.Ln(5)
    rs.close

    set rs=Server.CreateObject("ADODB.recordset")
totalquery="SELECT Count(distinct UIN) total_students FROM Applicants where Admission_decision = 'W' and Term_CD like '"&AdmitTerm&"' "
rs.Open totalquery,conn
    
    Total_Student = "Total number of students in Waitlist : "&rs("total_students")
    pdf.ChapterBody(Total_Student)
    pdf.Ln(5)
    rs.close




pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
