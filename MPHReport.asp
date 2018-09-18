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

pdf.ChapterTitle2("                          Report 10-    MPH Students - "&Termsel& "      "  &LastUpdatedDt&" "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(3)

rs.close



'//////// MPH-FT students  ////////////

pdf.OrangeTitle("Program Option MPH-FT")
pdf.GreyTitle("A. Concentration SCH")

set rs=Server.CreateObject("ADODB.recordset")
mphftsch_query="SELECT Count(distinct UIN) mphftsch_students FROM Applicants where Program_type='MPH-FT' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphftsch_query,conn
'pdf.ChapterBody(Total_Student)


'pdf.FancyTable()

'//////// MPH-FT SCH////////////
mphftsch_rows = rs("mphftsch_students")
mphftsch_cols = 5
Dim mphftsch_col(5)
mphftsch_col(1) = "Banner # "
mphftsch_col(2) = "First Name"
mphftsch_col(3) = "Last Name"
mphftsch_col(4) = "Admission Decision"
mphftsch_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admission Decision", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_ftsch_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-FT' and Concentration='SCH' and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open mph_ftsch_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPH_FTsch_Student = "Total number of students in MPH-FT SCH Concentration: "&mphftsch_rows
    pdf.ChapterBody(Total_MPH_FTsch_Student)
    pdf.Ln(5)
pdf.GreyTitle("B. Concentration Non SCH")
set rs=Server.CreateObject("ADODB.recordset")
mphftnonsch_query="SELECT Count(distinct UIN) mphftnonsch_students FROM Applicants where Program_type='MPH-FT' and Concentration<>'SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphftnonsch_query,conn
'pdf.ChapterBody(Total_Student)


'pdf.FancyTable()
'//////// MPH-FT non SCH////////////
mphftnonsch_rows = rs("mphftnonsch_students")
mphftnonsch_cols = 5
Dim mphftnonsch_col(5)
mphftnonsch_col(1) = "Banner # "
mphftnonsch_col(2) = "First Name"
mphftnonsch_col(3) = "Last Name"
mphftnonsch_col(4) = "Admission Decision"
mphftnonsch_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admission Decision", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_ftnonsch_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-FT' and Concentration<>'SCH' and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open mph_ftnonsch_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPH_FTnonsch_Student = "Total number of students in MPH-FT non SCH Concentration: "&mphftnonsch_rows
    pdf.ChapterBody(Total_MPH_FTnonsch_Student)
    pdf.Ln(5)

 set rs=Server.CreateObject("ADODB.recordset")
mphft_query="SELECT Count(distinct UIN) mphft_students FROM Applicants where Program_type='MPH-FT' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphft_query,conn

    Total_MPH_FT_Student = "Total number of students in MPH-FT : "&rs("mphft_students")
    pdf.ChapterBody(Total_MPH_FT_Student)
    pdf.Ln(5)
    rs.close

'//////// MPH-PM students ////////////
pdf.OrangeTitle("Program Option MPH-PM")
pdf.GreyTitle("A. Concentration SCH")
set rs=Server.CreateObject("ADODB.recordset")
mphpmsch_query="SELECT Count(distinct UIN) mphpmsch_students FROM Applicants where Program_type='MPH-PM'  and Concentration='SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphpmsch_query,conn
'pdf.ChapterBody(Total_Student)
'pdf.FancyTable()

'//////// MPH-PM SCH ////////////
mphpmsch_rows = rs("mphpmsch_students")
mphpmsch_cols = 5
Dim mphpmsch_col(5)
mphpmsch_col(1) = "Banner # "
mphpmsch_col(2) = "First Name"
mphpmsch_col(3) = "Last Name"
mphpmsch_col(4) = "Admission Decision"
mphpmsch_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admission Decision", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_pmsch_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-PM' and Concentration='SCH' and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open mph_pmsch_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPH_PMsch_Student = "Total number of students in MPH-PM SCH Concentration: "&mphpmsch_rows
    pdf.ChapterBody(Total_MPH_PMsch_Student)
    pdf.Ln(5)


'pdf.FancyTable()
pdf.GreyTitle("B. Concentration Non SCH")
'//////// MPH-PM non SCH////////////
    set rs=Server.CreateObject("ADODB.recordset")
mphpmnonsch_query="SELECT Count(distinct UIN) mphpmnonsch_students FROM Applicants where Program_type='MPH-PM' and Concentration <> 'SCH'  and Term_CD like '"&AdmitTerm&"' "
rs.Open mphpmnonsch_query,conn
mphpmnonsch_rows = rs("mphpmnonsch_students")
mphpmnonsch_cols = 5
Dim mphpmnonsch_col(5)
mphpmnonsch_col(1) = "Banner # "
mphpmnonsch_col(2) = "First Name"
mphpmnonsch_col(3) = "Last Name"
mphpmnonsch_col(4) = "Admission Decision"
mphpmnonsch_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admission Decision", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_pmnonsch_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-PM' and Concentration<>'SCH' and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open mph_pmnonsch_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPH_PMnonsch_Student = "Total number of students in MPH-PM and Non SCH Concentration: "&mphpmnonsch_rows
    pdf.ChapterBody(Total_MPH_PMnonsch_Student)
    pdf.Ln(5)

    set rs=Server.CreateObject("ADODB.recordset")
mphpm_query="SELECT Count(distinct UIN) mphpm_students FROM Applicants where Program_type='MPH-PM'  and Term_CD like '"&AdmitTerm&"' "
rs.Open mphpm_query,conn

    Total_MPH_PM_Student = "Total number of students in MPH-PM : "&rs("mphpm_students")
    pdf.ChapterBody(Total_MPH_PM_Student)
    pdf.Ln(5)
    rs.close
    
'////// MPH-Adv Students ////////
pdf.OrangeTitle("Program Option MPH-ADV")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// MPH Adv CHF ////////////

set rs=Server.CreateObject("ADODB.recordset")
mphadvchf_query="SELECT Count(distinct UIN) mphadvchf_students FROM Applicants where Program_type='MPH-ADV'  and Concentration='CHF' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadvchf_query,conn
'pdf.ChapterBody(Total_Student)

'pdf.FancyTable()

'//////// MPH-ADV CHF////////////
mphadvchf_rows = rs("mphadvchf_students")
mphadvchf_cols = 5
Dim mphadvchf_col(5)
mphadvchf_col(1) = "Banner # "
mphadvchf_col(2) = "First Name"
mphadvchf_col(3) = "Last Name"
mphadvchf_col(4) = "Admission Decision"
mphadvchf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admission Decision", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_advchf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-ADV'  and Concentration='CHF' and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open mph_advchf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPH_ADVchf_Student = "Total number of students in MPH-ADV CHF Concentration : "&mphadvchf_rows
    pdf.ChapterBody(Total_MPH_ADVchf_Student)
    pdf.Ln(5)

pdf.GreyTitle("B. Concentration CHUD")
'//////// MPH Adv CHUD ////////////

set rs=Server.CreateObject("ADODB.recordset")
mphadvchud_query="SELECT Count(distinct UIN) mphadvchud_students FROM Applicants where Program_type='MPH-ADV'  and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadvchud_query,conn
'pdf.ChapterBody(Total_Student)

'pdf.FancyTable()

'//////// MPH-ADVchud ////////////
mphadvchud_rows = rs("mphadvchud_students")
mphadvchud_cols = 5
Dim mphadvchud_col(5)
mphadvchud_col(1) = "Banner # "
mphadvchud_col(2) = "First Name"
mphadvchud_col(3) = "Last Name"
mphadvchud_col(4) = "Admission Decision"
mphadvchud_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admission Decision", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_advchud_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-ADV'  and Concentration='CHUD' and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open mph_advchud_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPH_ADVchud_Student = "Total number of students in MPH-ADV CHUD Concentration: "&mphadvchud_rows
    pdf.ChapterBody(Total_MPH_ADVchud_Student)
    pdf.Ln(5)
pdf.GreyTitle("C. Concentration MH")
'//////// MPH Adv MH ////////////

set rs=Server.CreateObject("ADODB.recordset")
mphadvmh_query="SELECT Count(distinct UIN) mphadvmh_students FROM Applicants where Program_type='MPH-ADV'  and Concentration='MH' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadvmh_query,conn
'pdf.ChapterBody(Total_Student)

'pdf.FancyTable()

'//////// MPH-ADVmh ////////////
mphadvmh_rows = rs("mphadvmh_students")
mphadvmh_cols = 5
Dim mphadvmh_col(5)
mphadvmh_col(1) = "Banner # "
mphadvmh_col(2) = "First Name"
mphadvmh_col(3) = "Last Name"
mphadvmh_col(4) = "Admission Decision"
mphadvmh_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admission Decision", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_advmh_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-ADV'  and Concentration='MH' and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open mph_advmh_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPH_ADVmh_Student = "Total number of students in MPH-ADV MH Concentration: "&mphadvmh_rows
    pdf.ChapterBody(Total_MPH_ADVmh_Student)
    pdf.Ln(5)

pdf.GreyTitle("D. Concentration SCH")
'//////// MPH Adv SCH ////////////

set rs=Server.CreateObject("ADODB.recordset")
mphadvsch_query="SELECT Count(distinct UIN) mphadvsch_students FROM Applicants where Program_type='MPH-ADV'  and Concentration='SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadvsch_query,conn
'pdf.ChapterBody(Total_Student)

'pdf.FancyTable()

'//////// MPH-ADV sch////////////
mphadvsch_rows = rs("mphadvsch_students")
mphadvsch_cols = 5
Dim mphadvsch_col(5)
mphadvsch_col(1) = "Banner # "
mphadvsch_col(2) = "First Name"
mphadvsch_col(3) = "Last Name"
mphadvsch_col(4) = "Admission Decision"
mphadvsch_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admission Decision", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mph_advsch_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-ADV'  and Concentration='SCH' and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open mph_advsch_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPH_ADVsch_Student = "Total number of students in MPH-ADV SCH Concentration: "&mphadvsch_rows
    pdf.ChapterBody(Total_MPH_ADVsch_Student)
    pdf.Ln(5)
  set rs=Server.CreateObject("ADODB.recordset")
mphadv_query="SELECT Count(distinct UIN) mphadv_students FROM Applicants where Program_type='MPH-ADV'  and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadv_query,conn

    Total_MPH_ADV_Student = "Total number of students in MPH-ADV : "&rs("mphadv_students")
    pdf.ChapterBody(Total_MPH_ADV_Student)
    pdf.Ln(5)


    '////// MPH-TR Students ////////
    pdf.OrangeTitle("Program Option MPH-TR")
    '//////// MPH-TR CHF////////////
    pdf.GreyTitle("A. Concentration CHF")
set rs=Server.CreateObject("ADODB.recordset")
mphtrchf_query="SELECT Count(distinct UIN) mphtrchf_students FROM Applicants where Program_type='MPH-TR'  and Concentration='CHF' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphtrchf_query,conn
'pdf.ChapterBody(Total_Student)
mphtrchf_rows = rs("mphtrchf_students")
mphtrchf_cols = 5
Dim mphtrchf_col(5)
mphtrchf_col(1) = "Banner # "
mphtrchf_col(2) = "First Name"
mphtrchf_col(3) = "Last Name"
mphtrchf_col(4) = "Admission Decision"
mphtrchf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admission Decision", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mphtrchf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-TR'  and Concentration='CHF' and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open mphtrchf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPHtrchf_Student = "Total number of students in MPH-TR CHF Concentration : "&mphtrchf_rows
    pdf.ChapterBody(Total_MPHtrchf_Student)
    pdf.Ln(5)

        '//////// MPH-TR CHUD////////////
    pdf.GreyTitle("B. Concentration CHUD")
set rs=Server.CreateObject("ADODB.recordset")
mphtrchud_query="SELECT Count(distinct UIN) mphtrchud_students FROM Applicants where Program_type='MPH-TR' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphtrchud_query,conn
'pdf.ChapterBody(Total_Student)
mphtrchud_rows = rs("mphtrchud_students")
mphtrchud_cols = 5
Dim mphtrchud_col(5)
mphtrchud_col(1) = "Banner # "
mphtrchud_col(2) = "First Name"
mphtrchud_col(3) = "Last Name"
mphtrchud_col(4) = "Admission Decision"
mphtrchud_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admission Decision", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mphtrchud_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-TR'  and Concentration='CHUD' and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open mphtrchud_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPHtrchud_Student = "Total number of students in MPH-TR CHUD Concentration : "&mphtrchud_rows
    pdf.ChapterBody(Total_MPHtrchud_Student)
    pdf.Ln(5)

        '//////// MPH-TR MH////////////
    pdf.GreyTitle("C. Concentration MH")
set rs=Server.CreateObject("ADODB.recordset")
mphtrmh_query="SELECT Count(distinct UIN) mphtrmh_students FROM Applicants where Program_type='MPH-TR'  and Concentration='MH' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphtrmh_query,conn
'pdf.ChapterBody(Total_Student)
mphtrmh_rows = rs("mphtrmh_students")
mphtrmh_cols = 5
Dim mphtrmh_col(5)
mphtrmh_col(1) = "Banner # "
mphtrmh_col(2) = "First Name"
mphtrmh_col(3) = "Last Name"
mphtrmh_col(4) = "Admission Decision"
mphtrmh_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admission Decision", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mphtrmh_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-TR'  and Concentration='MH' and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open mphtrmh_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPHtrmh_Student = "Total number of students in MPH-TR MH Concentration : "&mphtrmh_rows
    pdf.ChapterBody(Total_MPHtrmh_Student)
    pdf.Ln(5)

        '//////// MPH-TR SCH////////////
    pdf.GreyTitle("D. Concentration SCH")
set rs=Server.CreateObject("ADODB.recordset")
mphtrsch_query="SELECT Count(distinct UIN) mphtrsch_students FROM Applicants where Program_type='MPH-TR'  and Concentration='SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphtrsch_query,conn
'pdf.ChapterBody(Total_Student)
mphtrsch_rows = rs("mphtrsch_students")
mphtrsch_cols = 5
Dim mphtrsch_col(5)
mphtrsch_col(1) = "Banner # "
mphtrsch_col(2) = "First Name"
mphtrsch_col(3) = "Last Name"
mphtrsch_col(4) = "Admission Decision"
mphtrsch_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Admission Decision", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					mphtrsch_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-TR'  and Concentration='SCH' and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open mphtrsch_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    

    pdf.Ln(1)
    Total_MPHtrsch_Student = "Total number of students in MPH-TR SCH Concentration : "&mphtrsch_rows
    pdf.ChapterBody(Total_MPHtrsch_Student)
    pdf.Ln(5)
 set rs=Server.CreateObject("ADODB.recordset")
mphtr_query="SELECT Count(distinct UIN) mphtr_students FROM Applicants where Program_type='MPH-TR'  and Term_CD like '"&AdmitTerm&"' "
rs.Open mphtr_query,conn

    Total_MPH_TR_Student = "Total number of students in MPH-TR : "&rs("mphtr_students")
    pdf.ChapterBody(Total_MPH_TR_Student)
    pdf.Ln(5)
    rs.close
   
pdf.GreyTitle("")
set rs=Server.CreateObject("ADODB.recordset")
mph_query="SELECT Count(distinct UIN) mph_students FROM Applicants where Program_type in ('MPH-TR', 'MPH-ADV', 'MPH-FT', 'MPH-PM')  and Term_CD like '"&AdmitTerm&"' "
rs.Open mph_query,conn
accepted_students = "Total Students : "&rs("mph_students")
pdf.ChapterBody(accepted_students)
pdf.Ln(5)
rs.close


pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
