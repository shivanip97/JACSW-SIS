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

pdf.ChapterTitle2("                   Report 2- Admissions Report - "&Termsel& " - Applied      "  &LastUpdatedDt&" "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(3)

rs.close


'////// Adv Students ////////
pdf.OrangeTitle("Program Option ADV")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// Adv CHF ////////////



set rs=Server.CreateObject("ADODB.recordset")
adv_chf_query="SELECT Count(distinct UIN) adv_chf_students FROM Applicants where Program_type='Adv' and Concentration='CHF' and Term_CD like '"&AdmitTerm&"' "
rs.Open adv_chf_query,conn

adv_chf_rows = rs("adv_chf_students")
adv_chf_cols = 5
Dim adv_chf_col(5)
adv_chf_col(1) = "Banner # "
adv_chf_col(2) = "First Name"
adv_chf_col(3) = "Last Name"
adv_chf_col(4) = "Admission Decision"
adv_chf_col(5) = "Confirmed"


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
					advchf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='Adv' and Concentration='CHF'  and term_cd='"&AdmitTerm&"'"
					rs.Open advchf_query,conn 
                  
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
    Total_Adv_CHF_Student = "Total number of students in CHF : "&adv_chf_rows
    pdf.ChapterBody(Total_Adv_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
advchudquery="SELECT Count(distinct UIN) adv_chud_students FROM Applicants where Program_type='Adv' and Concentration='CHUD'  and Term_CD like '"&AdmitTerm&"' "
rs.Open advchudquery,conn
'//////// Courses Table ////////////
advchud_rows = rs("adv_chud_students")
adv_chud_cols = 5
Dim adv_chud_col(5)
adv_chud_col(1) = "Banner # "
adv_chud_col(2) = "First Name"
adv_chud_col(3) = "Last Name"
adv_chud_col(4) = "Admission Decision"
adv_chud_col(5) = "Confirmed"


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
					adv_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='Adv' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"'"

					rs.Open adv_chud,conn 
                  
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
                    
    pdf.Ln(1)
    rs.close
    Totaladvchud_Student = "Total number of students in CHUD : "&advchud_rows
    pdf.ChapterBody(Totaladvchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
advmhquery="SELECT Count(distinct UIN) adv_mh_students FROM Applicants where Program_type='Adv' and Concentration='MH' and Term_CD like '"&AdmitTerm&"' "
rs.Open advmhquery,conn
advmh_rows = rs("adv_mh_students")
adv_mh_cols = 5
Dim adv_mh_col(5)
adv_mh_col(1) = "Banner # "
adv_mh_col(2) = "First Name"
adv_mh_col(3) = "Last Name"
adv_mh_col(4) = "Admission Decision"
adv_mh_col(5) = "Confirmed"

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
					adv_mh="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='Adv' and Concentration='MH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open adv_mh,conn 
                  
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
pdf.ln(1)
    Totaladvmh_Student = "Total number of students in MH : "&advmh_rows
    pdf.ChapterBody(Totaladvmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
advschquery="SELECT Count(distinct UIN) adv_sch_students FROM Applicants where Program_type='Adv' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open advschquery,conn
advsch_rows = rs("adv_sch_students")
adv_sch_cols = 5
Dim adv_sch_col(5)
adv_sch_col(1) = "Banner # "
adv_sch_col(2) = "First Name"
adv_sch_col(3) = "Last Name"
adv_sch_col(4) = "Admission Decision"
adv_sch_col(5) = "Confirmed"


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
					adv_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='Adv' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open adv_chud,conn 
                  
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
    Totaladvsch_Student = "Total number of students in SCH : "&advsch_rows
    pdf.ChapterBody(Totaladvsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
advquery="SELECT Count(distinct UIN) adv_students FROM Applicants where Program_type='Adv' and Term_CD like '"&AdmitTerm&"' "
rs.Open advquery,conn
    
    Totaladv_Student = "Total number of students in ADV : "&rs("adv_students")
    pdf.ChapterBody(Totaladv_Student)
    pdf.Ln(5)
    advstu=rs("adv_students")
rs.close



'//////// FT students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
ft_chf_query="SELECT Count(distinct UIN) ft_chf_students FROM Applicants where Program_type='FT' and Concentration='CHF' and Term_CD like '"&AdmitTerm&"' "
rs.Open ft_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option FT")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// FT CHF ////////////
ft_chf_rows = rs("ft_chf_students")
ft_chf_cols = 5
Dim ft_chf_col(5)
ft_chf_col(1) = "Banner # "
ft_chf_col(2) = "First Name"
ft_chf_col(3) = "Last Name"
ft_chf_col(4) = "Admission Decision"
ft_chf_col(5) = "Confirmed"


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
					ftchf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='FT' and Concentration='CHF' and term_cd='"&AdmitTerm&"'"
					rs.Open ftchf_query,conn 
                  
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
    Total_FT_CHF_Student = "Total number of students in CHF : "&ft_chf_rows
    pdf.ChapterBody(Total_FT_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
ftchudquery="SELECT Count(distinct UIN) ft_chud_students FROM Applicants where Program_type='FT' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"' "
rs.Open ftchudquery,conn
'//////// FT CHUD ////////////
ftchud_rows = rs("ft_chud_students")
ft_chud_cols = 5
Dim ft_chud_col(5)
ft_chud_col(1) = "Banner # "
ft_chud_col(2) = "First Name"
ft_chud_col(3) = "Last Name"
ft_chud_col(4) = "Admission Decision"
ft_chud_col(5) = "Confirmed"


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
					ft_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='FT' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"'"

					rs.Open ft_chud,conn 
                  
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
                    
    pdf.Ln(1)
    rs.close
    Totalftchud_Student = "Total number of students in CHUD : "&ftchud_rows
    pdf.ChapterBody(Totalftchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
ftmhquery="SELECT Count(distinct UIN) ft_mh_students FROM Applicants where Program_type='FT' and Concentration='MH'  and Term_CD like '"&AdmitTerm&"' "
rs.Open ftmhquery,conn
ftmh_rows = rs("ft_mh_students")
ft_mh_cols = 5
Dim ft_mh_col(5)
ft_mh_col(1) = "Banner # "
ft_mh_col(2) = "First Name"
ft_mh_col(3) = "Last Name"
ft_mh_col(4) = "Admission Decision"
ft_mh_col(5) = "Confirmed"


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
					ft_mh="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='FT' and Concentration='MH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open ft_mh,conn 
                  
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
pdf.ln(1)
    Totalftmh_Student = "Total number of students in MH : "&ftmh_rows
    pdf.ChapterBody(Totalftmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
ftschquery="SELECT Count(distinct UIN) ft_sch_students FROM Applicants where Program_type='FT' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open ftschquery,conn
ftsch_rows = rs("ft_sch_students")
ft_sch_cols = 5
Dim ft_sch_col(5)
ft_sch_col(1) = "Banner # "
ft_sch_col(2) = "First Name"
ft_sch_col(3) = "Last Name"
ft_sch_col(4) = "Admission Decision"
ft_sch_col(5) = "Confirmed"


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
					ft_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='FT' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open ft_chud,conn 
                  
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
    Totalftsch_Student = "Total number of students in SCH : "&ftsch_rows
    pdf.ChapterBody(Totalftsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
ftquery="SELECT Count(distinct UIN) ft_students FROM Applicants where Program_type='FT' and Term_CD like '"&AdmitTerm&"' "
rs.Open ftquery,conn
    
    Totalft_Student = "Total number of students in FT : "&rs("ft_students")
    pdf.ChapterBody(Totalft_Student)
    pdf.Ln(5)
    ftstu=rs("ft_students")
    rs.close

    

'//////// PM students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
pm_chf_query="SELECT Count(distinct UIN) pm_chf_students FROM Applicants where Program_type='PM' and Concentration='CHF' and Term_CD like '"&AdmitTerm&"' "
rs.Open pm_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option PM")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// PM CHF ////////////
pm_chf_rows = rs("pm_chf_students")
pm_chf_cols = 5
Dim pm_chf_col(5)
pm_chf_col(1) = "Banner # "
pm_chf_col(2) = "First Name"
pm_chf_col(3) = "Last Name"
pm_chf_col(4) = "Admission Decision"
pm_chf_col(5) = "Confirmed"

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
					pmchf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='PM' and Concentration='CHF' and term_cd='"&AdmitTerm&"'"
					rs.Open pmchf_query,conn 
                  
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
    Total_PM_CHF_Student = "Total number of students in CHF : "&pm_chf_rows
    pdf.ChapterBody(Total_PM_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
pmchudquery="SELECT Count(distinct UIN) pm_chud_students FROM Applicants where Program_type='PM' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"' "
rs.Open pmchudquery,conn
'//////// PM CHUD ////////////
pmchud_rows = rs("pm_chud_students")
pm_chud_cols = 5
Dim pm_chud_col(5)
pm_chud_col(1) = "Banner # "
pm_chud_col(2) = "First Name"
pm_chud_col(3) = "Last Name"
pm_chud_col(4) = "Admission Decision"
pm_chud_col(5) = "Confirmed"


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
					pm_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='PM' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"'"

					rs.Open pm_chud,conn 
                  
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
                    
    pdf.Ln(1)
    rs.close
    Totalpmchud_Student = "Total number of students in CHUD : "&pmchud_rows
    pdf.ChapterBody(Totalpmchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
pmmhquery="SELECT Count(distinct UIN) pm_mh_students FROM Applicants where Program_type='PM' and Concentration='MH' and Term_CD like '"&AdmitTerm&"' "
rs.Open pmmhquery,conn
pmmh_rows = rs("pm_mh_students")
pm_mh_cols = 5
Dim pm_mh_col(5)
pm_mh_col(1) = "Banner # "
pm_mh_col(2) = "First Name"
pm_mh_col(3) = "Last Name"
pm_mh_col(4) = "Admission Decision"
pm_mh_col(5) = "Confirmed"


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
					pm_mh="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='PM' and Concentration='MH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open pm_mh,conn 
                  
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
pdf.ln(1)
    Totalpmmh_Student = "Total number of students in MH : "&pmmh_rows
    pdf.ChapterBody(Totalpmmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
pmschquery="SELECT Count(distinct UIN) pm_sch_students FROM Applicants where Program_type='PM' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open pmschquery,conn
pmsch_rows = rs("pm_sch_students")
pm_sch_cols = 5
Dim pm_sch_col(5)
pm_sch_col(1) = "Banner # "
pm_sch_col(2) = "First Name"
pm_sch_col(3) = "Last Name"
pm_sch_col(4) = "Admission Decision"
pm_sch_col(5) = "Confirmed"


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
					pm_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='PM' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open pm_chud,conn 
                  
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
    Totalpmsch_Student = "Total number of students in SCH : "&pmsch_rows
    pdf.ChapterBody(Totalpmsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
pmquery="SELECT Count(distinct UIN) pm_students FROM Applicants where Program_type='PM' and Term_CD like '"&AdmitTerm&"' "
rs.Open pmquery,conn
    
    Totalpm_Student = "Total number of students in PM : "&rs("pm_students")
    pdf.ChapterBody(Totalpm_Student)
    pdf.Ln(5)
    pmstu=rs("pm_students")
    rs.close
   
pdf.Ln(5)



'////// TR Students ////////
pdf.OrangeTitle("Program Option TR")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// Tr CHF ////////////



set rs=Server.CreateObject("ADODB.recordset")
tr_chf_query="SELECT Count(distinct UIN) tr_chf_students FROM Applicants where Program_type='TR' and Concentration='CHF' and Term_CD like '"&AdmitTerm&"' "
rs.Open tr_chf_query,conn

tr_chf_rows = rs("tr_chf_students")
tr_chf_cols = 5
Dim tr_chf_col(5)
tr_chf_col(1) = "Banner # "
tr_chf_col(2) = "First Name"
tr_chf_col(3) = "Last Name"
tr_chf_col(4) = "Admission Decision"
tr_chf_col(5) = "Confirmed"


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
					trchf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='TR' and Concentration='CHF'  and term_cd='"&AdmitTerm&"'"
					rs.Open trchf_query,conn 
                  
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
    Total_TR_CHF_Student = "Total number of students in CHF : "&tr_chf_rows
    pdf.ChapterBody(Total_TR_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
trchudquery="SELECT Count(distinct UIN) tr_chud_students FROM Applicants where Program_type='TR' and Concentration='CHUD'  and Term_CD like '"&AdmitTerm&"' "
rs.Open trchudquery,conn
'//////// Courses Table ////////////
trchud_rows = rs("tr_chud_students")
tr_chud_cols = 5
Dim tr_chud_col(5)
tr_chud_col(1) = "Banner # "
tr_chud_col(2) = "First Name"
tr_chud_col(3) = "Last Name"
tr_chud_col(4) = "Admission Decision"
tr_chud_col(5) = "Confirmed"


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
					tr_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='TR' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"'"

					rs.Open tr_chud,conn 
                  
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
                    
    pdf.Ln(1)
    rs.close
    Totaltrchud_Student = "Total number of students in CHUD : "&trchud_rows
    pdf.ChapterBody(Totaltrchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
trmhquery="SELECT Count(distinct UIN) tr_mh_students FROM Applicants where Program_type='TR' and Concentration='MH' and Term_CD like '"&AdmitTerm&"' "
rs.Open trmhquery,conn
trmh_rows = rs("tr_mh_students")
tr_mh_cols = 5
Dim tr_mh_col(5)
tr_mh_col(1) = "Banner # "
tr_mh_col(2) = "First Name"
tr_mh_col(3) = "Last Name"
tr_mh_col(4) = "Admission Decision"
tr_mh_col(5) = "Confirmed"

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
					tr_mh="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='TR' and Concentration='MH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open tr_mh,conn 
                  
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
pdf.ln(1)
    Totaltrmh_Student = "Total number of students in MH : "&trmh_rows
    pdf.ChapterBody(Totaltrmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
trschquery="SELECT Count(distinct UIN) tr_sch_students FROM Applicants where Program_type='TR' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open trschquery,conn
trsch_rows = rs("tr_sch_students")
tr_sch_cols = 5
Dim tr_sch_col(5)
tr_sch_col(1) = "Banner # "
tr_sch_col(2) = "First Name"
tr_sch_col(3) = "Last Name"
tr_sch_col(4) = "Admission Decision"
tr_sch_col(5) = "Confirmed"


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
					tr_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='TR' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open tr_chud,conn 
                  
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
    Totaltrsch_Student = "Total number of students in SCH : "&trsch_rows
    pdf.ChapterBody(Totaltrsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
trquery="SELECT Count(distinct UIN) tr_students FROM Applicants where Program_type='TR' and Term_CD like '"&AdmitTerm&"' "
rs.Open trquery,conn
    
    Totaltr_Student = "Total number of students in TR : "&rs("tr_students")
    pdf.ChapterBody(Totaltr_Student)
    pdf.Ln(5)
    trstu=rs("tr_students")
rs.close



'//////// TR-FT students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
trft_chf_query="SELECT Count(distinct UIN) trft_chf_students FROM Applicants where Program_type='TR-FT' and Concentration='CHF' and Term_CD like '"&AdmitTerm&"' "
rs.Open trft_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option TR-FT")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// TR-FT CHF ////////////
trft_chf_rows = rs("trft_chf_students")
trft_chf_cols = 5
Dim trft_chf_col(5)
trft_chf_col(1) = "Banner # "
trft_chf_col(2) = "First Name"
trft_chf_col(3) = "Last Name"
trft_chf_col(4) = "Admission Decision"
trft_chf_col(5) = "Confirmed"


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
					trftchf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='TR-FT' and Concentration='CHF' and term_cd='"&AdmitTerm&"'"
					rs.Open trftchf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
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
    Total_TR_FT_CHF_Student = "Total number of students in CHF : "&trft_chf_rows
    pdf.ChapterBody(Total_TR_FT_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
trftchudquery="SELECT Count(distinct UIN) trft_chud_students FROM Applicants where Program_type='TR-FT' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"' "
rs.Open trftchudquery,conn
'//////// TR-FT CHUD ////////////
trftchud_rows = rs("trft_chud_students")
trft_chud_cols = 5
Dim trft_chud_col(5)
trft_chud_col(1) = "Banner # "
trft_chud_col(2) = "First Name"
trft_chud_col(3) = "Last Name"
trft_chud_col(4) = "Admission Decision"
trft_chud_col(5) = "Confirmed"


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
					trft_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='TR-FT' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"'"

					rs.Open trft_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    
    pdf.Ln(1)
    rs.close
    Totaltrftchud_Student = "Total number of students in CHUD : "&trftchud_rows
    pdf.ChapterBody(Totaltrftchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
trftmhquery="SELECT Count(distinct UIN) trft_mh_students FROM Applicants where Program_type='TR-FT' and Concentration='MH'  and Term_CD like '"&AdmitTerm&"' "
rs.Open trftmhquery,conn
trftmh_rows = rs("trft_mh_students")
trft_mh_cols = 5
Dim trft_mh_col(5)
trft_mh_col(1) = "Banner # "
trft_mh_col(2) = "First Name"
trft_mh_col(3) = "Last Name"
trft_mh_col(4) = "Admission Decision"
trft_mh_col(5) = "Confirmed"


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
					trft_mh="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='TR-FT' and Concentration='MH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open trft_mh,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
pdf.ln(1)
    Totaltrftmh_Student = "Total number of students in MH : "&trftmh_rows
    pdf.ChapterBody(Totaltrftmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
trftschquery="SELECT Count(distinct UIN) trft_sch_students FROM Applicants where Program_type='TR-FT' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open trftschquery,conn
trftsch_rows = rs("trft_sch_students")
trft_sch_cols = 5
Dim trft_sch_col(5)
trft_sch_col(1) = "Banner # "
trft_sch_col(2) = "First Name"
trft_sch_col(3) = "Last Name"
trft_sch_col(4) = "Admission Decision"
trft_sch_col(5) = "Confirmed"


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
					trft_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='TR-FT' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open trft_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
    Totaltrftsch_Student = "Total number of students in SCH : "&trftsch_rows
    pdf.ChapterBody(Totaltrftsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
trftquery="SELECT Count(distinct UIN) trft_students FROM Applicants where Program_type='TR-FT' and Term_CD like '"&AdmitTerm&"' "
rs.Open trftquery,conn
    
    Totaltrft_Student = "Total number of students in TR-FT : "&rs("trft_students")
    pdf.ChapterBody(Totaltrft_Student)
    pdf.Ln(5)
    trftstu=rs("trft_students")
    rs.close

    

'//////// TR-PM students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
trpm_chf_query="SELECT Count(distinct UIN) trpm_chf_students FROM Applicants where Program_type='TR-PM' and Concentration='CHF' and Term_CD like '"&AdmitTerm&"' "
rs.Open trpm_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option TR-PM")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// TR-PM CHF ////////////
trpm_chf_rows = rs("trpm_chf_students")
trpm_chf_cols = 5
Dim trpm_chf_col(5)
trpm_chf_col(1) = "Banner # "
trpm_chf_col(2) = "First Name"
trpm_chf_col(3) = "Last Name"
trpm_chf_col(4) = "Admission Decision"
trpm_chf_col(5) = "Confirmed"

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
					trpmchf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='TR-PM' and Concentration='CHF' and term_cd='"&AdmitTerm&"'"
					rs.Open trpmchf_query,conn 
                  
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
    Total_TR_PM_CHF_Student = "Total number of students in CHF : "&trpm_chf_rows
    pdf.ChapterBody(Total_TR_PM_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
trpmchudquery="SELECT Count(distinct UIN) trpm_chud_students FROM Applicants where Program_type='TR-PM' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"' "
rs.Open trpmchudquery,conn
'//////// TR-PM CHUD ////////////
trpmchud_rows = rs("trpm_chud_students")
trpm_chud_cols = 5
Dim trpm_chud_col(5)
trpm_chud_col(1) = "Banner # "
trpm_chud_col(2) = "First Name"
trpm_chud_col(3) = "Last Name"
trpm_chud_col(4) = "Admission Decision"
trpm_chud_col(5) = "Confirmed"


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
					trpm_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='TR-PM' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"'"

					rs.Open trpm_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    
    pdf.Ln(1)
    rs.close
    Totaltrpmchud_Student = "Total number of students in CHUD : "&trpmchud_rows
    pdf.ChapterBody(Totaltrpmchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
trpmmhquery="SELECT Count(distinct UIN) trpm_mh_students FROM Applicants where Program_type='TR-PM' and Concentration='MH' and Term_CD like '"&AdmitTerm&"' "
rs.Open trpmmhquery,conn
trpmmh_rows = rs("trpm_mh_students")
trpm_mh_cols = 5
Dim trpm_mh_col(5)
trpm_mh_col(1) = "Banner # "
trpm_mh_col(2) = "First Name"
trpm_mh_col(3) = "Last Name"
trpm_mh_col(4) = "Admission Decision"
trpm_mh_col(5) = "Confirmed"


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
					trpm_mh="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='TR-PM' and Concentration='MH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open trpm_mh,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
pdf.ln(1)
    Totaltrpmmh_Student = "Total number of students in MH : "&trpmmh_rows
    pdf.ChapterBody(Totaltrpmmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
trpmschquery="SELECT Count(distinct UIN) trpm_sch_students FROM Applicants where Program_type='TR-PM' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open trpmschquery,conn
trpmsch_rows = rs("trpm_sch_students")
trpm_sch_cols = 5
Dim trpm_sch_col(5)
trpm_sch_col(1) = "Banner # "
trpm_sch_col(2) = "First Name"
trpm_sch_col(3) = "Last Name"
trpm_sch_col(4) = "Admission Decision"
trpm_sch_col(5) = "Confirmed"


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
					trpm_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='TR-PM' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open trpm_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Admission_decision"),"|",",")
                    e = Replace(rs("Confirmed"),"|",",")

                    pdf.Row a,b,c,d,e
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
    Totaltrpmsch_Student = "Total number of students in SCH : "&trpmsch_rows
    pdf.ChapterBody(Totaltrpmsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
trpmquery="SELECT Count(distinct UIN) trpm_students FROM Applicants where Program_type='TR-PM' and Term_CD like '"&AdmitTerm&"' "
rs.Open trpmquery,conn
    
    Totaltrpm_Student = "Total number of students in TR-PM : "&rs("trpm_students")
    pdf.ChapterBody(Totaltrpm_Student)
    pdf.Ln(5)
    trpmstu=rs("trpm_students")
    rs.close



'//////// MPH-FT students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
mphft_chf_query="SELECT Count(distinct UIN) mphft_chf_students FROM Applicants where Program_type='MPH-FT' and Concentration='CHF' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphft_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option MPH-FT")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// MPH-FT CHF ////////////
mphft_chf_rows = rs("mphft_chf_students")
mphft_chf_cols = 5
Dim mphft_chf_col(5)
mphft_chf_col(1) = "Banner # "
mphft_chf_col(2) = "First Name"
mphft_chf_col(3) = "Last Name"
mphft_chf_col(4) = "Admission Decision"
mphft_chf_col(5) = "Confirmed"


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
					mphftchf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-FT' and Concentration='CHF' and term_cd='"&AdmitTerm&"'"
					rs.Open mphftchf_query,conn 
                  
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
    Total_MPH_FT_CHF_Student = "Total number of students in CHF : "&mphft_chf_rows
    pdf.ChapterBody(Total_MPH_FT_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphftchudquery="SELECT Count(distinct UIN) mphft_chud_students FROM Applicants where Program_type='MPH-FT' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphftchudquery,conn
'//////// MPH-FT CHUD ////////////
mphftchud_rows = rs("mphft_chud_students")
mphft_chud_cols = 5
Dim mphft_chud_col(5)
mphft_chud_col(1) = "Banner # "
mphft_chud_col(2) = "First Name"
mphft_chud_col(3) = "Last Name"
mphft_chud_col(4) = "Admission Decision"
mphft_chud_col(5) = "Confirmed"


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
					mphft_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-FT' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mphft_chud,conn 
                  
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
                    
    pdf.Ln(1)
    rs.close
    Totalmphftchud_Student = "Total number of students in CHUD : "&mphftchud_rows
    pdf.ChapterBody(Totalmphftchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphftmhquery="SELECT Count(distinct UIN) mphft_mh_students FROM Applicants where Program_type='MPH-FT' and Concentration='MH'  and Term_CD like '"&AdmitTerm&"' "
rs.Open mphftmhquery,conn
mphftmh_rows = rs("mphft_mh_students")
mphft_mh_cols = 5
Dim mphft_mh_col(5)
mphft_mh_col(1) = "Banner # "
mphft_mh_col(2) = "First Name"
mphft_mh_col(3) = "Last Name"
mphft_mh_col(4) = "Admission Decision"
mphft_mh_col(5) = "Confirmed"


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
					mphft_mh="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-FT' and Concentration='MH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mphft_mh,conn 
                  
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
pdf.ln(1)
    Totalmphftmh_Student = "Total number of students in MH : "&mphftmh_rows
    pdf.ChapterBody(Totalmphftmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphftschquery="SELECT Count(distinct UIN) mphft_sch_students FROM Applicants where Program_type='MPH-FT' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphftschquery,conn
mphftsch_rows = rs("mphft_sch_students")
mphft_sch_cols = 5
Dim mphft_sch_col(5)
mphft_sch_col(1) = "Banner # "
mphft_sch_col(2) = "First Name"
mphft_sch_col(3) = "Last Name"
mphft_sch_col(4) = "Admission Decision"
mphft_sch_col(5) = "Confirmed"


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
					mphft_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-FT' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mphft_chud,conn 
                  
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
    Totalmphftsch_Student = "Total number of students in SCH : "&mphftsch_rows
    pdf.ChapterBody(Totalmphftsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
mphftquery="SELECT Count(distinct UIN) mphft_students FROM Applicants where Program_type='MPH-FT' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphftquery,conn
    
    Totalmphft_Student = "Total number of students in MPH-FT : "&rs("mphft_students")
    pdf.ChapterBody(Totalmphft_Student)
    pdf.Ln(5)
    mphftstu=rs("mphft_students")
    rs.close

    

'//////// MPH-PM students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
mphpm_chf_query="SELECT Count(distinct UIN) mphpm_chf_students FROM Applicants where Program_type='MPH-PM' and Concentration='CHF' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphpm_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option MPH-PM")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// MPH-PM CHF ////////////
mphpm_chf_rows = rs("mphpm_chf_students")
mphpm_chf_cols = 5
Dim mphpm_chf_col(5)
mphpm_chf_col(1) = "Banner # "
mphpm_chf_col(2) = "First Name"
mphpm_chf_col(3) = "Last Name"
mphpm_chf_col(4) = "Admission Decision"
mphpm_chf_col(5) = "Confirmed"

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
					mphpmchf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-PM' and Concentration='CHF' and term_cd='"&AdmitTerm&"'"
					rs.Open mphpmchf_query,conn 
                  
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
    Total_MPH_PM_CHF_Student = "Total number of students in CHF : "&mphpm_chf_rows
    pdf.ChapterBody(Total_MPH_PM_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphpmchudquery="SELECT Count(distinct UIN) mphpm_chud_students FROM Applicants where Program_type='MPH-PM' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphpmchudquery,conn
'//////// MPH-PM CHUD ////////////
mphpmchud_rows = rs("mphpm_chud_students")
mphpm_chud_cols = 5
Dim mphpm_chud_col(5)
mphpm_chud_col(1) = "Banner # "
mphpm_chud_col(2) = "First Name"
mphpm_chud_col(3) = "Last Name"
mphpm_chud_col(4) = "Admission Decision"
mphpm_chud_col(5) = "Confirmed"


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
					mphpm_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-PM' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mphpm_chud,conn 
                  
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
                    
    pdf.Ln(1)
    rs.close
    Totalmphpmchud_Student = "Total number of students in CHUD : "&mphpmchud_rows
    pdf.ChapterBody(Totalmphpmchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphpmmhquery="SELECT Count(distinct UIN) mphpm_mh_students FROM Applicants where Program_type='MPH-PM' and Concentration='MH' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphpmmhquery,conn
mphpmmh_rows = rs("mphpm_mh_students")
mphpm_mh_cols = 5
Dim mphpm_mh_col(5)
mphpm_mh_col(1) = "Banner # "
mphpm_mh_col(2) = "First Name"
mphpm_mh_col(3) = "Last Name"
mphpm_mh_col(4) = "Admission Decision"
mphpm_mh_col(5) = "Confirmed"


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
					mphpm_mh="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-PM' and Concentration='MH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mphpm_mh,conn 
                  
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
pdf.ln(1)
    Totalmphpmmh_Student = "Total number of students in MH : "&mphpmmh_rows
    pdf.ChapterBody(Totalmphpmmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphpmschquery="SELECT Count(distinct UIN) mphpm_sch_students FROM Applicants where Program_type='MPH-PM' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphpmschquery,conn
mphpmsch_rows = rs("mphpm_sch_students")
mphpm_sch_cols = 5
Dim mphpm_sch_col(5)
mphpm_sch_col(1) = "Banner # "
mphpm_sch_col(2) = "First Name"
mphpm_sch_col(3) = "Last Name"
mphpm_sch_col(4) = "Admission Decision"
mphpm_sch_col(5) = "Confirmed"


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
					mphpm_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-PM' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mphpm_chud,conn 
                  
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
    Totalmphpmsch_Student = "Total number of students in SCH : "&mphpmsch_rows
    pdf.ChapterBody(Totalmphpmsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
mphpmquery="SELECT Count(distinct UIN) mphpm_students FROM Applicants where Program_type='MPH-PM' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphpmquery,conn
    
    Totalmphpm_Student = "Total number of students in MPH-PM : "&rs("mphpm_students")
    pdf.ChapterBody(Totalmphpm_Student)
    pdf.Ln(5)
    mphpmstu=rs("mphpm_students")
    rs.close

    
'////// MPH-Adv Students ////////
pdf.OrangeTitle("Program Option MPH-ADV")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// MPH-Adv CHF ////////////



set rs=Server.CreateObject("ADODB.recordset")
mphadv_chf_query="SELECT Count(distinct UIN) mphadv_chf_students FROM Applicants where Program_type='MPH-Adv' and Concentration='CHF' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadv_chf_query,conn

mphadv_chf_rows = rs("mphadv_chf_students")
mphadv_chf_cols = 5
Dim mphadv_chf_col(5)
mphadv_chf_col(1) = "Banner # "
mphadv_chf_col(2) = "First Name"
mphadv_chf_col(3) = "Last Name"
mphadv_chf_col(4) = "Admission Decision"
mphadv_chf_col(5) = "Confirmed"


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
					mphadvchf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-Adv' and Concentration='CHF'  and term_cd='"&AdmitTerm&"'"
					rs.Open mphadvchf_query,conn 
                  
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
    Total_MPH_Adv_CHF_Student = "Total number of students in CHF : "&mphadv_chf_rows
    pdf.ChapterBody(Total_MPH_Adv_CHF_Student)
    pdf.Ln(1)
pdf.GreyTitle("B. Concentration CHUD")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphadvchudquery="SELECT Count(distinct UIN) mphadv_chud_students FROM Applicants where Program_type='MPH-Adv' and Concentration='CHUD'  and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadvchudquery,conn
'//////// Courses Table ////////////
mphadvchud_rows = rs("mphadv_chud_students")
mphadv_chud_cols = 5
Dim mphadv_chud_col(5)
mphadv_chud_col(1) = "Banner # "
mphadv_chud_col(2) = "First Name"
mphadv_chud_col(3) = "Last Name"
mphadv_chud_col(4) = "Admission Decision"
mphadv_chud_col(5) = "Confirmed"


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
					mphadv_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-Adv' and Concentration='CHUD' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mphadv_chud,conn 
                  
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
                    
    pdf.Ln(1)
    rs.close
    Totalmphadvchud_Student = "Total number of students in CHUD : "&mphadvchud_rows
    pdf.ChapterBody(Totalmphadvchud_Student)
    pdf.Ln(1)
pdf.GreyTitle("C. Concentration MH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphadvmhquery="SELECT Count(distinct UIN) mphadv_mh_students FROM Applicants where Program_type='MPH-Adv' and Concentration='MH' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadvmhquery,conn
mphadvmh_rows = rs("mphadv_mh_students")
mphadv_mh_cols = 5
Dim mphadv_mh_col(5)
mphadv_mh_col(1) = "Banner # "
mphadv_mh_col(2) = "First Name"
mphadv_mh_col(3) = "Last Name"
mphadv_mh_col(4) = "Admission Decision"
mphadv_mh_col(5) = "Confirmed"

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
					mphadv_mh="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-Adv' and Concentration='MH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mphadv_mh,conn 
                  
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
pdf.ln(1)
    Totalmphadvmh_Student = "Total number of students in MH : "&mphadvmh_rows
    pdf.ChapterBody(Totalmphadvmh_Student)
    pdf.ln(1)
pdf.GreyTitle("D. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
mphadvschquery="SELECT Count(distinct UIN) mphadv_sch_students FROM Applicants where Program_type='MPH-Adv' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadvschquery,conn
mphadvsch_rows = rs("mphadv_sch_students")
mphadv_sch_cols = 5
Dim mphadv_sch_col(5)
mphadv_sch_col(1) = "Banner # "
mphadv_sch_col(2) = "First Name"
mphadv_sch_col(3) = "Last Name"
mphadv_sch_col(4) = "Admission Decision"
mphadv_sch_col(5) = "Confirmed"


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
					mphadv_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='MPH-Adv' and Concentration='SCH' and Term_CD like '"&AdmitTerm&"'"

					rs.Open mphadv_chud,conn 
                  
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
    Totalmphadvsch_Student = "Total number of students in SCH : "&mphadvsch_rows
    pdf.ChapterBody(Totalmphadvsch_Student)
    pdf.Ln(2)
set rs=Server.CreateObject("ADODB.recordset")
mphadvquery="SELECT Count(distinct UIN) mphadv_students FROM Applicants where Program_type='MPH-Adv' and Term_CD like '"&AdmitTerm&"' "
rs.Open mphadvquery,conn
    
    Totalmphadv_Student = "Total number of students in MPH-ADV : "&rs("mphadv_students")
    pdf.ChapterBody(Totalmphadv_Student)
    pdf.Ln(5)
    mphadvstu=rs("mphadv_students")
rs.close





    total_applied=pmstu+ftstu+advstu+mphpmstu+mphftstu+trstu+mphadvstu
pdf.GreyTitle("")
applied_students = "Total Students Applied : "&total_applied
pdf.ChapterBody(applied_students)
pdf.Ln(5)



pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
