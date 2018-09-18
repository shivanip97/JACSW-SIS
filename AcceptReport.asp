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

pdf.ChapterTitle2("             Report 6 -Admissions Report - "&Termsel& " - Accept   "  &LastUpdatedDt&"  "&LastUpdatedTime)
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
adv_chf_query="SELECT Count(distinct UIN) adv_chf_students FROM Applicants where Program_type='Adv' and Concentration='CHF' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
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
					advchf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='Adv' and Concentration='CHF' and Admission_decision in ('A','S','ReAdmit') and term_cd='"&AdmitTerm&"' order by  Confirmed desc,LastName asc"
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
advchudquery="SELECT Count(distinct UIN) adv_chud_students FROM Applicants where Program_type='Adv' and Concentration='CHUD' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
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
					adv_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='Adv' and Concentration='CHUD' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' order by Confirmed desc,LastName asc"

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
advmhquery="SELECT Count(distinct UIN) adv_mh_students FROM Applicants where Program_type='Adv' and Concentration='MH' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
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
					adv_mh="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='Adv' and Concentration='MH' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' order by Confirmed desc,LastName asc"

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
advschquery="SELECT Count(distinct UIN) adv_sch_students FROM Applicants where Program_type='Adv' and Concentration='SCH' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
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
					adv_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='Adv' and Concentration='SCH' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' order by Confirmed desc,LastName asc"

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

pdf.GreyTitle("E. Without Concentration")
'pdf.FancyTable()

'//////// Adv CHF ////////////



set rs=Server.CreateObject("ADODB.recordset")
adv_blank_query="SELECT Count(distinct UIN) adv_blank_students FROM Applicants where Program_type='Adv' and Concentration='' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open adv_blank_query,conn

adv_blank_rows = rs("adv_blank_students")
adv_blank_cols = 5
Dim adv_blank_col(5)
adv_blank_col(1) = "Banner # "
adv_blank_col(2) = "First Name"
adv_blank_col(3) = "Last Name"
adv_blank_col(4) = "Admission Decision"
adv_blank_col(5) = "Confirmed"


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
					advblank_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='Adv' and Concentration='' and Admission_decision in ('A','S','ReAdmit') and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open advblank_query,conn 
                  
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
    Total_Adv_blank_Student = "Total number of students with no concentration : "&adv_blank_rows
    pdf.ChapterBody(Total_Adv_blank_Student)
    pdf.Ln(1)
set rs=Server.CreateObject("ADODB.recordset")
advquery="SELECT Count(distinct UIN) adv_students FROM Applicants where Program_type='Adv' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open advquery,conn
    
    Totaladv_Student = "Total number of students in ADV : "&rs("adv_students")
    pdf.ChapterBody(Totaladv_Student)
    pdf.Ln(5)
    advstu=rs("adv_students")
rs.close



'//////// FT students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
ft_chf_query="SELECT Count(distinct UIN) ft_chf_students FROM Applicants where Program_type='FT' and Concentration='CHF' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open ft_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option FT")

pdf.GreyTitle("A. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
ftschquery="SELECT Count(distinct UIN) ft_sch_students FROM Applicants where Program_type='FT' and Concentration='SCH' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
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
					ft_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='FT' and Concentration='SCH' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' order by Confirmed desc,LastName asc"

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
pdf.GreyTitle("B. Without Concentration")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
ftblankquery="SELECT Count(distinct UIN) ft_blank_students FROM Applicants where Program_type='FT' and Concentration='' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open ftblankquery,conn
'//////// FT CHUD ////////////
ftblank_rows = rs("ft_blank_students")
ft_blank_cols = 5
Dim ft_blank_col(5)
ft_blank_col(1) = "Banner # "
ft_blank_col(2) = "First Name"
ft_blank_col(3) = "Last Name"
ft_blank_col(4) = "Admission Decision"
ft_blank_col(5) = "Confirmed"


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
					ft_blank="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='FT' and Concentration='' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' order by Confirmed desc,LastName asc"

					rs.Open ft_blank,conn 
                  
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
    Totalftblank_Student = "Total number of students without concentration : "&ftblank_rows
    pdf.ChapterBody(Totalftblank_Student)
    pdf.Ln(1)
set rs=Server.CreateObject("ADODB.recordset")
ftquery="SELECT Count(distinct UIN) ft_students FROM Applicants where Program_type='FT' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open ftquery,conn
    
    Totalft_Student = "Total number of students in FT : "&rs("ft_students")
    pdf.ChapterBody(Totalft_Student)
    pdf.Ln(5)
    ftstu=rs("ft_students")
    rs.close

    

'//////// PM students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
pm_chf_query="SELECT Count(distinct UIN) pm_chf_students FROM Applicants where Program_type='PM' and Concentration='CHF' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open pm_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option PM")

pdf.GreyTitle("A. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
pmschquery="SELECT Count(distinct UIN) pm_sch_students FROM Applicants where Program_type='PM' and Concentration='SCH' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
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
					pm_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='PM' and Concentration='SCH' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' order by Confirmed desc,LastName asc"

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
    pdf.GreyTitle("B. Without Concentration")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
pm_blank_query="SELECT Count(distinct UIN) pm_blank_students FROM Applicants where Program_type='PM' and Concentration='' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open pm_blank_query,conn
pm_blank_rows = rs("pm_blank_students")
pm_blank_cols = 5
Dim pm_blank_col(5)
pm_blank_col(1) = "Banner # "
pm_blank_col(2) = "First Name"
pm_blank_col(3) = "Last Name"
pm_blank_col(4) = "Admission Decision"
pm_blank_col(5) = "Confirmed"

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
					pmblank_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type='PM' and Concentration='' and Admission_decision in ('A','S','ReAdmit') and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
					rs.Open pmblank_query,conn 
                  
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
    Total_PM_blank_Student = "Total number of students without concentration : "&pm_blank_rows
    pdf.ChapterBody(Total_PM_blank_Student)
    pdf.Ln(1)
set rs=Server.CreateObject("ADODB.recordset")
pmquery="SELECT Count(distinct UIN) pm_students FROM Applicants where Program_type='PM' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open pmquery,conn
    
    Totalpm_Student = "Total number of students in PM : "&rs("pm_students")
    pdf.ChapterBody(Totalpm_Student)
    pdf.Ln(5)
    pmstu=rs("pm_students")
    rs.close



    '//////// PM students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
tr_chf_query="SELECT Count(distinct UIN) tr_chf_students FROM Applicants where Program_type like 'TR%' and Concentration='CHF' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open tr_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option TR")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// PM CHF ////////////
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
					trchf_query="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type like 'TR%' and Concentration='CHF' and Admission_decision in ('A','S','ReAdmit') and term_cd='"&AdmitTerm&"' order by Confirmed desc,LastName asc"
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
trchudquery="SELECT Count(distinct UIN) tr_chud_students FROM Applicants where Program_type like 'TR%' and Concentration='CHUD' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open trchudquery,conn
'//////// PM CHUD ////////////
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
					tr_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type like 'TR%' and Concentration='CHUD' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' order by Confirmed desc,LastName asc"

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
trmhquery="SELECT Count(distinct UIN) tr_mh_students FROM Applicants where Program_type like 'TR%' and Concentration='MH' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
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
					tr_mh="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type like 'TR%' and Concentration='MH' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' order by Confirmed desc,LastName asc"

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
trschquery="SELECT Count(distinct UIN) tr_sch_students FROM Applicants where Program_type like 'TR%' and Concentration='SCH' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
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
					tr_chud="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type like 'TR%' and Concentration='SCH' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' order by Confirmed desc,LastName asc"

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
    pdf.GreyTitle("E. Without Concentration")

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
trblankquery="SELECT Count(distinct UIN) tr_blank_students FROM Applicants where Program_type like 'TR%' and Concentration='' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open trblankquery,conn
'//////// PM CHUD ////////////
trblank_rows = rs("tr_blank_students")
tr_blank_cols = 5
Dim tr_blank_col(5)
tr_blank_col(1) = "Banner # "
tr_blank_col(2) = "First Name"
tr_blank_col(3) = "Last Name"
tr_blank_col(4) = "Admission Decision"
tr_blank_col(5) = "Confirmed"


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
					tr_blank="SELECT UIN, Firstname, LastName, Admission_decision, Confirmed FROM Applicants where Program_type like 'TR%' and Concentration='' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' order by Confirmed desc,LastName asc"

					rs.Open tr_blank,conn 
                  
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
    Totaltrblank_Student = "Total number of students without concentration : "&trblank_rows
    pdf.ChapterBody(Totaltrblank_Student)
    pdf.Ln(1)
set rs=Server.CreateObject("ADODB.recordset")
trquery="SELECT Count(distinct UIN) tr_students FROM Applicants where Program_type like 'TR%' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open trquery,conn
    
    Totaltr_Student = "Total number of students in TR : "&rs("tr_students")
    pdf.ChapterBody(Totaltr_Student)
    pdf.Ln(5)
    trstu=rs("tr_students")
    rs.close





    total_accepted=pmstu+ftstu+advstu+trstu
pdf.GreyTitle("")
accepted_students = "Total Students Accepted : "&total_accepted
pdf.ChapterBody(accepted_students)
pdf.Ln(3)

set rs=Server.CreateObject("ADODB.recordset")
pmquery="SELECT Count(distinct UIN) conf_students FROM Applicants where Admission_decision in ('A','S','ReAdmit') and Confirmed='Y' and Term_CD like '"&AdmitTerm&"' "
rs.Open pmquery,conn

Totalconf_Student = "Total Students Confirmed : "&rs("conf_students")
    pdf.ChapterBody(Totalconf_Student)
    pdf.Ln(5)
    rs.close

pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
