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

pdf.ChapterTitle2("            Report 3 - Admissions Report - "&Termsel& "  Confirm   "&LastUpdatedDt&"  "&LastUpdatedTime)
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
adv_chf_query="SELECT Count(distinct UIN) adv_chf_students FROM Applicants where Program_type='Adv' and Concentration='CHF' and Admission_decision in ('A','S','ReAdmit') and confirmed='Y' and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' "
rs.Open adv_chf_query,conn

adv_chf_rows = rs("adv_chf_students")
adv_chf_cols = 4
Dim adv_chf_col(4)
adv_chf_col(1) = "Banner # "
adv_chf_col(2) = "First Name"
adv_chf_col(3) = "Last Name"
adv_chf_col(4) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					advchf_query="SELECT UIN, Firstname, LastName, Confirmed FROM Applicants where Program_type='Adv' and Concentration='CHF' and Admission_decision in ('A','S','ReAdmit') and confirmed='Y' and Withdrawn <> 'Y' and term_cd='"&AdmitTerm&"' order by Lastname"
					rs.Open advchf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d
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
advchudquery="SELECT Count(distinct UIN) adv_chud_students FROM Applicants where Program_type='Adv' and Concentration='CHUD' and confirmed='Y' and Admission_decision in ('A','S','ReAdmit') and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' "
rs.Open advchudquery,conn
'//////// Courses Table ////////////
advchud_rows = rs("adv_chud_students")
adv_chud_cols = 4
Dim adv_chud_col(4)
adv_chud_col(1) = "Banner # "
adv_chud_col(2) = "First Name"
adv_chud_col(3) = "Last Name"
adv_chud_col(4) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					adv_chud="SELECT UIN, Firstname, LastName, Confirmed FROM Applicants where Program_type='Adv' and Concentration='CHUD' and confirmed='Y' and Admission_decision in ('A','S','ReAdmit') and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' order by Lastname"

					rs.Open adv_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d
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
advmhquery="SELECT Count(distinct UIN) adv_mh_students FROM Applicants where Program_type='Adv' and Concentration='MH' and confirmed='Y' and Withdrawn <> 'Y' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open advmhquery,conn
advmh_rows = rs("adv_mh_students")
adv_mh_cols = 4
Dim adv_mh_col(4)
adv_mh_col(1) = "Banner # "
adv_mh_col(2) = "First Name"
adv_mh_col(3) = "Last Name"
adv_mh_col(4) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					adv_mh="SELECT UIN, Firstname, LastName, Confirmed FROM Applicants where Program_type='Adv' and Concentration='MH' and confirmed='Y'and Withdrawn <> 'Y' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' order by Lastname"

					rs.Open adv_mh,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d
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
advschquery="SELECT Count(distinct UIN) adv_sch_students FROM Applicants where Program_type='Adv' and Concentration='SCH' and confirmed='Y' and Withdrawn <> 'Y' and Admission_decision in ('A','S','ReAdmit') and Term_CD like '"&AdmitTerm&"' "
rs.Open advschquery,conn
advsch_rows = rs("adv_sch_students")
adv_sch_cols = 4
Dim adv_sch_col(4)
adv_sch_col(1) = "Banner # "
adv_sch_col(2) = "First Name"
adv_sch_col(3) = "Last Name"
adv_sch_col(4) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					adv_chud="SELECT UIN, Firstname, LastName, Confirmed FROM Applicants where Program_type='Adv' and Concentration='SCH' and confirmed='Y' and Admission_decision in ('A','S','ReAdmit') and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' order by Lastname"

					rs.Open adv_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
    Totaladvsch_Student = "Total number of students in SCH : "&advsch_rows
    pdf.ChapterBody(Totaladvsch_Student)
    pdf.Ln(2)

set rs=Server.CreateObject("ADODB.recordset")
advquery="SELECT Count(distinct UIN) adv_students FROM Applicants where Program_type='Adv' and confirmed='Y' and Admission_decision in ('A','S','ReAdmit') and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' "
rs.Open advquery,conn
    
    Totaladv_Student = "Total number of students in ADV : "&rs("adv_students")
    pdf.ChapterBody(Totaladv_Student)
    pdf.Ln(5)
rs.close



'//////// FT students confirmed ////////////

pdf.OrangeTitle("Program Option FT")

pdf.GreyTitle("A. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
ftschquery="SELECT Count(distinct UIN) ft_sch_students FROM Applicants where Program_type='FT' and Concentration='SCH' and confirmed='Y' and Admission_decision in ('A','S','ReAdmit') and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' "
rs.Open ftschquery,conn
ftsch_rows = rs("ft_sch_students")
ft_sch_cols = 4
Dim ft_sch_col(4)
ft_sch_col(1) = "Banner # "
ft_sch_col(2) = "First Name"
ft_sch_col(3) = "Last Name"
ft_sch_col(4) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					ft_chud="SELECT UIN, Firstname, LastName, Confirmed FROM Applicants where Program_type='FT' and Concentration='SCH' and Admission_decision in ('A','S','ReAdmit') and confirmed='Y' and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' order by Lastname"

					rs.Open ft_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d
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
ftnoquery="SELECT Count(distinct UIN) ft_no_students FROM Applicants where Program_type='FT' and Concentration='' and confirmed='Y' and Admission_decision in ('A','S','ReAdmit') and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' "
rs.Open ftnoquery,conn
ftno_rows = rs("ft_no_students")
ft_no_cols = 4
Dim ft_no_col(4)
ft_no_col(1) = "Banner # "
ft_no_col(2) = "First Name"
ft_no_col(3) = "Last Name"
ft_no_col(4) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					ft_no="SELECT UIN, Firstname, LastName, Confirmed FROM Applicants where Program_type='FT' and Concentration='' and confirmed='Y' and Admission_decision in ('A','S','ReAdmit') and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' order by Lastname"

					rs.Open ft_no,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
pdf.ln(1)
    Totalftno_Student = "Total number of students with no Concentration : "&ftno_rows
    pdf.ChapterBody(Totalftno_Student)
    pdf.ln(1)
set rs=Server.CreateObject("ADODB.recordset")
ftquery="SELECT Count(distinct UIN) ft_students FROM Applicants where Program_type='FT' and confirmed='Y' and Admission_decision in ('A','S','ReAdmit') and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' "
rs.Open ftquery,conn
    
    Totalft_Student = "Total number of students in FT : "&rs("ft_students")
    pdf.ChapterBody(Totalft_Student)
    pdf.Ln(5)
    rs.close

 

'//////// PM students confirmed ////////////

set rs=Server.CreateObject("ADODB.recordset")
pm_chf_query="SELECT Count(distinct UIN) pm_chf_students FROM Applicants where Program_type='PM' and Concentration='CHF' and Admission_decision in ('A','S','ReAdmit') and confirmed='Y' and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' "
rs.Open pm_chf_query,conn
'pdf.ChapterBody(Total_Student)



pdf.OrangeTitle("Program Option PM")

pdf.GreyTitle("A. Concentration SCH")
'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
pmschquery="SELECT Count(distinct UIN) pm_sch_students FROM Applicants where Program_type='PM' and Concentration='SCH' and Admission_decision in ('A','S','ReAdmit') and confirmed='Y' and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' "
rs.Open pmschquery,conn
pmsch_rows = rs("pm_sch_students")
pm_sch_cols = 4
Dim pm_sch_col(4)
pm_sch_col(1) = "Banner # "
pm_sch_col(2) = "First Name"
pm_sch_col(3) = "Last Name"
pm_sch_col(4) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					pm_chud="SELECT UIN, Firstname, LastName, Confirmed FROM Applicants where Program_type='PM' and Concentration='SCH' and Admission_decision in ('A','S','ReAdmit') and confirmed='Y' and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' order by Lastname"

					rs.Open pm_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d
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
pmnoquery="SELECT Count(distinct UIN) pm_no_students FROM Applicants where Program_type='PM' and Concentration='' and confirmed='Y' and Admission_decision in ('A','S','ReAdmit') and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' "
rs.Open pmnoquery,conn
pmno_rows = rs("pm_no_students")
pm_no_cols = 4
Dim pm_no_col(4)
pm_no_col(1) = "Banner # "
pm_no_col(2) = "First Name"
pm_no_col(3) = "Last Name"
pm_no_col(4) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,95,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					pm_no="SELECT UIN, Firstname, LastName, Confirmed FROM Applicants where Program_type='PM' and Concentration='' and Admission_decision in ('A','S','ReAdmit') and confirmed='Y' and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' order by Lastname"

					rs.Open pm_no,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Confirmed"),"|",",")
                    
                    pdf.Row a,b,c,d
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If
                    

rs.close
pdf.ln(1)
    Totalpmno_Student = "Total number of students with no Concentration : "&pmno_rows
    pdf.ChapterBody(Totalpmno_Student)
    pdf.ln(1)
set rs=Server.CreateObject("ADODB.recordset")
pmquery="SELECT Count(distinct UIN) pm_students FROM Applicants where Program_type='PM' and confirmed='Y' and Admission_decision in ('A','S','ReAdmit') and Withdrawn <> 'Y' and Term_CD like '"&AdmitTerm&"' "
rs.Open pmquery,conn
    
    Totalpm_Student = "Total number of students in PM : "&rs("pm_students")
    pdf.ChapterBody(Totalpm_Student)
    pdf.Ln(5)
    rs.close




pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
