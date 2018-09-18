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

pdf.ChapterTitle2("                           ADV Admissions Report - Accept   "  &LastUpdatedDt&"  "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(10)




'////// Adv Students ////////
pdf.OrangeTitle("Program Option ADV")
pdf.GreyTitle("A. Concentration CHF")
'pdf.FancyTable()

'//////// Adv CHF ////////////



set rs=Server.CreateObject("ADODB.recordset")
adv_chf_query="SELECT Count(distinct UIN) adv_chf_students FROM CurrentStudents where ProgramType='Adv' and Concentration='CHF' "
rs.Open adv_chf_query,conn

adv_chf_rows = rs("adv_chf_students")
adv_chf_cols = 5
Dim adv_chf_col(5)
adv_chf_col(1) = "Banner # "
adv_chf_col(2) = "Last Name"
adv_chf_col(3) = "First Name"
adv_chf_col(4) = "Admit Term"
adv_chf_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","Last Name","First name", "Admit Term", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					advchf_query="SELECT UIN, LastName, Firstname, AdmitTerm, Confirmed FROM CurrentStudents where ProgramType='Adv' and Concentration='CHF' order by LastName"
					rs.Open advchf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")                    
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("AdmitTerm"),"|",",")
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
advchudquery="SELECT Count(distinct UIN) adv_chud_students FROM CurrentStudents where ProgramType='Adv' and Concentration='CHUD' "
rs.Open advchudquery,conn
'//////// Courses Table ////////////
advchud_rows = rs("adv_chud_students")
adv_chud_cols = 5
Dim adv_chud_col(5)
adv_chud_col(1) = "Banner # "
adv_chud_col(2) = "Last Name"
adv_chud_col(3) = "First Name"
adv_chud_col(4) = "Admit Term"
adv_chud_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","Last Name", "First name","Admit Term", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					adv_chud="SELECT UIN, LastName,Firstname, AdmitTerm, Confirmed FROM CurrentStudents where ProgramType='Adv' and Concentration='CHUD' order by LastName"

					rs.Open adv_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")                    
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("AdmitTerm"),"|",",")
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
advmhquery="SELECT Count(distinct UIN) adv_mh_students FROM CurrentStudents where ProgramType='Adv' and Concentration='MH' "
rs.Open advmhquery,conn
advmh_rows = rs("adv_mh_students")
adv_mh_cols = 5
Dim adv_mh_col(5)
adv_mh_col(1) = "Banner # "
adv_mh_col(2) = "Last Name"
adv_mh_col(3) = "First Name"
adv_mh_col(4) = "Admit Term"
adv_mh_col(5) = "Confirmed"

pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","Last Name", "First name","Admit Term", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					adv_mh="SELECT UIN, LastName, Firstname, AdmitTerm, Confirmed FROM CurrentStudents where ProgramType='Adv' and Concentration='MH' order by LastName"

					rs.Open adv_mh,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("AdmitTerm"),"|",",")
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
advschquery="SELECT Count(distinct UIN) adv_sch_students FROM CurrentStudents where ProgramType='Adv' and Concentration='SCH'  "
rs.Open advschquery,conn
advsch_rows = rs("adv_sch_students")
adv_sch_cols = 5
Dim adv_sch_col(5)
adv_sch_col(1) = "Banner # "
adv_sch_col(2) = "Last Name"
adv_sch_col(3) = "First Name"
adv_sch_col(4) = "Admit Term"
adv_sch_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","Last Name","First name", "Admit Term", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					adv_chud="SELECT UIN, LastName, Firstname,  AdmitTerm, Confirmed FROM CurrentStudents where ProgramType='Adv' and Concentration='SCH' order by LastName"

					rs.Open adv_chud,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("AdmitTerm"),"|",",")
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
adv_blank_query="SELECT Count(distinct UIN) adv_blank_students FROM CurrentStudents where ProgramType='Adv' and Concentration='' "
rs.Open adv_blank_query,conn

adv_blank_rows = rs("adv_blank_students")
adv_blank_cols = 5
Dim adv_blank_col(5)
adv_blank_col(1) = "Banner # "
adv_blank_col(2) = "Last Name"
adv_blank_col(3) = "First Name"
adv_blank_col(4) = "Admit Term"
adv_blank_col(5) = "Confirmed"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,50,40,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","Last Name", "First name","Admit Term", "Confirmed"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					advblank_query="SELECT UIN, LastName, Firstname, AdmitTerm, Confirmed FROM CurrentStudents where ProgramType='Adv' and Concentration='' order by LastName"
					rs.Open advblank_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("AdmitTerm"),"|",",")
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
advquery="SELECT Count(distinct UIN) adv_students FROM CurrentStudents where ProgramType='Adv' "
rs.Open advquery,conn
    
    Totaladv_Student = "Total number of students in ADV : "&rs("adv_students")
    pdf.ChapterBody(Totaladv_Student)
    pdf.Ln(5)
    advstu=rs("adv_students")
rs.close



pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
