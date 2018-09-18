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

pdf.ChapterTitle2("                  Report 15 - Deferred Students - "&Termsel& "      "  &LastUpdatedDt&" "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(10)

rs.close
'////// Deferred Students ///////

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
deferred_query="SELECT Count(distinct UIN) deferred_students FROM Applicants where Admission_decision = 'DF' and Term_CD like '"&AdmitTerm&"' "
rs.Open deferred_query,conn

deferred_rows = rs("deferred_students")
deferred_cols = 3
Dim deferred_col(3)
deferred_col(1) = "Banner # "
deferred_col(2) = "First Name"
deferred_col(3) = "Last Name"

pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 30,95,50
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Banner #","First name","Last Name"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					deferredstudent_query="SELECT UIN, Firstname, LastName FROM Applicants where Admission_decision = 'DF' and term_cd='"&AdmitTerm&"' order by LastName"
					rs.Open deferredstudent_query,conn 
                  
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
                    
    set rs=Server.CreateObject("ADODB.recordset")
totalquery="SELECT Count(distinct UIN) total_students FROM Applicants where Admission_decision = 'DF' and Term_CD like '"&AdmitTerm&"' "
rs.Open totalquery,conn
    
    Total_Student = "Total number of students Deferred : "&rs("total_students")
    pdf.ChapterBody(Total_Student)
    pdf.Ln(5)
    rs.close

pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
