<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Login_Check.asp"-->
<!--#include file="fpdf.asp"-->
<!--#include file="DBconn.asp"-->
<%

AdmitTerm=Request("term")
set rs=Server.CreateObject("ADODB.recordset")
termquery = "select distinct Admit_Term from PHDApplicants where Term_CD like '"&AdmitTerm&"'"


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

pdf.ChapterTitle2("                       Admissions Report - "&Termsel& "  Confirm   "&LastUpdatedDt&" "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(10)

rs.close

'////// Phd Students ////////
pdf.OrangeTitle("Program Option PhD")

set rs=Server.CreateObject("ADODB.recordset")
phd_query="SELECT Count(distinct UIN) phd_students FROM Applicants where Degree_Program='PHD' and confirmed='Y' and Term_CD like '"&AdmitTerm&"' "
rs.Open phd_query,conn

phd_rows = rs("phd_students")
phd_cols = 4
Dim phd_col(4)
phd_col(1) = "Banner # "
phd_col(2) = "First Name"
phd_col(3) = "Last Name"
phd_col(4) = "Confirmed"


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
					phd1_query="SELECT UIN, Firstname, LastName, Confirmed FROM Applicants where Degree_Program='PHD' and confirmed='Y' and term_cd='"&AdmitTerm&"'"
					rs.Open phd1_query,conn 
                  
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
    Total_PHD_Student = "Total number of students in PHD : "&phd_rows
    pdf.ChapterBody(Total_PHD_Student)
    pdf.Ln(1)



pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
