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

pdf.ChapterTitle2("                       Current Students 10 day Withdrawal Report    "&LastUpdatedDt&"  "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(10)
startdate = Request("ID")

    set rs=Server.CreateObject("ADODB.recordset")
    cnf_students_blank_query="SELECT Count(distinct UIN) blank_cnf_students FROM CurrentStudents where DATEDIFF(day, '" & startdate & "', WithdrawnDate) >= 10"
    Response.Write (cnf_students_blank_query)
    rs.Open cnf_students_blank_query,conn
    If rs("blank_cnf_students") <> 0 Then
    pdf.OrangeTitle("Withdrawal Report")
'pdf.FancyTable()

'//////// Students ////////////

blank_cnf_rows = rs("blank_cnf_students")
blank_cnf_cols = 5
Dim blank_cnf_col(5)

blank_cnf_col(1) = "UIN"
blank_cnf_col(2) = "Last Name"
blank_cnf_col(3) = "First Name"
blank_cnf_col(4) = "Admit Date"
blank_cnf_col(5) = "Withdrawn Date"

pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,45,45,30,30
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "UIN", "Last Name", "First Name","Admit Date", "Withdrawn Date"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					blankcnf_query="SELECT UIN, LastName, Firstname, case AdmitTerm when 'Fall 2016' Then '2016-08-15' end as AdmitDate, WithdrawnDate FROM CurrentStudents where DATEDIFF(day, '" & startdate & "', WithdrawnDate) >= 10 order by LastName"
					rs.Open blankcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("UIN"),"|",",")
                    b = Replace(rs("LastName"),"|",",")
                    c = Replace(rs("Firstname"),"|",",")
                    d = Replace(rs("AdmitDate"),"|",",")
                    e = Replace(rs("WithdrawnDate"),"|",",")
                   
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
  
                      
pdf.Ln(5)
     set rs=Server.CreateObject("ADODB.recordset")
totalquery="SELECT Count(distinct UIN) total_students FROM CurrentStudents where DATEDIFF(day,  '" & startdate & "', WithdrawnDate) >= 10 "
rs.Open totalquery,conn
 Total_Student = "Total number of students: "&rs("total_students")
    pdf.ChapterBody(Total_Student)
pdf.Ln(10)
    rs.close
pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)
''" & startdate & "', WithdrawnDate
pdf.Close()
pdf.Output()
conn.close
%>