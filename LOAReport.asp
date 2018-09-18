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

pdf.ChapterTitle2("                            Current Students LOA Report    "&LastUpdatedDt&"   "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(10)


    set rs=Server.CreateObject("ADODB.recordset")
    cnf_students_blank_query="SELECT Count(distinct UIN) blank_cnf_students FROM CurrentStudents where (Graduated != 'Y'  or Graduated is null) and Status = 'LOA' and Decision <> 'DF' "
    Response.Write (cnf_students_blank_query)
    rs.Open cnf_students_blank_query,conn
    If rs("blank_cnf_students") <> 0 Then
    pdf.OrangeTitle("LOA Report")
'pdf.FancyTable()

'//////// Students ////////////

blank_cnf_rows = rs("blank_cnf_students")
blank_cnf_cols = 7
Dim blank_cnf_col(7)
blank_cnf_col(1) = "Degree Program"
blank_cnf_col(2) = "UIN"
blank_cnf_col(3) = "Last Name"
blank_cnf_col(4) = "First Name"
blank_cnf_col(5) = "Email"
blank_cnf_col(6) = "Concentration"
blank_cnf_col(7) = "Track"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 25,20,35,35,35,28,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Degree Program","UIN", "Last Name", "First Name","Email", "Concentration", "Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					blankcnf_query="SELECT DegreeProgram, UIN, LastName, Firstname, EMail,Concentration, Track FROM CurrentStudents where  (Graduated != 'Y'  or Graduated is null) and Status = 'LOA' and Decision <> 'DF' order by LastName"
					rs.Open blankcnf_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("DegreeProgram"),"|",",")
                    b = Replace(rs("UIN"),"|",",")
                    c = Replace(rs("LastName"),"|",",")
                    d = Replace(rs("Firstname"),"|",",")
                    e = Replace(rs("EMail"),"|",",")
                    f = Replace(rs("Concentration"),"|",",")
                    g = Replace(rs("Track"),"|",",")
                    pdf.Row a,b,c,d,e,f,g
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
totalquery="SELECT Count(distinct UIN) total_students FROM CurrentStudents where (Graduated != 'Y'  or Graduated is null) and Status = 'LOA' and Decision <> 'DF' "
rs.Open totalquery,conn
 Total_Student = "Total number of students: "&rs("total_students")
    pdf.ChapterBody(Total_Student)
pdf.Ln(10)
    rs.close
pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close
%>