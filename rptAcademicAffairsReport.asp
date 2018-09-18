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

pdf.ChapterTitle2("                                   Academic Affairs Report            "  &LastUpdatedDt&"    "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(10)


'////// Deferred Students ///////

'pdf.FancyTable()
set rs=Server.CreateObject("ADODB.recordset")
currentstudents_query="SELECT Count(distinct UIN) current_students FROM CurrentStudents  "
rs.Open currentstudents_query,conn

current_rows = rs("current_students")
current_cols = 7
Dim current_col(7)
current_col(1) = "Last Name"
current_col(2) = "First Name"
current_col(3) = "Degree Program"
current_col(4) = "Program Type"
current_col(5) = "Concentration"
current_col(6) = "Advisor"
current_col(7) = "Track"

pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 30,30,30,25,27,30,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "LastName", "Firstname", "DegreeProgram","ProgramType","Concentration","Advisor","Track"
pdf.SetFont "Arial","",10
					
					rs.close
					set rs=Server.CreateObject("ADODB.recordset")
					currentstudent_query="SELECT LastName, Firstname, DegreeProgram,isnull(ProgramType,'') as ProgramType,isnull(Concentration,'') as Concentration,isnull(Advisor,'') as Advisor,isnull(Track,'') as Track FROM CurrentStudents where (Graduated != 'Y'  or Graduated is null) and Status not in ('LOA', 'TRANS','WDN') and Decision <> 'DF' order by LastName"
					rs.Open currentstudent_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                    i=1
                    Do While NOT rs.Eof 
                   
                    a = Replace(rs("LastName"),"|",",")
                    b= Replace(rs("Firstname"),"|",",")
                    c = Replace(rs("DegreeProgram"),"|",",")
                    d = Replace(rs("ProgramType"),"|",",")
                    e= Replace(rs("Concentration"),"|",",")
                    f = Replace(rs("Advisor"),"|",",")
                    g = Replace(rs("Track"),"|",",")
                    
                    
                    pdf.Row a,b,c,d,e,f,g
                    i=i+1
                    rs.MoveNext    
                    Loop
                    End If

 rs.close
'rs.Open query,conn
                    
    set rs=Server.CreateObject("ADODB.recordset")
totalquery="SELECT Count(distinct UIN) total_students FROM CurrentStudents where (Graduated != 'Y'  or Graduated is null) and Status not in ('LOA', 'TRANS','WDN') and Decision <> 'DF' "
rs.Open totalquery,conn
    
    Total_Student = "Total number of students : "&rs("total_students")
    pdf.ChapterBody(Total_Student)
    pdf.Ln(7)
    rs.close

pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(7)

pdf.Close()
pdf.Output()
conn.close

%>
