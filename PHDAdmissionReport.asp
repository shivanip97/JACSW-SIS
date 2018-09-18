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

pdf.ChapterTitle2("                              Admissions Report - "&Termsel& "    "&LastUpdatedDt&" "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(10)

rs.close

               ' //// PHD ////
pdf.OrangeTitle("PhD")

mswphd_rows = 9
mswphd_cols = 3

Dim mswphd_col(3)

mswphd_col(1) = ""
mswphd_col(2) = "Total"
mswphd_col(3) = "Target"
'totalmsw_rows(2) = "Applicants"
'totalmsw_rows(3) = "Confirmed"

'totalmsw_rows(4) = "Denials"
'totalmsw_rows(5) = "Withdrawals"
'totalmsw_rows(6) = "No Decision"
'totalmsw_rows(7) = "Incomplete"
'totalmsw_rows(8) = "Deferment"
'totalmsw_rows(9) = "Wait List"


pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 35,50,60,30
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "","Total", "Target"
pdf.SetFont "Arial","",10

set rs=Server.CreateObject("ADODB.recordset")
					mswphdapplicants_query="SELECT Count (distinct UIN) number FROM PHDApplicants where Degree_Program='PHD' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswphdapplicants_query,conn 
    mswPHDapplicants = rs("number")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Applicants","|",",")
                    c = Replace(rs("number"),"|",",")
                    d = Replace("","|",",")
                    
                    pdf.Row b,c,d
                   
                    End If

 rs.close
set rs=Server.CreateObject("ADODB.recordset")
					mswphdaccepted_query="SELECT Count (Admission_decision) status, round(Count(Admission_decision)* 100 /Cast("&mswPHDapplicants&" as float), 2) percentaccepted FROM PHDApplicants where Admission_decision IN ('A', 'S', 'ReAdmit') and Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswphdaccepted_query,conn 
                  mswPHDaccepted = rs("status")
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                   
                    b= Replace("Accepted","|",",")
                    c = Replace(rs("status"),"|",",")
                    d = Replace(rs("percentaccepted"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
				
     If mswPHDaccepted <> 0 Then 
   
					mswphdconfirmed_query="SELECT Count (Confirmed) confirm, round(Count(Confirmed)* 100 /Cast( "&mswPHDaccepted&" as float),2) percentconfirmed FROM PHDApplicants where Confirmed = 'Y' and Degree_Program = 'PHD' and term_cd='"&AdmitTerm&"'"
					rs.Open mswphdconfirmed_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
        
                    b= Replace("Confirmed","|",",")
                    c = Replace(rs("confirm"),"|",",")
                    d = Replace(rs("percentconfirmed"),"|",",")+("% of all accepted")
                    
                    pdf.Row b,c,d
                   
                   End If
    Else

                  mswphdconfirmed_query="SELECT Count (Confirmed) confirm FROM PHDApplicants where Confirmed = 'Y' and Degree_Program = 'PHD' and term_cd='"&AdmitTerm&"'"
					rs.Open mswphdconfirmed_query,conn 
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Confirmed","|",",")
                    c = Replace(rs("confirm"),"|",",")
                    d = Replace("0","|",",")+("% of all accepted")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If
    
 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
					mswphddenied_query="SELECT Count (Admission_decision) denied,round(Count(Admission_decision)* 100 / Cast("&mswPHDapplicants&" as float),2) percentdenied FROM PHDApplicants where Admission_decision = 'D' and Degree_Program = 'PHD' and term_cd='"&AdmitTerm&"'"
					rs.Open mswphddenied_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Denials","|",",")
                    c = Replace(rs("denied"),"|",",")
                    d = Replace(rs("percentdenied"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If

 rs.close
    'set rs=Server.CreateObject("ADODB.recordset")
	'				mswphdwithdrawn_query="SELECT Count (Withdrawn) withdraw, round(Count (Withdrawn)*100/Cast("&mswPHDapplicants&"as float), 2) percentwithdraw FROM PHDApplicants where Withdrawn = 'Y' and Degree_Program = 'PHD' and term_cd='"&AdmitTerm&"'"
	'				rs.Open mswphdwithdrawn_query,conn 
                  
                    'If rs.EOF Then
                      
                   ' Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Withdrawals","|",",")
                    c = Replace("0","|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                   ' End If

 'rs.close
    set rs=Server.CreateObject("ADODB.recordset")
					mswphdnodecision_query="SELECT Count (Admission_decision) nodecision, round( Count (Admission_decision)*100/Cast("&mswPHDapplicants&" as float),2) percentnodecision FROM PHDApplicants where Admission_decision = '' and Degree_Program = 'PHD' and term_cd='"&AdmitTerm&"'"
					rs.Open mswphdnodecision_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("No Decision","|",",")
                    c = Replace(rs("nodecision"),"|",",")
                    d = Replace(rs("percentnodecision"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
					mswphdincomplete_query="SELECT Count (Application_Status) incomplete, round(Count (Application_Status) * 100/ Cast("&mswPHDapplicants&" as float),2) percentincomplete FROM PHDApplicants where Application_Status = 'IN-Incomplete' and Degree_Program = 'PHD' and term_cd='"&AdmitTerm&"'"
					rs.Open mswphdincomplete_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Incomplete","|",",")
                    c = Replace(rs("incomplete"),"|",",")
                    d = Replace(rs("percentincomplete"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
					mswphddefer_query="SELECT Count (Admission_decision) deferment, round(Count (Admission_decision) * 100/Cast( "&mswPHDapplicants&" as float),2) percentdefer FROM PHDApplicants where Admission_decision = 'DF' and Degree_Program = 'PHD' and term_cd='"&AdmitTerm&"'"
					rs.Open mswphddefer_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Deferment","|",",")
                    c = Replace(rs("deferment"),"|",",")
                    d = Replace(rs("percentdefer"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If

 rs.close
        set rs=Server.CreateObject("ADODB.recordset")
					mswphdwaitlist_query="SELECT  Count (Admission_decision) waitlist, round(Count (Admission_decision) * 100/Cast ("&mswPHDapplicants&" as float),2) percentwaitlist FROM PHDApplicants where Admission_decision = 'W' and Degree_Program = 'PHD' and term_cd='"&AdmitTerm&"'"
					rs.Open mswphdwaitlist_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Wait List","|",",")
                    c = Replace(rs("waitlist"),"|",",")
                    d = Replace(rs("percentwaitlist"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If

 rs.close  




    pdf.Ln(10)
pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
