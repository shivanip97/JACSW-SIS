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

pdf.ChapterTitle2("   Report 1 - Overall Summary Status - Admissions Report - "&Termsel& "   "  &LastUpdatedDt&" "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(10)

rs.close

   ' //// Total MSW ////
pdf.OrangeTitle("Total MSW")

totalmsw_rows = 9
totalmsw_cols = 3

Dim totalmsw_col(3)

totalmsw_col(1) = ""
totalmsw_col(2) = "Total"
totalmsw_col(3) = "Target"
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
					mswapplicants_query="SELECT Count (distinct UIN) number FROM Applicants where Degree_Program='MSW' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswapplicants_query,conn 
    mswapplicants = rs("number")
                  
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
     If mswapplicants <> 0 Then 
					mswaccepted_query="SELECT Count (distinct UIN) status, round(Count(distinct UIN)* 100 /Cast("&mswapplicants&" as float), 2) percentaccepted FROM Applicants where Admission_decision IN ('A', 'S', 'ReAdmit') and Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswaccepted_query,conn 
                  mswaccepted = rs("status")
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                   
                    b= Replace("Accepted","|",",")
                    c = Replace(rs("status"),"|",",")
                    d = Replace(rs("percentaccepted"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
    mswaccepted_query="SELECT Count (distinct UIN) status FROM Applicants where Admission_decision IN ('A', 'S', 'ReAdmit') and Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswaccepted_query,conn 
                  mswaccepted = rs("status")
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                   
                    b= Replace("Accepted","|",",")
                    c = Replace(rs("status"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswaccepted <> 0 Then 
   
					mswconfirmed_query="SELECT Count (distinct UIN) confirm, round(Count(distinct UIN)* 100 /Cast( "&mswaccepted&" as float),2) percentconfirmed FROM Applicants where Confirmed = 'Y' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswconfirmed_query,conn 
      If rs.EOF Then
                      
                    Else

                    b= Replace("Confirmed","|",",")
                    c = Replace(rs("confirm"),"|",",")
                    d = Replace(rs("percentconfirmed"),"|",",")+("% of all accepted")
                    
                    pdf.Row b,c,d
                   
                 End If  
    Else

                  mswconfirmed_query="SELECT Count (distinct UIN) confirm FROM Applicants where Confirmed = 'Y' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswconfirmed_query,conn 
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
     If mswapplicants <> 0 Then 
					mswdenied_query="SELECT Count (distinct UIN) denied,round(Count(distinct UIN)* 100 / Cast("&mswapplicants&" as float),2) percentdenied FROM Applicants where Admission_decision = 'D' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswdenied_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Denials","|",",")
                    c = Replace(rs("denied"),"|",",")
                    d = Replace(rs("percentdenied"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
        mswdenied_query="SELECT Count (distinct UIN) denied FROM Applicants where Admission_decision = 'D' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswdenied_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Denials","|",",")
                    c = Replace(rs("denied"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswapplicants <> 0 Then
					mswwithdrawn_query="SELECT Count (distinct UIN) withdraw, round(Count (distinct UIN)*100/Cast("&mswapplicants&"as float), 2) percentwithdraw FROM Applicants where Withdrawn = 'Y' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswwithdrawn_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Withdrawals","|",",")
                    c = Replace(rs("withdraw"),"|",",")
                    d = Replace(rs("percentwithdraw"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
        mswwithdrawn_query="SELECT Count (distinct UIN) withdraw FROM Applicants where Withdrawn = 'Y' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswwithdrawn_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Withdrawals","|",",")
                    c = Replace(rs("withdraw"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If
 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswapplicants <> 0 Then
					mswnodecision_query="SELECT Count (distinct UIN) nodecision, round( Count (distinct UIN)*100/Cast("&mswapplicants&" as float),2) percentnodecision FROM Applicants where Admission_decision = '' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"' or Admission_decision is null and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswnodecision_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("No Decision","|",",")
                    c = Replace(rs("nodecision"),"|",",")
                    d = Replace(rs("percentnodecision"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
        mswnodecision_query="SELECT Count (distinct UIN) nodecision FROM Applicants where Admission_decision = '' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"' or Admission_decision is null and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswnodecision_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("No Decision","|",",")
                    c = Replace(rs("nodecision"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswapplicants <> 0 Then
					mswincomplete_query="SELECT Count (distinct UIN) incomplete, round(Count (distinct UIN) * 100/ Cast("&mswapplicants&" as float),2) percentincomplete FROM Applicants where Application_Status = 'IN-Incomplete' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswincomplete_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Incomplete","|",",")
                    c = Replace(rs("incomplete"),"|",",")
                    d = Replace(rs("percentincomplete"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
        mswincomplete_query="SELECT Count (distinct UIN) incomplete FROM Applicants where Application_Status = 'IN-Incomplete' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswincomplete_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Incomplete","|",",")
                    c = Replace(rs("incomplete"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswapplicants <> 0 Then
					mswdefer_query="SELECT Count (distinct UIN) deferment, round(Count (distinct UIN) * 100/Cast( "&mswapplicants&" as float),2) percentdefer FROM Applicants where Admission_decision = 'DF' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswdefer_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Deferment","|",",")
                    c = Replace(rs("deferment"),"|",",")
                    d = Replace(rs("percentdefer"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
            mswdefer_query="SELECT Count (distinct UIN) deferment FROM Applicants where Admission_decision = 'DF' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswdefer_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Deferment","|",",")
                    c = Replace(rs("deferment"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
        set rs=Server.CreateObject("ADODB.recordset")
    If mswapplicants <> 0 Then
					mswwaitlist_query="SELECT  Count (distinct UIN) waitlist, round(Count (distinct UIN) * 100/Cast ("&mswapplicants&" as float),2) percentwaitlist FROM Applicants where Admission_decision = 'W' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswwaitlist_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Wait List","|",",")
                    c = Replace(rs("waitlist"),"|",",")
                    d = Replace(rs("percentwaitlist"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
    mswwaitlist_query="SELECT  Count (distinct UIN) waitlist FROM Applicants where Admission_decision = 'W' and Degree_Program = 'MSW' and term_cd='"&AdmitTerm&"'"
					rs.Open mswwaitlist_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Wait List","|",",")
                    c = Replace(rs("waitlist"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close  
    pdf.Ln(10)
       ' //// Full Time ////
pdf.OrangeTitle("Full-Time")

mswFT_rows = 9
mswFT_cols = 3

Dim mswFT_col(3)

mswFT_col(1) = ""
mswFT_col(2) = "Total"
mswFT_col(3) = "Target"
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
					mswftapplicants_query="SELECT Count (distinct UIN) number FROM Applicants where Degree_Program='MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswftapplicants_query,conn 
    mswFTapplicants = rs("number")
                  
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
    If mswFTapplicants <> 0 Then
					mswftaccepted_query="SELECT Count (distinct UIN) status, round(Count(distinct UIN)* 100 /Cast("&mswFTapplicants&" as float), 2) percentaccepted FROM Applicants where Admission_decision IN ('A', 'S', 'ReAdmit') and Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswftaccepted_query,conn 
                  mswFTaccepted = rs("status")
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                   
                    b= Replace("Accepted","|",",")
                    c = Replace(rs("status"),"|",",")
                    d = Replace(rs("percentaccepted"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
    mswftaccepted_query="SELECT Count (distinct UIN) status FROM Applicants where Admission_decision IN ('A', 'S', 'ReAdmit') and Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswftaccepted_query,conn 
                  mswFTaccepted = rs("status")
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                   
                    b= Replace("Accepted","|",",")
                    c = Replace(rs("status"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
End If
 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
		
    If mswFTaccepted <> 0 Then 
   
								mswftconfirmed_query="SELECT Count (distinct UIN) confirm, round(Count(distinct UIN)* 100 /Cast( "&mswFTaccepted&" as float),2) percentconfirmed FROM Applicants where Confirmed = 'Y' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftconfirmed_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
        
                    b= Replace("Confirmed","|",",")
                    c = Replace(rs("confirm"),"|",",")
                    d = Replace(rs("percentconfirmed"),"|",",")+("% of all accepted")
                    
                    pdf.Row b,c,d
                   
                   End If
    Else

                  			mswftconfirmed_query="SELECT Count (distinct UIN) confirm FROM Applicants where Confirmed = 'Y' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftconfirmed_query,conn 
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
    If mswFTapplicants <> 0 Then
					mswftdenied_query="SELECT Count (distinct UIN) denied,round(Count(distinct UIN)* 100 / Cast("&mswFTapplicants&" as float),2) percentdenied FROM Applicants where Admission_decision = 'D' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftdenied_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Denials","|",",")
                    c = Replace(rs("denied"),"|",",")
                    d = Replace(rs("percentdenied"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
            mswftdenied_query="SELECT Count (distinct UIN) denied FROM Applicants where Admission_decision = 'D' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftdenied_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Denials","|",",")
                    c = Replace(rs("denied"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If


 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswFTapplicants <> 0 Then
					mswftwithdrawn_query="SELECT Count (distinct UIN) withdraw, round(Count (distinct UIN)*100/Cast("&mswFTapplicants&"as float), 2) percentwithdraw FROM Applicants where Withdrawn = 'Y' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftwithdrawn_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Withdrawals","|",",")
                    c = Replace(rs("withdraw"),"|",",")
                    d = Replace(rs("percentwithdraw"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
    mswftwithdrawn_query="SELECT Count (distinct UIN) withdraw FROM Applicants where Withdrawn = 'Y' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftwithdrawn_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Withdrawals","|",",")
                    c = Replace(rs("withdraw"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If
 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswFTapplicants <> 0 Then
					mswftnodecision_query="SELECT Count (distinct UIN) nodecision, round( Count (distinct UIN)*100/Cast("&mswFTapplicants&" as float),2) percentnodecision FROM Applicants where Admission_decision = '' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"' or Admission_decision is null and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftnodecision_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("No Decision","|",",")
                    c = Replace(rs("nodecision"),"|",",")
                    d = Replace(rs("percentnodecision"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
    mswftnodecision_query="SELECT Count (distinct UIN) nodecision FROM Applicants where Admission_decision = '' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"' or Admission_decision is null and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftnodecision_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("No Decision","|",",")
                    c = Replace(rs("nodecision"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswFTapplicants <> 0 Then
					mswftincomplete_query="SELECT Count (distinct UIN) incomplete, round(Count (distinct UIN) * 100/ Cast("&mswFTapplicants&" as float),2) percentincomplete FROM Applicants where Application_Status = 'IN-Incomplete' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftincomplete_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Incomplete","|",",")
                    c = Replace(rs("incomplete"),"|",",")
                    d = Replace(rs("percentincomplete"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
    mswftincomplete_query="SELECT Count (distinct UIN) incomplete FROM Applicants where Application_Status = 'IN-Incomplete' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftincomplete_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Incomplete","|",",")
                    c = Replace(rs("incomplete"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswFTapplicants <> 0 Then
					mswftdefer_query="SELECT Count (distinct UIN) deferment, round(Count (distinct UIN) * 100/Cast( "&mswFTapplicants&" as float),2) percentdefer FROM Applicants where Admission_decision = 'DF' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftdefer_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Deferment","|",",")
                    c = Replace(rs("deferment"),"|",",")
                    d = Replace(rs("percentdefer"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
    mswftdefer_query="SELECT Count (distinct UIN) deferment FROM Applicants where Admission_decision = 'DF' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftdefer_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Deferment","|",",")
                    c = Replace(rs("deferment"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
        set rs=Server.CreateObject("ADODB.recordset")
    If mswFTapplicants <> 0 Then
					mswftwaitlist_query="SELECT  Count (distinct UIN) waitlist, round(Count (distinct UIN) * 100/Cast ("&mswFTapplicants&" as float),2) percentwaitlist FROM Applicants where Admission_decision = 'W' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftwaitlist_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Wait List","|",",")
                    c = Replace(rs("waitlist"),"|",",")
                    d = Replace(rs("percentwaitlist"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
    mswftwaitlist_query="SELECT  Count (distinct UIN) waitlist FROM Applicants where Admission_decision = 'W' and Degree_Program = 'MSW' and Program_Type = 'FT' and term_cd='"&AdmitTerm&"'"
					rs.Open mswftwaitlist_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Wait List","|",",")
                    c = Replace(rs("waitlist"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close  

    pdf.Ln(13)
       ' //// PM Program ////
pdf.OrangeTitle("PM Program")

mswPM_rows = 9
mswPM_cols = 3

Dim mswPM_col(3)

mswPM_col(1) = ""
mswPM_col(2) = "Total"
mswPM_col(3) = "Target"
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
					mswpmapplicants_query="SELECT Count (distinct UIN) number FROM Applicants where Degree_Program='MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswpmapplicants_query,conn 
    mswPMapplicants = rs("number")
                  
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
    If mswPMapplicants <> 0 Then
					mswpmaccepted_query="SELECT Count (distinct UIN) status, round(Count(distinct UIN)* 100 /Cast("&mswPMapplicants&" as float), 2) percentaccepted FROM Applicants where Admission_decision IN ('A', 'S', 'ReAdmit') and Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswpmaccepted_query,conn 
                  mswPMaccepted = rs("status")
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                   
                    b= Replace("Accepted","|",",")
                    c = Replace(rs("status"),"|",",")
                    d = Replace(rs("percentaccepted"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
            mswpmaccepted_query="SELECT Count (distinct UIN) status FROM Applicants where Admission_decision IN ('A', 'S', 'ReAdmit') and Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswpmaccepted_query,conn 
                  mswPMaccepted = rs("status")
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                   
                    b= Replace("Accepted","|",",")
                    c = Replace(rs("status"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
				
     If mswPMaccepted <> 0 Then 
   
					mswpmconfirmed_query="SELECT Count (distinct UIN) confirm, round(Count(distinct UIN)* 100 /Cast( "&mswPMaccepted&" as float),2) percentconfirmed FROM Applicants where Confirmed = 'Y' and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmconfirmed_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
        
                    b= Replace("Confirmed","|",",")
                    c = Replace(rs("confirm"),"|",",")
                    d = Replace(rs("percentconfirmed"),"|",",")+("% of all accepted")
                    
                    pdf.Row b,c,d
                   
                   End If
    Else

                  mswpmconfirmed_query="SELECT Count (distinct UIN) confirm FROM Applicants where Confirmed = 'Y' and Degree_Program = 'MSW'and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmconfirmed_query,conn 
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
    If mswPMapplicants <> 0 Then
					mswpmdenied_query="SELECT Count (distinct UIN) denied,round(Count(distinct UIN)* 100 / Cast("&mswPMapplicants&" as float),2) percentdenied FROM Applicants where Admission_decision = 'D' and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmdenied_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Denials","|",",")
                    c = Replace(rs("denied"),"|",",")
                    d = Replace(rs("percentdenied"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
            mswpmdenied_query="SELECT Count (distinct UIN) denied FROM Applicants where Admission_decision = 'D' and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmdenied_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Denials","|",",")
                    c = Replace(rs("denied"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswPMapplicants <> 0 Then
					mswpmwithdrawn_query="SELECT Count (distinct UIN) withdraw, round(Count (distinct UIN)*100/Cast("&mswPMapplicants&"as float), 2) percentwithdraw FROM Applicants where Withdrawn = 'Y' and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmwithdrawn_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Withdrawals","|",",")
                    c = Replace(rs("withdraw"),"|",",")
                    d = Replace(rs("percentwithdraw"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
            mswpmwithdrawn_query="SELECT Count (distinct UIN) withdraw FROM Applicants where Withdrawn = 'Y' and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmwithdrawn_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Withdrawals","|",",")
                    c = Replace(rs("withdraw"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If 

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswPMapplicants <> 0 Then
					mswpmnodecision_query="SELECT Count (distinct UIN) nodecision, round( Count (distinct UIN)*100/Cast("&mswPMapplicants&" as float),2) percentnodecision FROM Applicants where Admission_decision = '' and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"' or Admission_decision is null and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmnodecision_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("No Decision","|",",")
                    c = Replace(rs("nodecision"),"|",",")
                    d = Replace(rs("percentnodecision"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
            mswpmnodecision_query="SELECT Count (distinct UIN) nodecision FROM Applicants where Admission_decision = '' and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"' or Admission_decision is null and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmnodecision_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("No Decision","|",",")
                    c = Replace(rs("nodecision"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswPMapplicants <> 0 Then
					mswpmincomplete_query="SELECT Count (distinct UIN) incomplete, round(Count (distinct UIN) * 100/ Cast("&mswPMapplicants&" as float),2) percentincomplete FROM Applicants where Application_Status = 'IN-Incomplete' and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmincomplete_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Incomplete","|",",")
                    c = Replace(rs("incomplete"),"|",",")
                    d = Replace(rs("percentincomplete"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
                mswpmincomplete_query="SELECT Count (distinct UIN) incomplete FROM Applicants where Application_Status = 'IN-Incomplete' and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmincomplete_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Incomplete","|",",")
                    c = Replace(rs("incomplete"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswPMapplicants <> 0 Then
					mswpmdefer_query="SELECT Count (distinct UIN) deferment, round(Count (distinct UIN) * 100/Cast( "&mswPMapplicants&" as float),2) percentdefer FROM Applicants where Admission_decision = 'DF' and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmdefer_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Deferment","|",",")
                    c = Replace(rs("deferment"),"|",",")
                    d = Replace(rs("percentdefer"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
            mswpmdefer_query="SELECT Count (distinct UIN) deferment FROM Applicants where Admission_decision = 'DF' and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmdefer_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Deferment","|",",")
                    c = Replace(rs("deferment"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
        set rs=Server.CreateObject("ADODB.recordset")
    If mswPMapplicants <> 0 Then
					mswpmwaitlist_query="SELECT  Count (distinct UIN) waitlist, round(Count (distinct UIN) * 100/Cast ("&mswPMapplicants&" as float),2) percentwaitlist FROM Applicants where Admission_decision = 'W' and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmwaitlist_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Wait List","|",",")
                    c = Replace(rs("waitlist"),"|",",")
                    d = Replace(rs("percentwaitlist"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
            mswpmwaitlist_query="SELECT  Count (distinct UIN) waitlist FROM Applicants where Admission_decision = 'W' and Degree_Program = 'MSW' and Program_Type = 'PM' and term_cd='"&AdmitTerm&"'"
					rs.Open mswpmwaitlist_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Wait List","|",",")
                    c = Replace(rs("waitlist"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If 

 rs.close  

pdf.Ln(25)
           ' //// ADV Program ////
pdf.OrangeTitle("Advanced Standing")

mswADV_rows = 9
mswADV_cols = 3

Dim mswADV_col(3)

mswADV_col(1) = ""
mswADV_col(2) = "Total"
mswADV_col(3) = "Target"
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
					mswadvapplicants_query="SELECT Count (distinct UIN) number FROM Applicants where Degree_Program='MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswadvapplicants_query,conn 
    mswADVapplicants = rs("number")
                  
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
    If mswADVapplicants <> 0 Then
					mswadvaccepted_query="SELECT Count (distinct UIN) status, round(Count(distinct UIN)* 100 /Cast("&mswADVapplicants&" as float), 2) percentaccepted FROM Applicants where Admission_decision IN ('A', 'S', 'ReAdmit') and Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswadvaccepted_query,conn 
                    mswADVaccepted = rs("status")
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                   
                    b= Replace("Accepted","|",",")
                    c = Replace(rs("status"),"|",",")
                    d = Replace(rs("percentaccepted"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
        mswadvaccepted_query="SELECT Count (distinct UIN) status FROM Applicants where Admission_decision IN ('A', 'S', 'ReAdmit') and Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'"
					rs.Open mswadvaccepted_query,conn 
                  mswADVaccepted = rs("status")
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                   
                    b= Replace("Accepted","|",",")
                    c = Replace(rs("status"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
				
     If mswADVaccepted <> 0 Then 
   
					mswadvconfirmed_query="SELECT Count (distinct UIN) confirm, round(Count(distinct UIN)* 100 /Cast( "&mswADVaccepted&" as float),2) percentconfirmed FROM Applicants where Confirmed = 'Y' and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvconfirmed_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
        
                    b= Replace("Confirmed","|",",")
                    c = Replace(rs("confirm"),"|",",")
                    d = Replace(rs("percentconfirmed"),"|",",")+("% of all accepted")
                    
                    pdf.Row b,c,d
                   
                   End If
    Else

                  mswadvconfirmed_query="SELECT Count (distinct UIN) confirm FROM Applicants where Confirmed = 'Y' and Degree_Program = 'MSW'and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvconfirmed_query,conn 
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
    If mswADVapplicants <> 0 Then 
					mswadvdenied_query="SELECT Count (distinct UIN) denied,round(Count(distinct UIN)* 100 / Cast("&mswADVapplicants&" as float),2) percentdenied FROM Applicants where Admission_decision = 'D' and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvdenied_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Denials","|",",")
                    c = Replace(rs("denied"),"|",",")
                    d = Replace(rs("percentdenied"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
            mswadvdenied_query="SELECT Count (distinct UIN) denied FROM Applicants where Admission_decision = 'D' and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvdenied_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Denials","|",",")
                    c = Replace(rs("denied"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If


 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswADVapplicants <> 0 Then 
					mswadvwithdrawn_query="SELECT Count (distinct UIN) withdraw, round(Count (distinct UIN)*100/Cast("&mswADVapplicants&"as float), 2) percentwithdraw FROM Applicants where Withdrawn = 'Y' and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvwithdrawn_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Withdrawals","|",",")
                    c = Replace(rs("withdraw"),"|",",")
                    d = Replace(rs("percentwithdraw"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
    mswadvwithdrawn_query="SELECT Count (distinct UIN) withdraw FROM Applicants where Withdrawn = 'Y' and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvwithdrawn_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Withdrawals","|",",")
                    c = Replace(rs("withdraw"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If 

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswADVapplicants <> 0 Then 
					mswadvnodecision_query="SELECT Count (distinct UIN) nodecision, round( Count (distinct UIN)*100/Cast("&mswADVapplicants&" as float),2) percentnodecision FROM Applicants where Admission_decision = '' and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"' or Admission_decision is null and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvnodecision_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("No Decision","|",",")
                    c = Replace(rs("nodecision"),"|",",")
                    d = Replace(rs("percentnodecision"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
    mswadvnodecision_query="SELECT Count (distinct UIN) nodecision FROM Applicants where Admission_decision = '' and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"' or Admission_decision is null and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvnodecision_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("No Decision","|",",")
                    c = Replace(rs("nodecision"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
    If mswADVapplicants <> 0 Then
					mswadvincomplete_query="SELECT Count (distinct UIN) incomplete, round(Count (distinct UIN) * 100/ Cast("&mswADVapplicants&" as float),2) percentincomplete FROM Applicants where Application_Status = 'IN-Incomplete' and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvincomplete_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Incomplete","|",",")
                    c = Replace(rs("incomplete"),"|",",")
                    d = Replace(rs("percentincomplete"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
            mswadvincomplete_query="SELECT Count (distinct UIN) incomplete FROM Applicants where Application_Status = 'IN-Incomplete' and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvincomplete_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Incomplete","|",",")
                    c = Replace(rs("incomplete"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
     If mswADVapplicants <> 0 Then
					mswadvdefer_query="SELECT Count (distinct UIN) deferment, round(Count (distinct UIN) * 100/Cast( "&mswADVapplicants&" as float),2) percentdefer FROM Applicants where Admission_decision = 'DF' and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvdefer_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Deferment","|",",")
                    c = Replace(rs("deferment"),"|",",")
                    d = Replace(rs("percentdefer"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
            mswadvdefer_query="SELECT Count (distinct UIN) deferment FROM Applicants where Admission_decision = 'DF' and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvdefer_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Deferment","|",",")
                    c = Replace(rs("deferment"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close
        set rs=Server.CreateObject("ADODB.recordset")
    If mswADVapplicants <> 0 Then
					mswadvwaitlist_query="SELECT  Count (distinct UIN) waitlist, round(Count (distinct UIN) * 100/Cast ("&mswADVapplicants&" as float),2) percentwaitlist FROM Applicants where Admission_decision = 'W' and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvwaitlist_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Wait List","|",",")
                    c = Replace(rs("waitlist"),"|",",")
                    d = Replace(rs("percentwaitlist"),"|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    Else
    mswadvwaitlist_query="SELECT  Count (distinct UIN) waitlist FROM Applicants where Admission_decision = 'W' and Degree_Program = 'MSW' and Program_Type = 'ADV' and term_cd='"&AdmitTerm&"'"
					rs.Open mswadvwaitlist_query,conn 
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
                   
              
                   
                    
                    b= Replace("Wait List","|",",")
                    c = Replace(rs("waitlist"),"|",",")
                    d = Replace("0","|",",")+("% of all applicants")
                    
                    pdf.Row b,c,d
                   
                    End If
    End If

 rs.close  



    pdf.Ln(10)
pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close

%>
