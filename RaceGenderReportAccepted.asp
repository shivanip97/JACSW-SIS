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

pdf.ChapterTitle2("         Report 8 -  Race/Gender Ethinicity Report - Accepted - "&Termsel& "      "  &LastUpdatedDt&" "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(10)

rs.close


   ' //// Total MSW ////

    '///AfericanAmerican///

pdf.OrangeTitle("Total MSW")

totalmsw_rows = 9
totalmsw_cols = 9

Dim totalmsw_col(9)

totalmsw_col(1) = "Race"
totalmsw_col(2) = "Total Female"
totalmsw_col(3) = "% Female"
totalmsw_col(4) = "Total Male"
totalmsw_col(5) = "% Male"
totalmsw_col(6) = "Total Not Aailable"
totalmsw_col(7) = "% Not Available"
totalmsw_col(8) = "Total"
totalmsw_col(9) = "Percent"

pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 30,25,20,20,15,25,25,15,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Race","TotalFemale", "%Female", "Total Male", "%Male", "Total Not Aailable", "% Not Available", "Total", "Percent"
pdf.SetFont "Arial","",10

    set rs=Server.CreateObject("ADODB.recordset")
					mswTotalFemales_query="SELECT Count (distinct UIN) numberfemales FROM Applicants where Degree_Program='MSW' and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F'"
					rs.Open mswTotalFemales_query,conn 
    mswtotalfemales = rs("numberfemales")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
					mswTotalMales_query="SELECT Count (distinct UIN) numbermales FROM Applicants where Degree_Program='MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M'"
					rs.Open mswTotalMales_query,conn 
    mswtotalmales = rs("numbermales")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
					mswTotalnotavailable_query="SELECT Count (distinct UIN) numberunavailable FROM Applicants where Degree_Program='MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M')"
					rs.Open mswTotalnotavailable_query,conn 
    mswtotalunavailable = rs("numberunavailable")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
mswTotal = mswtotalfemales + mswtotalmales + mswtotalunavailable

                    set rs=Server.CreateObject("ADODB.recordset")
    If mswtotalfemales <> 0 Then
					mswAfricanAmericanfemale_query="SELECT Count (distinct UIN) africanamericancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswtotalfemales&" as float), 2) percentafricanamericanfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Black/African American'"
					rs.Open mswAfricanAmericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("African/American","|",",")
                    b = Replace(rs("africanamericancountfemale"),"|",",")
                    c = Replace(rs("percentafricanamericanfemales"),"|",",")
    End If
    Else
    mswAfricanAmericanfemale_query="SELECT Count (distinct UIN) africanamericancountfemale FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Black/African American'"
					rs.Open mswAfricanAmericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("African/American","|",",")
                    b = Replace(rs("africanamericancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close
    If mswtotalmales <> 0 Then
    mswAfricanAmericanmale_query = "SELECT Count (distinct UIN) africanamericancountmales, round(Count(distinct UIN)* 100 /Cast("&mswtotalmales&" as float), 2) percentafricanamericanmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Black/African American'"
    rs.Open mswAfricanAmericanmale_query,conn 
    
                    d= Replace(rs("africanamericancountmales"),"|",",")
                    e = Replace(rs("percentafricanamericanmales"),"|",",")
    Else
    mswAfricanAmericanmale_query = "SELECT Count (distinct UIN) africanamericancountmales FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswAfricanAmericanmale_query,conn 
    
                    d= Replace(rs("africanamericancountmales"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close

     If mswtotalunavailable <> 0 Then
    mswAfricanAmericanna_query = "SELECT Count (distinct UIN) africanamericancountna, round(Count(distinct UIN)* 100 /Cast("&mswtotalunavailable&" as float), 2) percentafricanamericanna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M') and Race_ethinicity = 'Black/African American'"
    rs.Open mswAfricanAmericanna_query,conn 
    
                    f= Replace(rs("africanamericancountna"),"|",",")
                    g = Replace(rs("percentafricanamericanna"),"|",",")
    Else
    mswAfricanAmericanna_query = "SELECT Count (distinct UIN) africanamericancountna FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswAfricanAmericanna_query,conn 
    
                    f= Replace(rs("africanamericancountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswTotal <> 0 Then
    mswAfricanAmerican_query = "SELECT Count (distinct UIN) africanamericancount, round(Count(distinct UIN)* 100 /Cast("&mswTotal&" as float), 2) percentafricanamerican FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswAfricanAmerican_query,conn 
    
                    h = Replace(rs("africanamericancount") ,"|",",")
                    i = Replace(rs("percentafricanamerican"),"|",",")
    Else
    mswAfricanAmerican_query = "SELECT Count (distinct UIN) africanamericancount FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswAfricanAmerican_query,conn 
    
                    h = Replace(rs("africanamericancount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                 

    '//Hispanic//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswtotalfemales <> 0 Then
					mswHispanicfemale_query="SELECT Count (distinct UIN) hispaniccountfemale, round(Count(distinct UIN)* 100 /Cast("&mswtotalfemales&" as float), 2) percenthispanicfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Hispanic'"
					rs.Open mswHispanicfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Hispanic","|",",")
                    b = Replace(rs("hispaniccountfemale"),"|",",")
                    c = Replace(rs("percenthispanicfemales"),"|",",")
    End If
    Else
    mswHispanicfemale_query="SELECT Count (distinct UIN) hispaniccountfemale FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' and Race_ethinicity = 'Hispanic'"
					rs.Open mswHispanicfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Hispanic","|",",")
                    b = Replace(rs("hispaniccountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswtotalmales <> 0 Then
    mswHispanicmale_query = "SELECT Count (distinct UIN) hispaniccountmales, round(Count(distinct UIN)* 100 /Cast("&mswtotalmales&" as float), 2) percenthispanicmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Hispanic'"
    rs.Open mswHispanicmale_query,conn 
    
                    d= Replace(rs("hispaniccountmales"),"|",",")
                    e = Replace(rs("percenthispanicmales"),"|",",")
    Else 
    mswHispanicmale_query = "SELECT Count (distinct UIN) hispaniccountmales FROM Applicants where Degree_Program = 'MSW'  and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswHispanicmale_query,conn 
    
                    d= Replace(rs("hispaniccountmales"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close

     If mswtotalunavailable <> 0 Then
    mswHispanicna_query = "SELECT Count (distinct UIN) hispaniccountna, round(Count(distinct UIN)* 100 /Cast("&mswtotalunavailable&" as float), 2) percenthispanicna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M') and Race_ethinicity = 'Hispanic'"
    rs.Open mswHispanicna_query,conn 
    
                    f= Replace(rs("hispaniccountna"),"|",",")
                    g = Replace(rs("percenthispanicna"),"|",",")
    Else
    mswHispanicna_query = "SELECT Count (distinct UIN) hispaniccountna FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswHispanicna_query,conn 
    
                    f= Replace(rs("hispaniccountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswTotal <> 0 Then
    mswHispanic_query = "SELECT Count (distinct UIN) hispaniccount, round(Count(distinct UIN)* 100 /Cast("&mswTotal&" as float), 2) percenthispanic FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Hispanic'"
    rs.Open mswHispanic_query,conn 
    
                    h = Replace(rs("hispaniccount") ,"|",",")
                    i = Replace(rs("percenthispanic"),"|",",")
    Else
     mswHispanic_query = "SELECT Count (distinct UIN) hispaniccount FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswHispanic_query,conn 
    
                    h = Replace(rs("hispaniccount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                 
     '//Asian//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswtotalfemales <> 0 Then
					mswAsianfemale_query="SELECT Count (distinct UIN) asiancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswtotalfemales&" as float), 2) percentasianfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Asian'"
					rs.Open mswAsianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Asian","|",",")
                    b = Replace(rs("asiancountfemale"),"|",",")
                    c = Replace(rs("percentasianfemales"),"|",",")
    End If
    Else
    mswAsianfemale_query="SELECT Count (distinct UIN) asiancountfemale FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
					rs.Open mswAsianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Asian","|",",")
                    b = Replace(rs("asiancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswtotalmales <> 0 Then
    mswAsianmale_query = "SELECT Count (distinct UIN) asiancountmales, round(Count(distinct UIN)* 100 /Cast("&mswtotalmales&" as float), 2) percentasianmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Asian'"
    rs.Open mswAsianmale_query,conn 
    
                    d= Replace(rs("asiancountmales"),"|",",")
                    e = Replace(rs("percentasianmales"),"|",",")
    Else
    mswAsianmale_query = "SELECT Count (distinct UIN) asiancountmales FROM Applicants where Degree_Program = 'MSW'  and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswAsianmale_query,conn 
    
                    d= Replace(rs("asiancountmales"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
     If mswtotalunavailable <> 0 Then
    mswAsianna_query = "SELECT Count (distinct UIN) asiancountna, round(Count(distinct UIN)* 100 /Cast("&mswtotalunavailable&" as float), 2) percentasianna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M') and Race_ethinicity = 'Asian'"
    rs.Open mswAsianna_query,conn 
    
                    f= Replace(rs("asiancountna"),"|",",")
                    g = Replace(rs("percentasianna"),"|",",")
    Else
    mswAsianna_query = "SELECT Count (distinct UIN) asiancountna FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswAsianna_query,conn 
    
                    f= Replace(rs("asiancountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswTotal <> 0 Then
    mswAsian_query = "SELECT Count (distinct UIN) asiancount, round(Count(distinct UIN)* 100 /Cast("&mswTotal&" as float), 2) percentasian FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Asian'"
    rs.Open mswAsian_query,conn 
    
                    h = Replace(rs("asiancount") ,"|",",")
                    i = Replace(rs("percentasian"),"|",",")
    Else
    mswAsian_query = "SELECT Count (distinct UIN) asiancount FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswAsian_query,conn 
    
                    h = Replace(rs("asiancount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                 
     
    '//Native American//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswtotalfemales <> 0 Then
					mswNativeamericanfemale_query="SELECT Count (distinct UIN) nativeamericancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswtotalfemales&" as float), 2) percentnativeamericanfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and NHPI_Race_Ind = 'Y'"
					rs.Open mswNativeamericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Native American","|",",")
                    b = Replace(rs("nativeamericancountfemale"),"|",",")
                    c = Replace(rs("percentnativeamericanfemales"),"|",",")
    End If
    Else
    mswNativeamericanfemale_query="SELECT Count (distinct UIN) nativeamericancountfemale FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
					rs.Open mswNativeamericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Native American","|",",")
                    b = Replace(rs("nativeamericancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close
    
    If mswtotalmales <> 0 Then
    mswNativeamericanmale_query = "SELECT Count (distinct UIN) nativeamericancountmale, round(Count(distinct UIN)* 100 /Cast("&mswtotalmales&" as float), 2) percentnativeamericanmales FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'M' and NHPI_Race_Ind = 'Y'"
    rs.Open mswNativeamericanmale_query,conn 
    
                    d= Replace(rs("nativeamericancountmale"),"|",",")
                    e = Replace(rs("percentnativeamericanmales"),"|",",")
    Else
    mswNativeamericanmale_query = "SELECT Count (distinct UIN) nativeamericancountmale FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswNativeamericanmale_query,conn 
    
                    d= Replace(rs("nativeamericancountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
     If mswtotalunavailable <> 0 Then
    mswNativeamericanna_query = "SELECT Count (distinct UIN) nativeamericancountna, round(Count(distinct UIN)* 100 /Cast("&mswtotalunavailable&" as float), 2) percentnativeamericanna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M') and NHPI_Race_Ind = 'Y'"
    rs.Open mswNativeamericanna_query,conn 
    
                    f= Replace(rs("nativeamericancountna"),"|",",")
                    g = Replace(rs("percentnativeamericanna"),"|",",")
    Else
    mswNativeamericanna_query = "SELECT Count (distinct UIN) nativeamericancountna FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswNativeamericanna_query,conn 
    
                    f= Replace(rs("nativeamericancountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswTotal <> 0 Then
    mswNativeAmerican_query = "SELECT Count (distinct UIN) nativeamericancount, round(Count(distinct UIN)* 100 /Cast("&mswTotal&" as float), 2) percentnativeamerican FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and NHPI_Race_Ind = 'Y'"
    rs.Open mswNativeAmerican_query,conn 
    
                    h = Replace(rs("nativeamericancount") ,"|",",")
                    i = Replace(rs("percentnativeamerican"),"|",",")
    Else
     mswNativeAmerican_query = "SELECT Count (distinct UIN) nativeamericancount FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswNativeAmerican_query,conn 
    
                    h = Replace(rs("nativeamericancount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
 
       
    '//International//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswtotalfemales <> 0 Then
					mswinternationalfemale_query="SELECT Count (distinct UIN) internationalcountfemale, round(Count(distinct UIN)* 100 /Cast("&mswtotalfemales&" as float), 2) percentinternationalfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'International'"
					rs.Open mswinternationalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("International","|",",")
                    b = Replace(rs("internationalcountfemale"),"|",",")
                    c = Replace(rs("percentinternationalfemales"),"|",",")
    End If
    Else
    mswinternationalfemale_query="SELECT Count (distinct UIN) internationalcountfemale FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
					rs.Open mswinternationalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("International","|",",")
                    b = Replace(rs("internationalcountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If 
    End If
     rs.close

    If mswtotalmales <> 0 Then
    mswinternationalmale_query = "SELECT Count (distinct UIN) internationalcountmale, round(Count(distinct UIN)* 100 /Cast("&mswtotalmales&" as float), 2) percentinternationalmales FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'M' and Race_ethinicity = 'International'"
    rs.Open mswinternationalmale_query,conn 
    
                    d= Replace(rs("internationalcountmale"),"|",",")
                    e = Replace(rs("percentinternationalmales"),"|",",")
    Else
     mswinternationalmale_query = "SELECT Count (distinct UIN) internationalcountmale FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswinternationalmale_query,conn 
    
                    d= Replace(rs("internationalcountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
     If mswtotalunavailable <> 0 Then
    mswinternationalna_query = "SELECT Count (distinct UIN) internationalcountna, round(Count(distinct UIN)* 100 /Cast("&mswtotalunavailable&" as float), 2) percentinternationalna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M') and Race_ethinicity = 'International'"
    rs.Open mswinternationalna_query,conn 
    
                    f= Replace(rs("internationalcountna"),"|",",")
                    g = Replace(rs("percentinternationalna"),"|",",")
    Else
    mswinternationalna_query = "SELECT Count (distinct UIN) internationalcountna FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswinternationalna_query,conn 
    
                    f= Replace(rs("internationalcountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswTotal <> 0 Then 
    mswinternational_query = "SELECT Count (distinct UIN) internationalcount, round(Count(distinct UIN)* 100 /Cast("&mswTotal&" as float), 2) percentinternational FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit')  and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'International'"
    rs.Open mswinternational_query,conn 
    
                    h = Replace(rs("internationalcount") ,"|",",")
                    i = Replace(rs("percentinternational"),"|",",")
    Else 
     mswinternational_query = "SELECT Count (distinct UIN) internationalcount FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswinternational_query,conn 
    
                    h = Replace(rs("internationalcount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
               

     '//Multi-Race//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswtotalfemales <> 0 Then
					mswmultiracefemale_query="SELECT Count (distinct UIN) multiracecountfemale, round(Count(distinct UIN)* 100 /Cast("&mswtotalfemales&" as float), 2) percentmultiracefemales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Multi-Race'"
					rs.Open mswmultiracefemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Multi-Race","|",",")
                    b = Replace(rs("multiracecountfemale"),"|",",")
                    c = Replace(rs("percentmultiracefemales"),"|",",")
    End If
    Else 
    mswmultiracefemale_query="SELECT Count (distinct UIN) multiracecountfemale FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
					rs.Open mswmultiracefemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Multi-Race","|",",")
                    b = Replace(rs("multiracecountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswtotalmales <> 0 Then
    mswmultiracemale_query = "SELECT Count (distinct UIN) multiracecountmale, round(Count(distinct UIN)* 100 /Cast("&mswtotalmales&" as float), 2) percentmultiracemales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Multi-Race'"
    rs.Open mswmultiracemale_query,conn 
    
                    d= Replace(rs("multiracecountmale"),"|",",")
                    e = Replace(rs("percentmultiracemales"),"|",",")
    Else
    mswmultiracemale_query = "SELECT Count (distinct UIN) multiracecountmale FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswmultiracemale_query,conn 
    
                    d= Replace(rs("multiracecountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
     
     If mswtotalunavailable <> 0 Then
    mswmultiracena_query = "SELECT Count (distinct UIN) multiracecountna, round(Count(distinct UIN)* 100 /Cast("&mswtotalunavailable&" as float), 2) percentmultiracena FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswmultiracena_query,conn 
    
                    f= Replace(rs("multiracecountna"),"|",",")
                    g = Replace(rs("percentmultiracena"),"|",",")
    Else
    mswmultiracena_query = "SELECT Count (distinct UIN) multiracecountna FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswmultiracena_query,conn 
    
                    f= Replace(rs("multiracecountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswTotal <> 0 Then
    mswmultirace_query = "SELECT Count (distinct UIN) multiracecount, round(Count(distinct UIN)* 100 /Cast("&mswTotal&" as float), 2) percentmultirace FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Multi-Race'"
    rs.Open mswmultirace_query,conn 
    
                    h = Replace(rs("multiracecount") ,"|",",")
                    i = Replace(rs("percentmultirace"),"|",",")
    Else
    mswmultirace_query = "SELECT Count (distinct UIN) multiracecount FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswmultirace_query,conn 
    
                    h = Replace(rs("multiracecount") ,"|",",")
                    i = Replace("0","|",",")
     End If 
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                  
    
    '//Total Minority//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswtotalfemales <> 0 Then
					mswminorityfemale_query="SELECT Count (distinct UIN) minoritycountfemale, round(Count(distinct UIN)* 100 /Cast("&mswtotalfemales&" as float), 2) percentminorityfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity != 'White'"
					rs.Open mswminorityfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total Minority","|",",")
                    b = Replace(rs("minoritycountfemale"),"|",",")
                    c = Replace(rs("percentminorityfemales"),"|",",")
    End If
    Else
    mswminorityfemale_query="SELECT Count (distinct UIN) minoritycountfemale FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' and Race_ethinicity != 'White'"
					rs.Open mswminorityfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total Minority","|",",")
                    b = Replace(rs("minoritycountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If 
    End If
     rs.close

    If mswtotalmales <> 0 Then
    mswminoritymale_query = "SELECT Count (distinct UIN) minoritycountmale, round(Count(distinct UIN)* 100 /Cast("&mswtotalmales&" as float), 2) percentminoritymales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit')  and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity != 'White'"
    rs.Open mswminoritymale_query,conn 
    
                    d= Replace(rs("minoritycountmale"),"|",",")
                    e = Replace(rs("percentminoritymales"),"|",",")
    Else
     mswminoritymale_query = "SELECT Count (distinct UIN) minoritycountmale FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswminoritymale_query,conn 
    
                    d= Replace(rs("minoritycountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
 
     If mswtotalunavailable <> 0 Then
    mswminorityna_query = "SELECT Count (distinct UIN) minoritycountna, round(Count(distinct UIN)* 100 /Cast("&mswtotalunavailable&" as float), 2) percentminorityna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M') and Race_ethinicity != 'White'"
    rs.Open mswminorityna_query,conn 
    
                    f= Replace(rs("minoritycountna"),"|",",")
                    g = Replace(rs("percentminorityna"),"|",",")
    Else
    mswminorityna_query = "SELECT Count (distinct UIN) minoritycountna FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswminorityna_query,conn 
    
                    f= Replace(rs("minoritycountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswTotal <> 0 Then 
    mswMinority_query = "SELECT Count (distinct UIN) minoritycount, round(Count(distinct UIN)* 100 /Cast("&mswTotal&" as float), 2) percentminority FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity != 'White'"
    rs.Open mswMinority_query,conn 
    
                    h = Replace(rs("minoritycount") ,"|",",")
                    i = Replace(rs("percentminority"),"|",",")
    Else 
     mswMinority_query = "SELECT Count (distinct UIN) minoritycount, round(Count(distinct UIN)* 100 /Cast("&mswTotal&" as float), 2) percentminority FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity != 'White'"
    rs.Open mswMinority_query,conn 
    
                    h = Replace(rs("minoritycount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
               

     '//Caucasian//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswtotalfemales <> 0 Then
					mswcaucasianfemale_query="SELECT Count (distinct UIN) caucasiancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswtotalfemales&" as float), 2) percentcaucasianfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit')  and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'White'"
					rs.Open mswcaucasianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Caucasian","|",",")
                    b = Replace(rs("caucasiancountfemale"),"|",",")
                    c = Replace(rs("percentcaucasianfemales"),"|",",")
    End If
    Else 
    mswcaucasianfemale_query="SELECT Count (distinct UIN) caucasiancountfemale FROM Applicants where Degree_Program = 'MSW'  and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
					rs.Open mswcaucasianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Caucasian","|",",")
                    b = Replace(rs("caucasiancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswtotalmales <> 0 Then
    mswcaucasianmale_query = "SELECT Count (distinct UIN) caucasiancountmale, round(Count(distinct UIN)* 100 /Cast("&mswtotalmales&" as float), 2) percentcaucasianmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'White'"
    rs.Open mswcaucasianmale_query,conn 
    
                    d= Replace(rs("caucasiancountmale"),"|",",")
                    e = Replace(rs("percentcaucasianmales"),"|",",")
    Else
    mswcaucasianmale_query = "SELECT Count (distinct UIN) caucasiancountmale FROM Applicants where Degree_Program = 'MSW'  and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswcaucasianmale_query,conn 
    
                    d= Replace(rs("caucasiancountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
 
     If mswtotalunavailable <> 0 Then
    mswcaucasianna_query = "SELECT Count (distinct UIN) caucasiancountna, round(Count(distinct UIN)* 100 /Cast("&mswtotalunavailable&" as float), 2) percentcaucasianna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M') and Race_ethinicity = 'White'"
    rs.Open mswcaucasianna_query,conn 
    
                    f= Replace(rs("caucasiancountna"),"|",",")
                    g = Replace(rs("percentcaucasianna"),"|",",")
    Else
    mswcaucasianna_query = "SELECT Count (distinct UIN) caucasiancountna FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswcaucasianna_query,conn 
    
                    f= Replace(rs("caucasiancountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswTotal <> 0 Then
    mswCaucasian_query = "SELECT Count (distinct UIN) caucasiancount, round(Count(distinct UIN)* 100 /Cast("&mswTotal&" as float), 2) percentcaucasian FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'White'"
    rs.Open mswCaucasian_query,conn 
    
                    h = Replace(rs("caucasiancount") ,"|",",")
                    i = Replace(rs("percentcaucasian"),"|",",")
    Else
    mswCaucasian_query = "SELECT Count (distinct UIN) caucasiancount FROM Applicants where Degree_Program = 'MSW'  and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswCaucasian_query,conn 
    
                    h = Replace(rs("caucasiancount") ,"|",",")
                    i = Replace("0","|",",")
     End If 
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                


    '//Unknown//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswtotalfemales <> 0 Then
					mswUnknownfemale_query="SELECT Count (distinct UIN) unknowncountfemale, round(Count(distinct UIN)* 100 /Cast("&mswtotalfemales&" as float), 2) percentunknownfemales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Unknown'"
					rs.Open mswUnknownfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Undeclared","|",",")
                    b = Replace(rs("unknowncountfemale"),"|",",")
                    c = Replace(rs("percentunknownfemales"),"|",",")
    End If
    Else
    mswUnknownfemale_query="SELECT Count (distinct UIN) unknowncountfemale FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
					rs.Open mswUnknownfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Undeclared","|",",")
                    b = Replace(rs("unknowncountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswtotalmales <> 0 Then
    mswUnknownmale_query = "SELECT Count (distinct UIN) unknowncountmale, round(Count(distinct UIN)* 100 /Cast("&mswtotalmales&" as float), 2) percentunknownmales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Unknown'"
    rs.Open mswUnknownmale_query,conn 
    
                    d= Replace(rs("unknowncountmale"),"|",",")
                    e = Replace(rs("percentunknownmales"),"|",",")
    Else
    mswUnknownmale_query = "SELECT Count (distinct UIN) unknowncountmale FROM Applicants where Degree_Program = 'MSW'  and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswUnknownmale_query,conn 
    
                    d= Replace(rs("unknowncountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
     If mswtotalunavailable <> 0 Then
    mswUnknownna_query = "SELECT Count (distinct UIN) unknowncountna, round(Count(distinct UIN)* 100 /Cast("&mswtotalunavailable&" as float), 2) percentunknownna FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender not in ('F', 'M') and Race_ethinicity = 'Unknown'"
    rs.Open mswUnknownna_query,conn 
    
                    f= Replace(rs("unknowncountna"),"|",",")
                    g = Replace(rs("percentunknownna"),"|",",")
    Else
    mswUnknownna_query = "SELECT Count (distinct UIN) unknowncountna FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswUnknownna_query,conn 
    
                    f= Replace(rs("unknowncountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswTotal <> 0 Then
    mswUnknown_query = "SELECT Count (distinct UIN) unknowncount, round(Count(distinct UIN)* 100 /Cast("&mswTotal&" as float), 2) percentunknown FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswUnknown_query,conn 
    
                    h = Replace(rs("unknowncount") ,"|",",")
                    i = Replace(rs("percentunknown"),"|",",")
    Else
    mswUnknown_query = "SELECT Count (distinct UIN) unknowncount FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswUnknown_query,conn 
    
                    h = Replace(rs("unknowncount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
             

     '//Total//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswtotalfemales <> 0 Then
					mswTotalfemale_query="SELECT Count (distinct UIN) totalcountfemale, round(Count(distinct UIN)* 100 /Cast("&mswtotalfemales&" as float), 2) percenttotalfemales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' "
					rs.Open mswTotalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total","|",",")
                    b = Replace(rs("totalcountfemale"),"|",",")
                    c = Replace(rs("percenttotalfemales"),"|",",")
    End If 
    Else
    	mswTotalfemale_query="SELECT Count (distinct UIN) totalcountfemale FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' "
					rs.Open mswTotalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total","|",",")
                    b = Replace(rs("totalcountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If 
    End If
     rs.close

    If mswtotalmales <> 0 Then
    mswTotalmale_query = "SELECT Count (distinct UIN) totalcountmale, round(Count(distinct UIN)* 100 /Cast("&mswtotalmales&" as float), 2) percenttotalmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' "
    rs.Open mswTotalmale_query,conn 
    
                    d= Replace(rs("totalcountmale"),"|",",")
                    e = Replace(rs("percenttotalmales"),"|",",")
    Else
    mswTotalmale_query = "SELECT Count (distinct UIN) totalcountmale FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'M' "
    rs.Open mswTotalmale_query,conn 
    
                    d= Replace(rs("totalcountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
     If mswtotalunavailable <> 0 Then
    mswTotalna_query = "SELECT Count (distinct UIN) totalcountna, round(Count(distinct UIN)* 100 /Cast("&mswtotalunavailable&" as float), 2) percenttotalna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M')"
    rs.Open mswTotalna_query,conn 
    
                    f= Replace(rs("totalcountna"),"|",",")
                    g = Replace(rs("percenttotalna"),"|",",")
    Else
    mswTotalna_query = "SELECT Count (distinct UIN) totalcountna FROM Applicants where Degree_Program = 'MSW' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender not in ('F', 'M')"
    rs.Open mswTotalna_query,conn 
    
                    f= Replace(rs("totalcountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close
    
     If mswTotal <> 0 Then
    mswTotal_query = "SELECT Count (distinct UIN) totalcount, round(Count(distinct UIN)* 100 /Cast("&mswTotal&" as float), 2) percenttotal FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' "
    rs.Open mswTotal_query,conn 
    
                    h = Replace(rs("totalcount") ,"|",",")
                    i = Replace(rs("percenttotal"),"|",",")
    Else
    mswTotal_query = "SELECT Count (distinct UIN) totalcount FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' "
    rs.Open mswTotal_query,conn 
    
                    h = Replace(rs("totalcount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                

    pdf.Ln(10)

   ' //// Full Time ////

    '///AfericanAmerican///

pdf.OrangeTitle("Full Time")

totalfulltime_rows = 9
totalfulltime_cols = 9

Dim totalft_col(9)

totalft_col(1) = "Race"
totalft_col(2) = "Total Female"
totalft_col(3) = "% Female"
totalft_col(4) = "Total Male"
totalft_col(5) = "% Male"
totalft_col(6) = "Total Not Available"
totalft_col(7) = "% Not Available"
totalft_col(8) = "Total"
totalft_col(9) = "Percent"

pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 30,25,20,20,15,25,25,15,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Race","TotalFemale", "%Female", "Total Male", "%Male","Total Not Available", "% Not Available", "Total", "Percent"
pdf.SetFont "Arial","",10

    set rs=Server.CreateObject("ADODB.recordset")
					mswftTotalFemales_query="SELECT Count (distinct UIN) numberfemales FROM Applicants where Degree_Program='MSW' and Program_Type = 'FT'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' "
					rs.Open mswftTotalFemales_query,conn 
    mswfttotalfemales = rs("numberfemales")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
					mswftTotalMales_query="SELECT Count (distinct UIN) numbermales FROM Applicants where Degree_Program='MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit')  and Term_CD='"&AdmitTerm&"' and Gender = 'M'"
					rs.Open mswftTotalMales_query,conn 
    mswfttotalmales = rs("numbermales")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
					mswftTotalNA_query="SELECT Count (distinct UIN) numberna FROM Applicants where Degree_Program='MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F', 'M') "
					rs.Open mswftTotalNA_query,conn 
    mswfttotalna = rs("numberna")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
mswftTotal = mswfttotalfemales + mswfttotalmales + mswfttotalna

                    set rs=Server.CreateObject("ADODB.recordset")
    If mswfttotalfemales <> 0 Then
					mswftAfricanAmericanfemale_query="SELECT Count (distinct UIN) africanamericancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalfemales&" as float), 2) percentafricanamericanfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Black/African American'"
					rs.Open mswftAfricanAmericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("African/American","|",",")
                    b = Replace(rs("africanamericancountfemale"),"|",",")
                    c = Replace(rs("percentafricanamericanfemales"),"|",",")
    End If
    Else
    		mswftAfricanAmericanfemale_query="SELECT Count (distinct UIN) africanamericancountfemale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
					rs.Open mswftAfricanAmericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("African/American","|",",")
                    b = Replace(rs("africanamericancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswfttotalmales <> 0 Then
    mswftAfricanAmericanmale_query = "SELECT Count (distinct UIN) africanamericancountmales, round(Count(distinct UIN)* 100 /Cast("&mswfttotalmales&" as float), 2) percentafricanamericanmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit')  and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Black/African American'"
    rs.Open mswftAfricanAmericanmale_query,conn 
    
                    d= Replace(rs("africanamericancountmales"),"|",",")
                    e = Replace(rs("percentafricanamericanmales"),"|",",")
    Else
    mswftAfricanAmericanmale_query = "SELECT Count (distinct UIN) africanamericancountmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Admission_decision IN ('A', 'S', 'ReAdmit')  and Race_ethinicity = 'Black/African American'"
    rs.Open mswftAfricanAmericanmale_query,conn 
    
                    d= Replace(rs("africanamericancountmales"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswfttotalna <> 0 Then
    mswftAfricanAmericanna_query = "SELECT Count (distinct UIN) africanamericancountna, round(Count(distinct UIN)* 100 /Cast("&mswfttotalmales&" as float), 2) percentafricanamericanna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F') and Race_ethinicity = 'Black/African American'"
    rs.Open mswftAfricanAmericanna_query,conn 
    
                    f = Replace(rs("africanamericancountna"),"|",",")
                    g = Replace(rs("percentafricanamericanna"),"|",",")
    Else
    mswftAfricanAmericanna_query = "SELECT Count (distinct UIN) africanamericancountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswftAfricanAmericanna_query,conn 
    
                    f = Replace(rs("africanamericancountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close
    
    If mswftTotal <> 0 Then
    mswftAfricanAmerican_query = "SELECT Count (distinct UIN) africanamericancount, round(Count(distinct UIN)* 100 /Cast("&mswftTotal&" as float), 2) percentafricanamerican FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswftAfricanAmerican_query,conn 
    
                    h = Replace(rs("africanamericancount") ,"|",",")
                    i = Replace(rs("percentafricanamerican"),"|",",")
    Else
    mswftAfricanAmerican_query = "SELECT Count (distinct UIN) africanamericancount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswftAfricanAmerican_query,conn 
    
                    h = Replace(rs("africanamericancount") ,"|",",")
                    i = Replace("0","|",",")
       End If 
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
               

    '//Hispanic//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswfttotalfemales <> 0 Then
					mswftHispanicfemale_query="SELECT Count (distinct UIN) hispaniccountfemale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalfemales&" as float), 2) percenthispanicfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT'  and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Hispanic'"
					rs.Open mswftHispanicfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Hispanic","|",",")
                    b = Replace(rs("hispaniccountfemale"),"|",",")
                    c = Replace(rs("percenthispanicfemales"),"|",",")
    End If
    Else
    	mswftHispanicfemale_query="SELECT Count (distinct UIN) hispaniccountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
					rs.Open mswftHispanicfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Hispanic","|",",")
                    b = Replace(rs("hispaniccountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswfttotalmales <> 0 Then
    mswftHispanicmale_query = "SELECT Count (distinct UIN) hispaniccountmales, round(Count(distinct UIN)* 100 /Cast("&mswfttotalmales&" as float), 2) percenthispanicmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Hispanic'"
    rs.Open mswftHispanicmale_query,conn 
    
                    d= Replace(rs("hispaniccountmales"),"|",",")
                    e = Replace(rs("percenthispanicmales"),"|",",")
    Else
    mswftHispanicmale_query = "SELECT Count (distinct UIN) hispaniccountmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswftHispanicmale_query,conn 
    
                    d= Replace(rs("hispaniccountmales"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswfttotalna <> 0 Then
    mswftHispanicna_query = "SELECT Count (distinct UIN) hispaniccountna, round(Count(distinct UIN)* 100 /Cast("&mswfttotalmales&" as float), 2) percenthispanicna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F') and Race_ethinicity = 'Hispanic'"
    rs.Open mswftHispanicna_query,conn 
    
                    f = Replace(rs("hispaniccountna"),"|",",")
                    g = Replace(rs("percenthispanicna"),"|",",")
    Else
    mswftHispanicna_query = "SELECT Count (distinct UIN) hispaniccountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswftHispanicna_query,conn 
    
                    f = Replace(rs("hispaniccountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close
    
     If mswftTotal <> 0 Then
    mswftHispanic_query = "SELECT Count (distinct UIN) hispaniccount, round(Count(distinct UIN)* 100 /Cast("&mswftTotal&" as float), 2) percenthispanic FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Hispanic'"
    rs.Open mswftHispanic_query,conn 
    
                    h = Replace(rs("hispaniccount") ,"|",",")
                    i = Replace(rs("percenthispanic"),"|",",")
    Else
     mswftHispanic_query = "SELECT Count (distinct UIN) hispaniccount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswftHispanic_query,conn 
    
                    h = Replace(rs("hispaniccount") ,"|",",")
                    i = Replace("0","|",",")
   End If 
     rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                 
     '//Asian//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswfttotalfemales <> 0 Then
					mswftAsianfemale_query="SELECT Count (distinct UIN) asiancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalfemales&" as float), 2) percentasianfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT'  and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Asian'"
					rs.Open mswftAsianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Asian","|",",")
                    b = Replace(rs("asiancountfemale"),"|",",")
                    c = Replace(rs("percentasianfemales"),"|",",")
    End If
    Else
    mswftAsianfemale_query="SELECT Count (distinct UIN) asiancountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
					rs.Open mswftAsianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Asian","|",",")
                    b = Replace(rs("asiancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswfttotalmales <> 0 Then
    mswftAsianmale_query = "SELECT Count (distinct UIN) asiancountmales, round(Count(distinct UIN)* 100 /Cast("&mswfttotalmales&" as float), 2) percentasianmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT'  and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Asian'"
    rs.Open mswftAsianmale_query,conn 
    
                    d= Replace(rs("asiancountmales"),"|",",")
                    e = Replace(rs("percentasianmales"),"|",",")
    Else
    mswftAsianmale_query = "SELECT Count (distinct UIN) asiancountmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswftAsianmale_query,conn 
    
                    d= Replace(rs("asiancountmales"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswfttotalna <> 0 Then
    mswftAsianna_query = "SELECT Count (distinct UIN) asiancountna, round(Count(distinct UIN)* 100 /Cast("&mswfttotalna&" as float), 2) percentasianna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F') and Race_ethinicity = 'Asian'"
    rs.Open mswftAsianna_query,conn 
    
                    f = Replace(rs("asiancountna"),"|",",")
                    g = Replace(rs("percentasianna"),"|",",")
    Else
    mswftAsianna_query = "SELECT Count (distinct UIN) asiancountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswftAsianna_query,conn 
    
                    f = Replace(rs("asiancountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close
    
     If mswftTotal <> 0 Then
    mswftAsian_query = "SELECT Count (distinct UIN) asiancount, round(Count(distinct UIN)* 100 /Cast("&mswftTotal&" as float), 2) percentasian FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Asian'"
    rs.Open mswftAsian_query,conn 
    
                    h = Replace(rs("asiancount") ,"|",",")
                    i = Replace(rs("percentasian"),"|",",")
    Else
    mswftAsian_query = "SELECT Count (distinct UIN) asiancount FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT'  and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswftAsian_query,conn 
    
                    h = Replace(rs("asiancount") ,"|",",")
                    i = Replace("0","|",",")
     End If 
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                
     
    '//Native American//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswfttotalfemales <> 0 Then
					mswftNativeamericanfemale_query="SELECT Count (distinct UIN) nativeamericancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalfemales&" as float), 2) percentnativeamericanfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit')  and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and NHPI_Race_Ind = 'Y'"
					rs.Open mswftNativeamericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Native American","|",",")
                    b = Replace(rs("nativeamericancountfemale"),"|",",")
                    c = Replace(rs("percentnativeamericanfemales"),"|",",")
    End If
    Else
    mswftNativeamericanfemale_query="SELECT Count (distinct UIN) nativeamericancountfemale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' and NHPI_Race_Ind = 'Y'"
					rs.Open mswftNativeamericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Native American","|",",")
                    b = Replace(rs("nativeamericancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswfttotalmales <> 0 Then
    mswftNativeamericanmale_query = "SELECT Count (distinct UIN) nativeamericancountmale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalmales&" as float), 2) percentnativeamericanmales FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and NHPI_Race_Ind = 'Y'"
    rs.Open mswftNativeamericanmale_query,conn 
    
                    d= Replace(rs("nativeamericancountmale"),"|",",")
                    e = Replace(rs("percentnativeamericanmales"),"|",",")
    Else
    mswftNativeamericanmale_query = "SELECT Count (distinct UIN) nativeamericancountmale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswftNativeamericanmale_query,conn 
    
                    d= Replace(rs("nativeamericancountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
     
    If mswfttotalna <> 0 Then
    mswftNativeamericanna_query = "SELECT Count (distinct UIN) nativeamericancountna, round(Count(distinct UIN)* 100 /Cast("&mswfttotalna&" as float), 2) percentnativeamericanna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F') and NHPI_Race_Ind = 'Y'"
    rs.Open mswftNativeamericanna_query,conn 
    
                    f = Replace(rs("nativeamericancountna"),"|",",")
                    g = Replace(rs("percentnativeamericanna"),"|",",")
    Else
    mswftNativeamericanna_query = "SELECT Count (distinct UIN) nativeamericancountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswftNativeamericanna_query,conn 
    
                    f = Replace(rs("nativeamericancountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close
    
     If mswftTotal <> 0 Then
    mswftNativeAmerican_query = "SELECT Count (distinct UIN) nativeamericancount, round(Count(distinct UIN)* 100 /Cast("&mswftTotal&" as float), 2) percentnativeamerican FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and NHPI_Race_Ind = 'Y'"
    rs.Open mswftNativeAmerican_query,conn 
    
                    h = Replace(rs("nativeamericancount") ,"|",",")
                    i = Replace(rs("percentnativeamerican"),"|",",")
    Else
    mswftNativeAmerican_query = "SELECT Count (distinct UIN) nativeamericancount FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswftNativeAmerican_query,conn 
    
                    h = Replace(rs("nativeamericancount") ,"|",",")
                    i = Replace("0","|",",")
     End If 
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
    
    '//International//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswfttotalfemales <> 0 Then
					mswftinternationalfemale_query="SELECT Count (distinct UIN) internationalcountfemale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalfemales&" as float), 2) percentinternationalfemales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'International'"
					rs.Open mswftinternationalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("International","|",",")
                    b = Replace(rs("internationalcountfemale"),"|",",")
                    c = Replace(rs("percentinternationalfemales"),"|",",")
    End If
    Else
    mswftinternationalfemale_query="SELECT Count (distinct UIN) internationalcountfemale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
					rs.Open mswftinternationalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("International","|",",")
                    b = Replace(rs("internationalcountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close


    If mswfttotalmales <> 0 Then
    mswftinternationalmale_query = "SELECT Count (distinct UIN) internationalcountmale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalmales&" as float), 2) percentinternationalmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'International'"
    rs.Open mswftinternationalmale_query,conn 
    
                    d= Replace(rs("internationalcountmale"),"|",",")
                    e = Replace(rs("percentinternationalmales"),"|",",")
    Else
    mswftinternationalmale_query = "SELECT Count (distinct UIN) internationalcountmale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswftinternationalmale_query,conn 
    
                    d= Replace(rs("internationalcountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close

    If mswfttotalna <> 0 Then
    mswftinternationalna_query = "SELECT Count (distinct UIN) internationalcountna, round(Count(distinct UIN)* 100 /Cast("&mswfttotalna&" as float), 2) percentinternationalna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F') and Race_ethinicity = 'International'"
    rs.Open mswftinternationalna_query,conn 
    
                    f = Replace(rs("internationalcountna"),"|",",")
                    g = Replace(rs("percentinternationalna"),"|",",")
    Else
    mswftinternationalna_query = "SELECT Count (distinct UIN) internationalcountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswftinternationalna_query,conn 
    
                    f = Replace(rs("internationalcountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close
    
    If mswftTotal <> 0 Then
    mswftinternational_query = "SELECT Count (distinct UIN) internationalcount, round(Count(distinct UIN)* 100 /Cast("&mswftTotal&" as float), 2) percentinternational FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'International'"
    rs.Open mswftinternational_query,conn 
    
                    h = Replace(rs("internationalcount") ,"|",",")
                    i = Replace(rs("percentinternational"),"|",",")
    Else
    mswftinternational_query = "SELECT Count (distinct UIN) internationalcount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswftinternational_query,conn 
    
                    h = Replace(rs("internationalcount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                

     '//Multi-Race//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswfttotalfemales <> 0 Then
					mswftmultiracefemale_query="SELECT Count (distinct UIN) multiracecountfemale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalfemales&" as float), 2) percentmultiracefemales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Multi-Race'"
					rs.Open mswftmultiracefemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Multi-Race","|",",")
                    b = Replace(rs("multiracecountfemale"),"|",",")
                    c = Replace(rs("percentmultiracefemales"),"|",",")
    End If
    Else
    mswftmultiracefemale_query="SELECT Count (distinct UIN) multiracecountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
					rs.Open mswftmultiracefemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Multi-Race","|",",")
                    b = Replace(rs("multiracecountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

     If mswfttotalmales <> 0 Then
    mswftmultiracemale_query = "SELECT Count (distinct UIN) multiracecountmale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalmales&" as float), 2) percentmultiracemales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Multi-Race'"
    rs.Open mswftmultiracemale_query,conn 
    
                    d= Replace(rs("multiracecountmale"),"|",",")
                    e = Replace(rs("percentmultiracemales"),"|",",")
    Else
     mswftmultiracemale_query = "SELECT Count (distinct UIN) multiracecountmale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswftmultiracemale_query,conn 
    
                    d= Replace(rs("multiracecountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswfttotalna <> 0 Then
    mswftmultiracena_query = "SELECT Count (distinct UIN) multiracecountna, round(Count(distinct UIN)* 100 /Cast("&mswfttotalna&" as float), 2) percentmultiracena FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswftmultiracena_query,conn 
    
                    f = Replace(rs("multiracecountna"),"|",",")
                    g = Replace(rs("percentmultiracena"),"|",",")
    Else
    mswftmultiracena_query = "SELECT Count (distinct UIN) multiracecountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswftmultiracena_query,conn 
    
                    f = Replace(rs("multiracecountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswftTotal <> 0 Then
    mswftmultirace_query = "SELECT Count (distinct UIN) multiracecount, round(Count(distinct UIN)* 100 /Cast("&mswftTotal&" as float), 2) percentmultirace FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Multi-Race'"
    rs.Open mswftmultirace_query,conn 
    
                    h = Replace(rs("multiracecount") ,"|",",")
                    i = Replace(rs("percentmultirace"),"|",",")
    Else
    mswftmultirace_query = "SELECT Count (distinct UIN) multiracecount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswftmultirace_query,conn 
    
                    h = Replace(rs("multiracecount") ,"|",",")
                    i = Replace("0","|",",")
       End If 
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
             
    
    '//Total Minority//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswfttotalfemales <> 0 Then
					mswftminorityfemale_query="SELECT Count (distinct UIN) minoritycountfemale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalfemales&" as float), 2) percentminorityfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity != 'White'"
					rs.Open mswftminorityfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total Minority","|",",")
                    b = Replace(rs("minoritycountfemale"),"|",",")
                    c = Replace(rs("percentminorityfemales"),"|",",")
    End If
    Else
    mswftminorityfemale_query="SELECT Count (distinct UIN) minoritycountfemale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' and Race_ethinicity != 'White'"
					rs.Open mswftminorityfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total Minority","|",",")
                    b = Replace(rs("minoritycountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close


    If mswfttotalmales <> 0 Then
    mswftminoritymale_query = "SELECT Count (distinct UIN) minoritycountmale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalmales&" as float), 2) percentminoritymales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity != 'White'"
    rs.Open mswftminoritymale_query,conn 
    
                    d= Replace(rs("minoritycountmale"),"|",",")
                    e = Replace(rs("percentminoritymales"),"|",",")
    Else
    mswftminoritymale_query = "SELECT Count (distinct UIN) minoritycountmale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswftminoritymale_query,conn 
    
                    d= Replace(rs("minoritycountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswfttotalna <> 0 Then
    mswftminorityna_query = "SELECT Count (distinct UIN) minoritycountna, round(Count(distinct UIN)* 100 /Cast("&mswfttotalna&" as float), 2) percentminorityna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT'  and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F') and Race_ethinicity != 'White'"
    rs.Open mswftminorityna_query,conn 
    
                    f = Replace(rs("minoritycountna"),"|",",")
                    g = Replace(rs("percentminorityna"),"|",",")
    Else
    mswftminorityna_query = "SELECT Count (distinct UIN) minoritycountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswftminorityna_query,conn 
    
                    f = Replace(rs("minoritycountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswftTotal <> 0 Then
    mswftMinority_query = "SELECT Count (distinct UIN) minoritycount, round(Count(distinct UIN)* 100 /Cast("&mswftTotal&" as float), 2) percentminority FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Race_ethinicity != 'White'"
    rs.Open mswftMinority_query,conn 
    
                    h = Replace(rs("minoritycount") ,"|",",")
                    i = Replace(rs("percentminority"),"|",",")
    Else
    mswftMinority_query = "SELECT Count (distinct UIN) minoritycount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswftMinority_query,conn 
    
                    h = Replace(rs("minoritycount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                

     '//Caucasian//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswfttotalfemales <> 0 Then
					mswftcaucasianfemale_query="SELECT Count (distinct UIN) caucasiancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalfemales&" as float), 2) percentcaucasianfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'White'"
					rs.Open mswftcaucasianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Caucasian","|",",")
                    b = Replace(rs("caucasiancountfemale"),"|",",")
                    c = Replace(rs("percentcaucasianfemales"),"|",",")
    End If
    Else
    mswftcaucasianfemale_query="SELECT Count (distinct UIN) caucasiancountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' and Race_ethinicity = 'White'"
					rs.Open mswftcaucasianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Caucasian","|",",")
                    b = Replace(rs("caucasiancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

     If mswfttotalmales <> 0 Then
    mswftcaucasianmale_query = "SELECT Count (distinct UIN) caucasiancountmale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalmales&" as float), 2) percentcaucasianmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'White'"
    rs.Open mswftcaucasianmale_query,conn 
    
                    d= Replace(rs("caucasiancountmale"),"|",",")
                    e = Replace(rs("percentcaucasianmales"),"|",",")
    Else
     mswftcaucasianmale_query = "SELECT Count (distinct UIN) caucasiancountmale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'M' and Race_ethinicity = 'White'"
    rs.Open mswftcaucasianmale_query,conn 
    
                    d= Replace(rs("caucasiancountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswfttotalna <> 0 Then
    mswftcaucasianna_query = "SELECT Count (distinct UIN) caucasiancountna, round(Count(distinct UIN)* 100 /Cast("&mswfttotalna&" as float), 2) percentcaucasianna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F') and Race_ethinicity = 'White'"
    rs.Open mswftcaucasianna_query,conn 
    
                    f = Replace(rs("caucasiancountna"),"|",",")
                    g = Replace(rs("percentcaucasianna"),"|",",")
    Else
    mswftcaucasianna_query = "SELECT Count (distinct UIN) caucasiancountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswftcaucasianna_query,conn 
    
                    f = Replace(rs("caucasiancountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswftTotal <> 0 Then
    mswftCaucasian_query = "SELECT Count (distinct UIN) caucasiancount, round(Count(distinct UIN)* 100 /Cast("&mswftTotal&" as float), 2) percentcaucasian FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'White'"
    rs.Open mswftCaucasian_query,conn 
    
                    h = Replace(rs("caucasiancount") ,"|",",")
                    i = Replace(rs("percentcaucasian"),"|",",")
    Else
    mswftCaucasian_query = "SELECT Count (distinct UIN) caucasiancount FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswftCaucasian_query,conn 
    
                    h = Replace(rs("caucasiancount") ,"|",",")
                    i = Replace("0","|",",")
       End If 
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   

    '//Unknown//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswfttotalfemales <> 0 Then
					mswftUnknownfemale_query="SELECT Count (distinct UIN) unknowncountfemale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalfemales&" as float), 2) percentunknownfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Unknown'"
					rs.Open mswftUnknownfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Undeclared","|",",")
                    b = Replace(rs("unknowncountfemale"),"|",",")
                    c = Replace(rs("percentunknownfemales"),"|",",")
    End If
    Else
    mswftUnknownfemale_query="SELECT Count (distinct UIN) unknowncountfemale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
					rs.Open mswftUnknownfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Undeclared","|",",")
                    b = Replace(rs("unknowncountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswfttotalmales <> 0 Then
    mswftUnknownmale_query = "SELECT Count (distinct UIN) unknowncountmale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalmales&" as float), 2) percentunknownmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Unknown'"
    rs.Open mswftUnknownmale_query,conn 
    
                    d= Replace(rs("unknowncountmale"),"|",",")
                    e = Replace(rs("percentunknownmales"),"|",",")
    Else
    mswftUnknownmale_query = "SELECT Count (distinct UIN) unknowncountmale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswftUnknownmale_query,conn 
    
                    d= Replace(rs("unknowncountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswfttotalna <> 0 Then
    mswftUnknownna_query = "SELECT Count (distinct UIN) unknowncountna, round(Count(distinct UIN)* 100 /Cast("&mswfttotalna&" as float), 2) percentunknownna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswftUnknownna_query,conn 
    
                    f = Replace(rs("unknowncountna"),"|",",")
                    g = Replace(rs("percentunknownna"),"|",",")
    Else
    mswftUnknownna_query = "SELECT Count (distinct UIN) unknowncountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswftUnknownna_query,conn 
    
                    f = Replace(rs("unknowncountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswftTotal <> 0 Then
    mswftUnknown_query = "SELECT Count (distinct UIN) unknowncount, round(Count(distinct UIN)* 100 /Cast("&mswftTotal&" as float), 2) percentunknown FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswftUnknown_query,conn 
    
                    h = Replace(rs("unknowncount") ,"|",",")
                    i = Replace(rs("percentunknown"),"|",",")
    Else
    mswftUnknown_query = "SELECT Count (distinct UIN) unknowncount FROM Applicants where Degree_Program = 'MSW'   and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswftUnknown_query,conn 
    
                    h = Replace(rs("unknowncount") ,"|",",")
                    i = Replace("0","|",",")
      End If 
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   

     '//Total//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswfttotalfemales <> 0 Then
					mswftTotalfemale_query="SELECT Count (distinct UIN) totalcountfemale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalfemales&" as float), 2) percenttotalfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender = 'F' "
					rs.Open mswftTotalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total","|",",")
                    b = Replace(rs("totalcountfemale"),"|",",")
                    c = Replace(rs("percenttotalfemales"),"|",",")
    End If
    Else
    mswftTotalfemale_query="SELECT Count (distinct UIN) totalcountfemale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' "
					rs.Open mswftTotalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total","|",",")
                    b = Replace(rs("totalcountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswfttotalmales <> 0 Then
    mswftTotalmale_query = "SELECT Count (distinct UIN) totalcountmale, round(Count(distinct UIN)* 100 /Cast("&mswfttotalmales&" as float), 2) percenttotalmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' "
    rs.Open mswftTotalmale_query,conn 
    
                    d= Replace(rs("totalcountmale"),"|",",")
                    e = Replace(rs("percenttotalmales"),"|",",")
    Else
    mswftTotalmale_query = "SELECT Count (distinct UIN) totalcountmale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'M' "
    rs.Open mswftTotalmale_query,conn 
    
                    d= Replace(rs("totalcountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close

    If mswfttotalna <> 0 Then
    mswftTotalna_query = "SELECT Count (distinct UIN) totalcountna, round(Count(distinct UIN)* 100 /Cast("&mswfttotalna&" as float), 2) percenttotalna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' and Gender not in ('M', 'F')"
    rs.Open mswftTotalna_query,conn 
    
                    f = Replace(rs("totalcountna"),"|",",")
                    g = Replace(rs("percenttotalna"),"|",",")
    Else
    mswftTotalna_query = "SELECT Count (distinct UIN) totalcountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender not in ('M', 'F')"
    rs.Open mswftTotalna_query,conn 
    
                    f = Replace(rs("totalcountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswftTotal <> 0 Then
    mswFTTotal_query = "SELECT Count (distinct UIN) totalcount, round(Count(distinct UIN)* 100 /Cast("&mswftTotal&" as float), 2) percenttotal FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'FT' and Term_CD='"&AdmitTerm&"' "
    rs.Open mswFTTotal_query,conn 
    
                    f = Replace(rs("totalcount") ,"|",",")
                    g = Replace(rs("percenttotal"),"|",",")
    Else
    mswFTTotal_query = "SELECT Count (distinct UIN) totalcount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'FT'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' "
    rs.Open mswFTTotal_query,conn 
    
                    f = Replace(rs("totalcount") ,"|",",")
                    g = Replace("0","|",",")
    End If 
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                 
pdf.Ln(10)
    
   ' //// PM ////

    '///AfericanAmerican///

pdf.OrangeTitle("PM")

totalpm_rows = 9
totalpm_cols = 9

Dim totalpm_col(9)

totalpm_col(1) = "Race"
totalpm_col(2) = "Total Female"
totalpm_col(3) = "% Female"
totalpm_col(4) = "Total Male"
totalpm_col(5) = "% Male"
totalpm_col(4) = "Total Not Available"
totalpm_col(5) = "% Not Available"
totalpm_col(6) = "Total"
totalpm_col(7) = "Percent"

pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 30,25,20,20,15,25,25,15,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Race","TotalFemale", "%Female", "Total Male", "%Male","Total Not Available", "% Not Available", "Total", "Percent"
pdf.SetFont "Arial","",10

    set rs=Server.CreateObject("ADODB.recordset")
					mswpmTotalFemales_query="SELECT Count (distinct UIN) numberfemales FROM Applicants where Degree_Program='MSW' and Program_Type = 'PM'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' "
					rs.Open mswpmTotalFemales_query,conn 
    mswpmtotalfemales = rs("numberfemales")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
					mswpmTotalMales_query="SELECT Count (distinct UIN) numbermales FROM Applicants where Degree_Program='MSW' and Program_Type = 'PM'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M'"
					rs.Open mswpmTotalMales_query,conn 
    mswpmtotalmales = rs("numbermales")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
					mswpmTotalNA_query="SELECT Count (distinct UIN) numberna FROM Applicants where Degree_Program='MSW' and Program_Type = 'PM'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')"
					rs.Open mswpmTotalNA_query,conn 
    mswpmtotalna = rs("numberna")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
mswpmTotal = mswpmtotalfemales + mswpmtotalmales

                    set rs=Server.CreateObject("ADODB.recordset")
    If mswpmtotalfemales <> 0 Then
					mswpmAfricanAmericanfemale_query="SELECT Count (distinct UIN) africanamericancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalfemales&" as float), 2) percentafricanamericanfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Black/African American'"
					rs.Open mswpmAfricanAmericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("African/American","|",",")
                    b = Replace(rs("africanamericancountfemale"),"|",",")
                    c = Replace(rs("percentafricanamericanfemales"),"|",",")
    End If
    Else
    mswpmAfricanAmericanfemale_query="SELECT Count (distinct UIN) africanamericancountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' and Race_ethinicity = 'Black/African American'"
					rs.Open mswpmAfricanAmericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("African/American","|",",")
                    b = Replace(rs("africanamericancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswpmtotalmales <> 0 Then
    mswpmAfricanAmericanmale_query = "SELECT Count (distinct UIN) africanamericancountmales, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalmales&" as float), 2) percentafricanamericanmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Black/African American'"
    rs.Open mswpmAfricanAmericanmale_query,conn 
    
                    d= Replace(rs("africanamericancountmales"),"|",",")
                    e = Replace(rs("percentafricanamericanmales"),"|",",")
    Else
    mswpmAfricanAmericanmale_query = "SELECT Count (distinct UIN) africanamericancountmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswpmAfricanAmericanmale_query,conn 
    
                    d= Replace(rs("africanamericancountmales"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close

    If mswpmtotalna <> 0 Then
    mswpmAfricanAmericanna_query = "SELECT Count (distinct UIN) africanamericancountna, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalna&" as float), 2) percentafricanamericanna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity = 'Black/African American'"
    rs.Open mswpmAfricanAmericanna_query,conn 
    
                    f = Replace(rs("africanamericancountna"),"|",",")
                    g = Replace(rs("percentafricanamericanna"),"|",",")
    Else
    mswpmAfricanAmericanna_query = "SELECT Count (distinct UIN) africanamericancountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswpmAfricanAmericanna_query,conn 
    
                    f = Replace(rs("africanamericancountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswpmTotal <> 0 Then
    mswpmAfricanAmerican_query = "SELECT Count (distinct UIN) africanamericancount, round(Count(distinct UIN)* 100 /Cast("&mswpmTotal&" as float), 2) percentafricanamerican FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Black/African American'"
    rs.Open mswpmAfricanAmerican_query,conn 
    
                    h = Replace(rs("africanamericancount") ,"|",",")
                    i = Replace(rs("percentafricanamerican"),"|",",")
    Else
    mswpmAfricanAmerican_query = "SELECT Count (distinct UIN) africanamericancount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswpmAfricanAmerican_query,conn 
    
                    h = Replace(rs("africanamericancount") ,"|",",")
                    i = Replace("0","|",",")
     End If  
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                

    '//Hispanic//
   
                    set rs=Server.CreateObject("ADODB.recordset")
     If mswpmtotalfemales <> 0 Then
					mswpmHispanicfemale_query="SELECT Count (distinct UIN) hispaniccountfemale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalfemales&" as float), 2) percenthispanicfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Hispanic'"
					rs.Open mswpmHispanicfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Hispanic","|",",")
                    b = Replace(rs("hispaniccountfemale"),"|",",")
                    c = Replace(rs("percenthispanicfemales"),"|",",")
    End If
    Else 
    mswpmHispanicfemale_query="SELECT Count (distinct UIN) hispaniccountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
					rs.Open mswpmHispanicfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Hispanic","|",",")
                    b = Replace(rs("hispaniccountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

     If mswpmtotalmales <> 0 Then
    mswpmHispanicmale_query = "SELECT Count (distinct UIN) hispaniccountmales, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalmales&" as float), 2) percenthispanicmales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Hispanic'"
    rs.Open mswpmHispanicmale_query,conn 
    
                    d= Replace(rs("hispaniccountmales"),"|",",")
                    e = Replace(rs("percenthispanicmales"),"|",",")
    Else
    mswpmHispanicmale_query = "SELECT Count (distinct UIN) hispaniccountmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswpmHispanicmale_query,conn 
    
                    d= Replace(rs("hispaniccountmales"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswpmtotalna <> 0 Then
    mswpmHispanicna_query = "SELECT Count (distinct UIN) hispaniccountna, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalna&" as float), 2) percenthispanicna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity = 'Hispanic'"
    rs.Open mswpmHispanicna_query,conn 
    
                    f = Replace(rs("hispaniccountna"),"|",",")
                    g = Replace(rs("percenthispanicna"),"|",",")
    Else
    mswpmHispanicna_query = "SELECT Count (distinct UIN) hispaniccountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswpmHispanicna_query,conn 
    
                    f = Replace(rs("hispaniccountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswpmTotal <> 0 Then
    mswpmHispanic_query = "SELECT Count (distinct UIN) hispaniccount, round(Count(distinct UIN)* 100 /Cast("&mswpmTotal&" as float), 2) percenthispanic FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Hispanic'"
    rs.Open mswpmHispanic_query,conn 
    
                    h = Replace(rs("hispaniccount") ,"|",",")
                    i = Replace(rs("percenthispanic"),"|",",")
    Else
    mswpmHispanic_query = "SELECT Count (distinct UIN) hispaniccount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswpmHispanic_query,conn 
    
                    h = Replace(rs("hispaniccount") ,"|",",")
                    i = Replace("0","|",",")
        End If  
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
             
     '//Asian//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswpmtotalfemales <> 0 Then
					mswpmAsianfemale_query="SELECT Count (distinct UIN) asiancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalfemales&" as float), 2) percentasianfemales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM'  and Admission_decision IN ('A', 'S', 'ReAdmit')  and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Asian'"
					rs.Open mswpmAsianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Asian","|",",")
                    b = Replace(rs("asiancountfemale"),"|",",")
                    c = Replace(rs("percentasianfemales"),"|",",")
    End If
    Else
    mswpmAsianfemale_query="SELECT Count (distinct UIN) asiancountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM'  and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
					rs.Open mswpmAsianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Asian","|",",")
                    b = Replace(rs("asiancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswpmtotalmales <> 0 Then
    mswpmAsianmale_query = "SELECT Count (distinct UIN) asiancountmales, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalmales&" as float), 2) percentasianmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM'  and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Asian'"
    rs.Open mswpmAsianmale_query,conn 
    
                    d= Replace(rs("asiancountmales"),"|",",")
                    e = Replace(rs("percentasianmales"),"|",",")
    Else
    mswpmAsianmale_query = "SELECT Count (distinct UIN) asiancountmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM'  and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswpmAsianmale_query,conn 
    
                    d= Replace(rs("asiancountmales"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswpmtotalna <> 0 Then
    mswpmAsianna_query = "SELECT Count (distinct UIN) asiancountna, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalna&" as float), 2) percentasianna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity = 'Asian'"
    rs.Open mswpmAsianna_query,conn 
    
                    f = Replace(rs("asiancountna"),"|",",")
                    g = Replace(rs("percentasianna"),"|",",")
    Else
    mswpmAsianna_query = "SELECT Count (distinct UIN) asiancountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswpmAsianna_query,conn 
    
                    f = Replace(rs("asiancountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswpmTotal <> 0 Then
    mswpmAsian_query = "SELECT Count (distinct UIN) asiancount, round(Count(distinct UIN)* 100 /Cast("&mswpmTotal&" as float), 2) percentasian FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Asian'"
    rs.Open mswpmAsian_query,conn 
    
                    h = Replace(rs("asiancount") ,"|",",")
                    i = Replace(rs("percentasian"),"|",",")
    Else
    mswpmAsian_query = "SELECT Count (distinct UIN) asiancount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswpmAsian_query,conn 
    
                    h = Replace(rs("asiancount") ,"|",",")
                    i = Replace("0","|",",")
       End If 
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
               
     
    '//Native American//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswpmtotalfemales <> 0 Then
					mswpmNativeamericanfemale_query="SELECT Count (distinct UIN) nativeamericancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalfemales&" as float), 2) percentnativeamericanfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and NHPI_Race_Ind = 'Y'"
					rs.Open mswpmNativeamericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Native American","|",",")
                    b = Replace(rs("nativeamericancountfemale"),"|",",")
                    c = Replace(rs("percentnativeamericanfemales"),"|",",")
    End If
    Else
    mswpmNativeamericanfemale_query="SELECT Count (distinct UIN) nativeamericancountfemale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
					rs.Open mswpmNativeamericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Native American","|",",")
                    b = Replace(rs("nativeamericancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswpmtotalmales <> 0 Then
    mswpmNativeamericanmale_query = "SELECT Count (distinct UIN) nativeamericancountmale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalmales&" as float), 2) percentnativeamericanmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and NHPI_Race_Ind = 'Y'"
    rs.Open mswpmNativeamericanmale_query,conn 
    
                    d= Replace(rs("nativeamericancountmale"),"|",",")
                    e = Replace(rs("percentnativeamericanmales"),"|",",")
    Else
    mswpmNativeamericanmale_query = "SELECT Count (distinct UIN) nativeamericancountmale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswpmNativeamericanmale_query,conn 
    
                    d= Replace(rs("nativeamericancountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
     
    If mswpmtotalna <> 0 Then
    mswpmNativeamericanna_query = "SELECT Count (distinct UIN) nativeamericancountna, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalna&" as float), 2) percentnativeamericanna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and NHPI_Race_Ind = 'Y'"
    rs.Open mswpmNativeamericanna_query,conn 
    
                    f = Replace(rs("nativeamericancountna"),"|",",")
                    g = Replace(rs("percentnativeamericanna"),"|",",")
    Else
    mswpmNativeamericanna_query = "SELECT Count (distinct UIN) nativeamericancountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender not in ('M','F') and NHPI_Race_Ind = 'Y'"
    rs.Open mswpmNativeamericanna_query,conn 
    
                    f = Replace(rs("nativeamericancountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswpmTotal <> 0 Then
    mswpmNativeAmerican_query = "SELECT Count (distinct UIN) nativeamericancount, round(Count(distinct UIN)* 100 /Cast("&mswpmTotal&" as float), 2) percentnativeamerican FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and NHPI_Race_Ind = 'Y'"
    rs.Open mswpmNativeAmerican_query,conn 
    
                    h = Replace(rs("nativeamericancount") ,"|",",")
                    i = Replace(rs("percentnativeamerican"),"|",",")
    Else
    mswpmNativeAmerican_query = "SELECT Count (distinct UIN) nativeamericancount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswpmNativeAmerican_query,conn 
    
                    h = Replace(rs("nativeamericancount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                 
    
    '//International//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswpmtotalfemales <> 0 Then 
					mswpminternationalfemale_query="SELECT Count (distinct UIN) internationalcountfemale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalfemales&" as float), 2) percentinternationalfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'International'"
					rs.Open mswpminternationalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("International","|",",")
                    b = Replace(rs("internationalcountfemale"),"|",",")
                    c = Replace(rs("percentinternationalfemales"),"|",",")
    End If
    Else
    mswpminternationalfemale_query="SELECT Count (distinct UIN) internationalcountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
					rs.Open mswpminternationalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("International","|",",")
                    b = Replace(rs("internationalcountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswpmtotalmales <> 0 Then
    mswpminternationalmale_query = "SELECT Count (distinct UIN) internationalcountmale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalmales&" as float), 2) percentinternationalmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'International'"
    rs.Open mswpminternationalmale_query,conn 
    
                    d= Replace(rs("internationalcountmale"),"|",",")
                    e = Replace(rs("percentinternationalmales"),"|",",")
    Else
    mswpminternationalmale_query = "SELECT Count (distinct UIN) internationalcountmale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswpminternationalmale_query,conn 
    
                    d= Replace(rs("internationalcountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
     
    If mswpmtotalna <> 0 Then
    mswpminternationalna_query = "SELECT Count (distinct UIN) internationalcountna, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalna&" as float), 2) percentinternationalna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity = 'International'"
    rs.Open mswpminternationalna_query,conn 
    
                    f = Replace(rs("internationalcountna"),"|",",")
                    g = Replace(rs("percentinternationalna"),"|",",")
    Else
    mswpminternationalna_query = "SELECT Count (distinct UIN) internationalcountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswpminternationalna_query,conn 
    
                    f = Replace(rs("internationalcountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswpmTotal <> 0 Then
    mswpminternational_query = "SELECT Count (distinct UIN) internationalcount, round(Count(distinct UIN)* 100 /Cast("&mswpmTotal&" as float), 2) percentinternational FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'International'"
    rs.Open mswpminternational_query,conn 
    
                    h = Replace(rs("internationalcount") ,"|",",")
                    i = Replace(rs("percentinternational"),"|",",")
    Else
    mswpminternational_query = "SELECT Count (distinct UIN) internationalcount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswpminternational_query,conn 
    
                    h = Replace(rs("internationalcount") ,"|",",")
                    i = Replace("0","|",",")
    End If

   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                

     '//Multi-Race//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswpmtotalfemales <> 0 Then
					mswpmmultiracefemale_query="SELECT Count (distinct UIN) multiracecountfemale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalfemales&" as float), 2) percentmultiracefemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Multi-Race'"
					rs.Open mswpmmultiracefemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Multi-Race","|",",")
                    b = Replace(rs("multiracecountfemale"),"|",",")
                    c = Replace(rs("percentmultiracefemales"),"|",",")
    End If
    Else
    mswpmmultiracefemale_query="SELECT Count (distinct UIN) multiracecountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
					rs.Open mswpmmultiracefemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Multi-Race","|",",")
                    b = Replace(rs("multiracecountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswpmtotalmales <> 0 Then
    mswpmmultiracemale_query = "SELECT Count (distinct UIN) multiracecountmale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalmales&" as float), 2) percentmultiracemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Multi-Race'"
    rs.Open mswpmmultiracemale_query,conn 
    
                    d= Replace(rs("multiracecountmale"),"|",",")
                    e = Replace(rs("percentmultiracemales"),"|",",")
    Else
    mswpmmultiracemale_query = "SELECT Count (distinct UIN) multiracecountmale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswpmmultiracemale_query,conn 
    
                    d= Replace(rs("multiracecountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
      
    If mswpmtotalna <> 0 Then
    mswpmmultiracena_query = "SELECT Count (distinct UIN) multiracecountna, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalna&" as float), 2) percentmultiracena FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswpmmultiracena_query,conn 
    
                    f = Replace(rs("multiracecountna"),"|",",")
                    g = Replace(rs("percentmultiracena"),"|",",")
    Else
    mswpmmultiracena_query = "SELECT Count (distinct UIN) multiracecountna FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswpmmultiracena_query,conn 
    
                    f = Replace(rs("multiracecountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswpmTotal <> 0 Then
    mswpmmultirace_query = "SELECT Count (distinct UIN) multiracecount, round(Count(distinct UIN)* 100 /Cast("&mswpmTotal&" as float), 2) percentmultirace FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Multi-Race'"
    rs.Open mswpmmultirace_query,conn 
    
                    h = Replace(rs("multiracecount") ,"|",",")
                    i = Replace(rs("percentmultirace"),"|",",")
    Else
     mswpmmultirace_query = "SELECT Count (distinct UIN) multiracecount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswpmmultirace_query,conn 
    
                    h = Replace(rs("multiracecount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
             
              
    
    '//Total Minority//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswpmtotalfemales <> 0 Then 
					mswpmminorityfemale_query="SELECT Count (distinct UIN) minoritycountfemale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalfemales&" as float), 2) percentminorityfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity != 'White'"
					rs.Open mswpmminorityfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total Minority","|",",")
                    b = Replace(rs("minoritycountfemale"),"|",",")
                    c = Replace(rs("percentminorityfemales"),"|",",")
    End If
    Else
    mswpmminorityfemale_query="SELECT Count (distinct UIN) minoritycountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
					rs.Open mswpmminorityfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total Minority","|",",")
                    b = Replace(rs("minoritycountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswpmtotalmales <> 0 Then
    mswpmminoritymale_query = "SELECT Count (distinct UIN) minoritycountmale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalmales&" as float), 2) percentminoritymales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity != 'White'"
    rs.Open mswpmminoritymale_query,conn 
    
                    d= Replace(rs("minoritycountmale"),"|",",")
                    e = Replace(rs("percentminoritymales"),"|",",")
    Else
    mswpmminoritymale_query = "SELECT Count (distinct UIN) minoritycountmale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswpmminoritymale_query,conn 
    
                    d= Replace(rs("minoritycountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
      
    If mswpmtotalna <> 0 Then
    mswpmminorityna_query = "SELECT Count (distinct UIN) minoritycountna, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalna&" as float), 2) percentminorityna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity != 'White'"
    rs.Open mswpmminorityna_query,conn 
    
                    f = Replace(rs("minoritycountna"),"|",",")
                    g = Replace(rs("percentminorityna"),"|",",")
    Else
    mswpmminorityna_query = "SELECT Count (distinct UIN) minoritycountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswpmminorityna_query,conn 
    
                    f = Replace(rs("minoritycountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswpmTotal <> 0 Then
    mswpmMinority_query = "SELECT Count (distinct UIN) minoritycount, round(Count(distinct UIN)* 100 /Cast("&mswpmTotal&" as float), 2) percentminority FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity != 'White'"
    rs.Open mswpmMinority_query,conn 
    
                    h = Replace(rs("minoritycount") ,"|",",")
                    i = Replace(rs("percentminority"),"|",",")
    Else
    mswpmMinority_query = "SELECT Count (distinct UIN) minoritycount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswpmMinority_query,conn 
    
                    h = Replace(rs("minoritycount") ,"|",",")
                    i = Replace("0","|",",")
    End If

   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                

     '//Caucasian//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswpmtotalfemales <> 0 Then
					mswpmcaucasianfemale_query="SELECT Count (distinct UIN) caucasiancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalfemales&" as float), 2) percentcaucasianfemales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'White'"
					rs.Open mswpmcaucasianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Caucasian","|",",")
                    b = Replace(rs("caucasiancountfemale"),"|",",")
                    c = Replace(rs("percentcaucasianfemales"),"|",",")
    End If
    Else
    mswpmcaucasianfemale_query="SELECT Count (distinct UIN) caucasiancountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
					rs.Open mswpmcaucasianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Caucasian","|",",")
                    b = Replace(rs("caucasiancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswpmtotalmales <> 0 Then
    mswpmcaucasianmale_query = "SELECT Count (distinct UIN) caucasiancountmale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalmales&" as float), 2) percentcaucasianmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'White'"
    rs.Open mswpmcaucasianmale_query,conn 
    
                    d= Replace(rs("caucasiancountmale"),"|",",")
                    e = Replace(rs("percentcaucasianmales"),"|",",")
    Else
    mswpmcaucasianmale_query = "SELECT Count (distinct UIN) caucasiancountmale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswpmcaucasianmale_query,conn 
    
                    d= Replace(rs("caucasiancountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
      
    If mswpmtotalna <> 0 Then
    mswpmcaucasianna_query = "SELECT Count (distinct UIN) caucasiancountna, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalna&" as float), 2) percentcaucasianna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity = 'White'"
    rs.Open mswpmcaucasianna_query,conn 
    
                    f = Replace(rs("caucasiancountna"),"|",",")
                    g = Replace(rs("percentcaucasianna"),"|",",")
    Else
    mswpmcaucasianna_query = "SELECT Count (distinct UIN) caucasiancountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswpmcaucasianna_query,conn 
    
                    f = Replace(rs("caucasiancountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswpmTotal <> 0 Then
    mswpmCaucasian_query = "SELECT Count (distinct UIN) caucasiancount, round(Count(distinct UIN)* 100 /Cast("&mswpmTotal&" as float), 2) percentcaucasian FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'White'"
    rs.Open mswpmCaucasian_query,conn 
    
                    h = Replace(rs("caucasiancount") ,"|",",")
                    i = Replace(rs("percentcaucasian"),"|",",")
    Else
     mswpmCaucasian_query = "SELECT Count (distinct UIN) caucasiancount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswpmCaucasian_query,conn 
    
                    h = Replace(rs("caucasiancount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
             


    '//Unknown//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    IF mswpmtotalfemales <> 0 Then
					mswpmUnknownfemale_query="SELECT Count (distinct UIN) unknowncountfemale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalfemales&" as float), 2) percentunknownfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Unknown'"
					rs.Open mswpmUnknownfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Undeclared","|",",")
                    b = Replace(rs("unknowncountfemale"),"|",",")
                    c = Replace(rs("percentunknownfemales"),"|",",")
    End If
    Else
    mswpmUnknownfemale_query="SELECT Count (distinct UIN) unknowncountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' and Race_ethinicity = 'Unknown'"
					rs.Open mswpmUnknownfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Undeclared","|",",")
                    b = Replace(rs("unknowncountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswpmtotalmales <> 0 Then
    mswpmUnknownmale_query = "SELECT Count (distinct UIN) unknowncountmale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalmales&" as float), 2) percentunknownmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Unknown'"
    rs.Open mswpmUnknownmale_query,conn 
    
                    d= Replace(rs("unknowncountmale"),"|",",")
                    e = Replace(rs("percentunknownmales"),"|",",")
    Else
    mswpmUnknownmale_query = "SELECT Count (distinct UIN) unknowncountmale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswpmUnknownmale_query,conn 
    
                    d= Replace(rs("unknowncountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
      
    If mswpmtotalna <> 0 Then
    mswpmUnknownna_query = "SELECT Count (distinct UIN) unknowncountna, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalna&" as float), 2) percentcaucasianna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity = 'Unknown'"
    rs.Open mswpmUnknownna_query,conn 
    
                    f = Replace(rs("unknowncountna"),"|",",")
                    g = Replace(rs("percentunknownna"),"|",",")
    Else
    mswpmUnknownna_query = "SELECT Count (distinct UIN) unknowncountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswpmUnknownna_query,conn 
    
                    f = Replace(rs("unknowncountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswpmTotal <> 0 Then
    mswpmUnknown_query = "SELECT Count (distinct UIN) unknowncount, round(Count(distinct UIN)* 100 /Cast("&mswpmTotal&" as float), 2) percentunknown FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Unknown'"
    rs.Open mswpmUnknown_query,conn 
    
                    h = Replace(rs("unknowncount") ,"|",",")
                    i = Replace(rs("percentunknown"),"|",",")
    Else
    mswpmUnknown_query = "SELECT Count (distinct UIN) unknowncount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswpmUnknown_query,conn 
    
                    h = Replace(rs("unknowncount") ,"|",",")
                    i = Replace("0","|",",")
    End If 

   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                 
     '//Total//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswpmtotalfemales <> 0 Then
					mswpmTotalfemale_query="SELECT Count (distinct UIN) totalcountfemale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalfemales&" as float), 2) percenttotalfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'F' "
					rs.Open mswpmTotalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total","|",",")
                    b = Replace(rs("totalcountfemale"),"|",",")
                    c = Replace(rs("percenttotalfemales"),"|",",")
    End If 
    Else
    mswpmTotalfemale_query="SELECT Count (distinct UIN) totalcountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' "
					rs.Open mswpmTotalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total","|",",")
                    b = Replace(rs("totalcountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If 
    End If
     rs.close

    If mswpmtotalmales <> 0 Then
    mswpmTotalmale_query = "SELECT Count (distinct UIN) totalcountmale, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalmales&" as float), 2) percenttotalmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender = 'M' "
    rs.Open mswpmTotalmale_query,conn 
    
                    d= Replace(rs("totalcountmale"),"|",",")
                    e = Replace(rs("percenttotalmales"),"|",",")
    Else
    mswpmTotalmale_query = "SELECT Count (distinct UIN) totalcountmale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'M' "
    rs.Open mswpmTotalmale_query,conn 
    
                    d= Replace(rs("totalcountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
      
    If mswpmtotalna <> 0 Then
    mswpmTotalna_query = "SELECT Count (distinct UIN) totalcountna, round(Count(distinct UIN)* 100 /Cast("&mswpmtotalna&" as float), 2) percenttotalna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')"
    rs.Open mswpmTotalna_query,conn 
    
                    f = Replace(rs("totalcountna"),"|",",")
                    g = Replace(rs("percenttotalna"),"|",",")
    Else
    mswpmTotalna_query = "SELECT Count (distinct UIN) totalcountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender not in ('M','F')"
    rs.Open mswpmTotalna_query,conn 
    
                    f = Replace(rs("totalcountna"),"|",",")
                    g = Replace("0","|",",")
    End If
    rs.close

    If mswpmTotal <> 0 Then
    mswPMTotal_query = "SELECT Count (distinct UIN) totalcount, round(Count(distinct UIN)* 100 /Cast("&mswpmTotal&" as float), 2) percenttotal FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'PM' and Term_CD='"&AdmitTerm&"' "
    rs.Open mswPMTotal_query,conn 
    
                    h = Replace(rs("totalcount") ,"|",",")
                    i = Replace(rs("percenttotal"),"|",",")
    Else
    mswPMTotal_query = "SELECT Count (distinct UIN) totalcount FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'PM'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' "
    rs.Open mswPMTotal_query,conn 
    
                    h = Replace(rs("totalcount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                

    
pdf.Ln(10)
    
   ' //// ADV ////

    '///AfericanAmerican///

pdf.OrangeTitle("Advanced Standing")

totaladv_rows = 9
totaladv_cols = 9

Dim totaladv_col(9)

totaladv_col(1) = "Race"
totaladv_col(2) = "Total Female"
totaladv_col(3) = "% Female"
totaladv_col(4) = "Total Male"
totaladv_col(5) = "% Male"
totaladv_col(6) = "Total Not Available"
totaladv_col(7) = "% Not Available"
totaladv_col(8) = "Total"
totaladv_col(9) = "Percent"

pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 30,25,20,20,15,25,25,15,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Race","TotalFemale", "%Female", "Total Male", "%Male","Total Not Available", "% Not Available", "Total", "Percent"
pdf.SetFont "Arial","",10

    set rs=Server.CreateObject("ADODB.recordset")
					mswadvTotalFemales_query="SELECT Count (distinct UIN) numberfemales FROM Applicants where Degree_Program='MSW' and Program_Type = 'ADV'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' "
					rs.Open mswadvTotalFemales_query,conn 
    mswadvtotalfemales = rs("numberfemales")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
					mswadvTotalMales_query="SELECT Count (distinct UIN) numbermales FROM Applicants where Degree_Program='MSW' and Program_Type = 'ADV'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M'"
					rs.Open mswadvTotalMales_query,conn 
    mswadvtotalmales = rs("numbermales")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
					mswadvTotalNA_query="SELECT Count (distinct UIN) numberna FROM Applicants where Degree_Program='MSW' and Program_Type = 'ADV'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')"
					rs.Open mswadvTotalNA_query,conn 
    mswadvtotalna = rs("numberna")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
mswadvTotal = mswadvtotalfemales + mswadvtotalmales + mswadvtotalna

                    set rs=Server.CreateObject("ADODB.recordset")
    If mswadvtotalfemales <> 0 Then
					mswadvAfricanAmericanfemale_query="SELECT Count (distinct UIN) africanamericancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalfemales&" as float), 2) percentafricanamericanfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Black/African American'"
					rs.Open mswadvAfricanAmericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("African/American","|",",")
                    b = Replace(rs("africanamericancountfemale"),"|",",")
                    c = Replace(rs("percentafricanamericanfemales"),"|",",")
    End If
    Else
    mswadvAfricanAmericanfemale_query="SELECT Count (distinct UIN) africanamericancountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' and Race_ethinicity = 'Black/African American'"
					rs.Open mswadvAfricanAmericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("African/American","|",",")
                    b = Replace(rs("africanamericancountfemale"),"|",",")
                    c = Replace(rs("0"),"|",",")
    End If
    End If
     rs.close

    If mswadvtotalmales <> 0 Then
    mswadvAfricanAmericanmale_query = "SELECT Count (distinct UIN) africanamericancountmales, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalmales&" as float), 2) percentafricanamericanmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Black/African American'"
    rs.Open mswadvAfricanAmericanmale_query,conn 
    
                    d= Replace(rs("africanamericancountmales"),"|",",")
                    e = Replace(rs("percentafricanamericanmales"),"|",",")



    Else
    mswadvAfricanAmericanmale_query = "SELECT Count (distinct UIN) africanamericancountmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswadvAfricanAmericanmale_query,conn 
    
                    d= Replace(rs("africanamericancountmales"),"|",",")
                    e = Replace("0","|",",")

    End If
    rs.close
    
    If mswadvtotalna <> 0 Then
    mswadvAfricanAmericanna_query = "SELECT Count (distinct UIN) africanamericancountna, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalna&" as float), 2) percentafricanamericanna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity = 'Black/African American'"
    rs.Open mswadvAfricanAmericanna_query,conn 
    
                    f = Replace(rs("africanamericancountna"),"|",",")
                    g = Replace(rs("percentafricanamericanna"),"|",",")



    Else
    mswadvAfricanAmericanna_query = "SELECT Count (distinct UIN) africanamericancountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswadvAfricanAmericanna_query,conn 
    
                    f = Replace(rs("africanamericancountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswadvTotal <> 0 Then
    mswadvAfricanAmerican_query = "SELECT Count (distinct UIN) africanamericancount, round(Count(distinct UIN)* 100 /Cast("&mswadvTotal&" as float), 2) percentafricanamerican FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Black/African American'"
    rs.Open mswadvAfricanAmerican_query,conn 
    
                    h = Replace(rs("africanamericancount") ,"|",",")
                    i = Replace(rs("percentafricanamerican"),"|",",")
    
    Else 
    mswadvAfricanAmerican_query = "SELECT Count (distinct UIN) africanamericancount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswadvAfricanAmerican_query,conn 
    
                    h = Replace(rs("africanamericancount") ,"|",",")
                    i = Replace("0","|",",")
    
    End If
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
         

    '//Hispanic//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswadvtotalfemales <> 0 Then
					mswadvHispanicfemale_query="SELECT Count (distinct UIN) hispaniccountfemale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalfemales&" as float), 2) percenthispanicfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Hispanic'"
					rs.Open mswadvHispanicfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Hispanic","|",",")
                    b = Replace(rs("hispaniccountfemale"),"|",",")
                    c = Replace(rs("percenthispanicfemales"),"|",",")
    End If
    Else
    	mswadvHispanicfemale_query="SELECT Count (distinct UIN) hispaniccountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
					rs.Open mswadvHispanicfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Hispanic","|",",")
                    b = Replace(rs("hispaniccountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswadvtotalmales <> 0 Then
    mswadvHispanicmale_query = "SELECT Count (distinct UIN) hispaniccountmales, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalmales&" as float), 2) percenthispanicmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Hispanic'"
    rs.Open mswadvHispanicmale_query,conn 
    
                    d= Replace(rs("hispaniccountmales"),"|",",")
                    e = Replace(rs("percenthispanicmales"),"|",",")
    
    Else
    
    mswadvHispanicmale_query = "SELECT Count (distinct UIN) hispaniccountmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswadvHispanicmale_query,conn 
    
                    d= Replace(rs("hispaniccountmales"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswadvtotalna <> 0 Then
    mswadvHispanicna_query = "SELECT Count (distinct UIN) hispaniccountna, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalna&" as float), 2) percenthispanicna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender not in ('M','F') and Race_ethinicity = 'Hispanic'"
    rs.Open mswadvHispanicna_query,conn 
    
                    f = Replace(rs("hispaniccountna"),"|",",")
                    g = Replace(rs("percenthispanicna"),"|",",")



    Else
    mswadvHispanicna_query = "SELECT Count (distinct UIN) hispaniccountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswadvHispanicna_query,conn 
    
                    f = Replace(rs("hispaniccountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswadvTotal <> 0 Then
    mswadvHispanic_query = "SELECT Count (distinct UIN) hispaniccount, round(Count(distinct UIN)* 100 /Cast("&mswadvTotal&" as float), 2) percenthispanic FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Hispanic'"
    rs.Open mswadvHispanic_query,conn 
    
                    h = Replace(rs("hispaniccount") ,"|",",")
                    i = Replace(rs("percenthispanic"),"|",",")
    
    Else
    mswadvHispanic_query = "SELECT Count (distinct UIN) hispaniccount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswadvHispanic_query,conn 
    
                    h = Replace(rs("hispaniccount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
         
     '//Asian//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswadvtotalfemales <> 0 Then
					mswadvAsianfemale_query="SELECT Count (distinct UIN) asiancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalfemales&" as float), 2) percentasianfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Asian'"
					rs.Open mswadvAsianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Asian","|",",")
                    b = Replace(rs("asiancountfemale"),"|",",")
                    c = Replace(rs("percentasianfemales"),"|",",")
    End If
    Else
    mswadvAsianfemale_query="SELECT Count (distinct UIN) asiancountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' and Race_ethinicity = 'Asian'"
					rs.Open mswadvAsianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Asian","|",",")
                    b = Replace(rs("asiancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswadvtotalmales <> 0 Then
    mswadvAsianmale_query = "SELECT Count (distinct UIN) asiancountmales, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalmales&" as float), 2) percentasianmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV'  and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Asian'"
    rs.Open mswadvAsianmale_query,conn 
    
                    d= Replace(rs("asiancountmales"),"|",",")
                    e = Replace(rs("percentasianmales"),"|",",")
    
    Else
    mswadvAsianmale_query = "SELECT Count (distinct UIN) asiancountmales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswadvAsianmale_query,conn 
    
                    d= Replace(rs("asiancountmales"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close

    If mswadvtotalna <> 0 Then
    mswadvAsianna_query = "SELECT Count (distinct UIN) asiancountna, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalna&" as float), 2) percentasianna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity = 'Asian'"
    rs.Open mswadvAsianna_query,conn 
    
                    f = Replace(rs("asiancountna"),"|",",")
                    g = Replace(rs("percentasianna"),"|",",")



    Else
    mswadvAsianna_query = "SELECT Count (distinct UIN) asiancountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswadvAsianna_query,conn 
    
                    f = Replace(rs("asiancountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswadvTotal <> 0 Then
    mswadvAsian_query = "SELECT Count (distinct UIN) asiancount, round(Count(distinct UIN)* 100 /Cast("&mswadvTotal&" as float), 2) percentasian FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Asian'"
    rs.Open mswadvAsian_query,conn 
    
                    h = Replace(rs("asiancount") ,"|",",")
                    i = Replace(rs("percentasian"),"|",",")
    
    Else
    mswadvAsian_query = "SELECT Count (distinct UIN) asiancount FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswadvAsian_query,conn 
    
                    h = Replace(rs("asiancount") ,"|",",")
                    i = Replace("0","|",",")
   End If
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
         
     
    '//Native American//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswadvtotalfemales <> 0 Then
					mswadvNativeamericanfemale_query="SELECT Count (distinct UIN) nativeamericancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalfemales&" as float), 2) percentnativeamericanfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and NHPI_Race_Ind = 'Y'"
					rs.Open mswadvNativeamericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Native American","|",",")
                    b = Replace(rs("nativeamericancountfemale"),"|",",")
                    c = Replace(rs("percentnativeamericanfemales"),"|",",")
    End If
    Else
    mswadvNativeamericanfemale_query="SELECT Count (distinct UIN) nativeamericancountfemale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
					rs.Open mswadvNativeamericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Native American","|",",")
                    b = Replace(rs("nativeamericancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If

     rs.close

    If mswadvtotalmales <> 0 Then
    mswadvNativeamericanmale_query = "SELECT Count (distinct UIN) nativeamericancountmale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalmales&" as float), 2) percentnativeamericanmales FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and NHPI_Race_Ind = 'Y'"
    rs.Open mswadvNativeamericanmale_query,conn 
    
                    d= Replace(rs("nativeamericancountmale"),"|",",")
                    e = Replace(rs("percentnativeamericanmales"),"|",",")
    
    Else
     mswadvNativeamericanmale_query = "SELECT Count (distinct UIN) nativeamericancountmale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswadvNativeamericanmale_query,conn 
    
                    d= Replace(rs("nativeamericancountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswadvtotalna <> 0 Then
    mswadvNativeamericanna_query = "SELECT Count (distinct UIN) nativeamericancountna, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalna&" as float), 2) percentnativeamericanna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and NHPI_Race_Ind = 'Y'"
    rs.Open mswadvNativeamericanna_query,conn 
    
                    f = Replace(rs("nativeamericancountna"),"|",",")
                    g = Replace(rs("percentnativeamericanna"),"|",",")



    Else
    mswadvNativeamericanna_query = "SELECT Count (distinct UIN) nativeamericancountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswadvNativeamericanna_query,conn 
    
                    f = Replace(rs("nativeamericancountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

     If mswadvTotal <> 0 Then
    mswadvNativeAmerican_query = "SELECT Count (distinct UIN) nativeamericancount, round(Count(distinct UIN)* 100 /Cast("&mswadvTotal&" as float), 2) percentnativeamerican FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and NHPI_Race_Ind = 'Y'"
    rs.Open mswadvNativeAmerican_query,conn 
    
                    h = Replace(rs("nativeamericancount") ,"|",",")
                    i = Replace(rs("percentnativeamerican"),"|",",")
    
    Else
     mswadvNativeAmerican_query = "SELECT Count (distinct UIN) nativeamericancount FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswadvNativeAmerican_query,conn 
    
                    h = Replace(rs("nativeamericancount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i

                    
    
    '//International//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswadvtotalfemales <> 0 Then
					mswadvinternationalfemale_query="SELECT Count (distinct UIN) internationalcountfemale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalfemales&" as float), 2) percentinternationalfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'International'"
					rs.Open mswadvinternationalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("International","|",",")
                    b = Replace(rs("internationalcountfemale"),"|",",")
                    c = Replace(rs("percentinternationalfemales"),"|",",")
    End If
    Else
    mswadvinternationalfemale_query="SELECT Count (distinct UIN) minoritycountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
					rs.Open mswadvinternationalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("International","|",",")
                    b = Replace(rs("internationalcountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If 
    rs.close

    If mswadvtotalmales <> 0 Then
    mswadvinternationalmale_query = "SELECT Count (distinct UIN) internationalcountmale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalmales&" as float), 2) percentinternationalmales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'International'"
    rs.Open mswadvinternationalmale_query,conn 
    
                    d= Replace(rs("internationalcountmale"),"|",",")
                    e = Replace(rs("percentinternationalmales"),"|",",")
     
    Else
     mswadvinternationalmale_query = "SELECT Count (distinct UIN) internationalcountmale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswadvinternationalmale_query,conn 
    
                    d= Replace(rs("internationalcountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    
    rs.close
    
    If mswadvtotalna <> 0 Then
    mswadvinternationalna_query = "SELECT Count (distinct UIN) internationalcountna, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalna&" as float), 2) percentinternationalna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity = 'International'"
    rs.Open mswadvinternationalna_query,conn 
    
                    f = Replace(rs("internationalcountna"),"|",",")
                    g = Replace(rs("percentinternationalna"),"|",",")



    Else
    mswadvinternationalna_query = "SELECT Count (distinct UIN) internationalcountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswadvinternationalna_query,conn 
    
                    f = Replace(rs("internationalcountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswadvTotal <> 0 Then
    mswadvinternational_query = "SELECT Count (distinct UIN) internationalcount, round(Count(distinct UIN)* 100 /Cast("&mswadvTotal&" as float), 2) percentinternational FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'International'"
    rs.Open mswadvinternational_query,conn 
    
                    h = Replace(rs("internationalcount") ,"|",",")
                    i = Replace(rs("percentinternational"),"|",",")
    
    Else
    mswadvinternational_query = "SELECT Count (distinct UIN) internationalcount FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswadvinternational_query,conn 
    
                    h = Replace(rs("internationalcount") ,"|",",")
                    i = Replace("0","|",",")
    End If


   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                 

     '//Multi-Race//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    
    If mswadvtotalfemales <> 0 Then
					mswadvmultiracefemale_query="SELECT Count (distinct UIN) multiracecountfemale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalfemales&" as float), 2) percentmultiracefemales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Multi-Race'"
					rs.Open mswadvmultiracefemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Multi-Race","|",",")
                    b = Replace(rs("multiracecountfemale"),"|",",")
                    c = Replace(rs("percentmultiracefemales"),"|",",")
    End If
    Else
    mswadvmultiracefemale_query="SELECT Count (distinct UIN) multiracecountfemale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
					rs.Open mswadvmultiracefemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Multi-Race","|",",")
                    b = Replace(rs("multiracecountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If

     rs.close

    If mswadvtotalmales <> 0 Then
    mswadvmultiracemale_query = "SELECT Count (distinct UIN) multiracecountmale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalmales&" as float), 2) percentmultiracemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Multi-Race'"
    rs.Open mswadvmultiracemale_query,conn 
    
                    d= Replace(rs("multiracecountmale"),"|",",")
                    e = Replace(rs("percentmultiracemales"),"|",",")
    
    Else
    mswadvmultiracemale_query = "SELECT Count (distinct UIN) multiracecountmale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswadvmultiracemale_query,conn 
    
                    d= Replace(rs("multiracecountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
     
    If mswadvtotalna <> 0 Then
    mswadvmultiracena_query = "SELECT Count (distinct UIN) multiracecountna, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalna&" as float), 2) percentmultiracena FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswadvmultiracena_query,conn 
    
                    f = Replace(rs("multiracecountna"),"|",",")
                    g = Replace(rs("percentmultiracena"),"|",",")



    Else
    mswadvmultiracena_query = "SELECT Count (distinct UIN) multiracecountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV'  and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswadvmultiracena_query,conn 
    
                    f = Replace(rs("multiracecountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswadvTotal <> 0 Then
    mswadvmultirace_query = "SELECT Count (distinct UIN) multiracecount, round(Count(distinct UIN)* 100 /Cast("&mswadvTotal&" as float), 2) percentmultirace FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Multi-Race'"
    rs.Open mswadvmultirace_query,conn 
    
                    h = Replace(rs("multiracecount") ,"|",",")
                    i = Replace(rs("percentmultirace"),"|",",")
    
    Else
    mswadvmultirace_query = "SELECT Count (distinct UIN) multiracecount FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswadvmultirace_query,conn 
    
                    h = Replace(rs("multiracecount") ,"|",",")
                    i = Replace("0","|",",")
   End If
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   

                   
    
    '//Total Minority//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswadvtotalfemales <> 0 Then
					mswadvminorityfemale_query="SELECT Count (distinct UIN) minoritycountfemale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalfemales&" as float), 2) percentminorityfemales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity != 'White'"
					rs.Open mswadvminorityfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total Minority","|",",")
                    b = Replace(rs("minoritycountfemale"),"|",",")
                    c = Replace(rs("percentminorityfemales"),"|",",")
    End If
    Else
    mswadvminorityfemale_query="SELECT Count (distinct UIN) minoritycountfemale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
					rs.Open mswadvminorityfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total Minority","|",",")
                    b = Replace(rs("minoritycountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If 
    rs.close

    If mswadvtotalmales <> 0 Then
    mswadvminoritymale_query = "SELECT Count (distinct UIN) minoritycountmale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalmales&" as float), 2) percentminoritymales FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity != 'White'"
    rs.Open mswadvminoritymale_query,conn 
    
                    d= Replace(rs("minoritycountmale"),"|",",")
                    e = Replace(rs("percentminoritymales"),"|",",")
     
    Else
     mswadvminoritymale_query = "SELECT Count (distinct UIN) minoritycountmale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswadvminoritymale_query,conn 
    
                    d= Replace(rs("minoritycountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    
    rs.close
    
    If mswadvtotalna <> 0 Then
    mswadvminorityna_query = "SELECT Count (distinct UIN) minoritycountna, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalna&" as float), 2) percentminorityna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity != 'White'"
    rs.Open mswadvminorityna_query,conn 
    
                    f = Replace(rs("minoritycountna"),"|",",")
                    g = Replace(rs("percentminorityna"),"|",",")



    Else
    mswadvminorityna_query = "SELECT Count (distinct UIN) minoritycountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswadvminorityna_query,conn 
    
                    f = Replace(rs("minoritycountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close


    If mswadvTotal <> 0 Then
    mswadvMinority_query = "SELECT Count (distinct UIN) minoritycount, round(Count(distinct UIN)* 100 /Cast("&mswadvTotal&" as float), 2) percentminority FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Race_ethinicity != 'White'"
    rs.Open mswadvMinority_query,conn 
    
                    h = Replace(rs("minoritycount") ,"|",",")
                    i = Replace(rs("percentminority"),"|",",")
    
    Else
    mswadvMinority_query = "SELECT Count (distinct UIN) minoritycount FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswadvMinority_query,conn 
    
                    h = Replace(rs("minoritycount") ,"|",",")
                    i = Replace("0","|",",")
    End If


   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                 

     '//Caucasian//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    
    If mswadvtotalfemales <> 0 Then
					mswadvcaucasianfemale_query="SELECT Count (distinct UIN) caucasiancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalfemales&" as float), 2) percentcaucasianfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'White'"
					rs.Open mswadvcaucasianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Caucasian","|",",")
                    b = Replace(rs("caucasiancountfemale"),"|",",")
                    c = Replace(rs("percentcaucasianfemales"),"|",",")
    End If
    Else
    mswadvcaucasianfemale_query="SELECT Count (distinct UIN) caucasiancountfemale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
					rs.Open mswadvcaucasianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Caucasian","|",",")
                    b = Replace(rs("caucasiancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If

     rs.close

    If mswadvtotalmales <> 0 Then
    mswadvcaucasianmale_query = "SELECT Count (distinct UIN) caucasiancountmale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalmales&" as float), 2) percentcaucasianmales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'White'"
    rs.Open mswadvcaucasianmale_query,conn 
    
                    d= Replace(rs("caucasiancountmale"),"|",",")
                    e = Replace(rs("percentcaucasianmales"),"|",",")
    
    Else
    mswadvcaucasianmale_query = "SELECT Count (distinct UIN) caucasiancountmale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswadvcaucasianmale_query,conn 
    
                    d= Replace(rs("caucasiancountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswadvtotalna <> 0 Then
    mswadvcaucasianna_query = "SELECT Count (distinct UIN) caucasiancountna, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalna&" as float), 2) percentcaucasianna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity = 'White'"
    rs.Open mswadvcaucasianna_query,conn 
    
                    f = Replace(rs("caucasiancountna"),"|",",")
                    g = Replace(rs("percentcaucasianna"),"|",",")



    Else
    mswadvcaucasianna_query = "SELECT Count (distinct UIN) caucasiancountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswadvcaucasianna_query,conn 
    
                    f = Replace(rs("caucasiancountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswadvTotal <> 0 Then
    mswadvCaucasian_query = "SELECT Count (distinct UIN) caucasiancount, round(Count(distinct UIN)* 100 /Cast("&mswadvTotal&" as float), 2) percentcaucasian FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'White'"
    rs.Open mswadvCaucasian_query,conn 
    
                    h = Replace(rs("caucasiancount") ,"|",",")
                    i = Replace(rs("percentcaucasian"),"|",",")
    
    Else
    mswadvCaucasian_query = "SELECT Count (distinct UIN) caucasiancount FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswadvCaucasian_query,conn 
    
                    h = Replace(rs("caucasiancount") ,"|",",")
                    i = Replace("0","|",",")
   End If
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   


    '//Unknown//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswadvtotalfemales <> 0 Then
					mswadvUnknownfemale_query="SELECT Count (distinct UIN) unknowncountfemale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalfemales&" as float), 2) percentunknownfemales FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Unknown'"
					rs.Open mswadvUnknownfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Undeclared","|",",")
                    b = Replace(rs("unknowncountfemale"),"|",",")
                    c = Replace(rs("percentunknownfemales"),"|",",")
    End If
    Else
    mswadvUnknownfemale_query="SELECT Count (distinct UIN) unknowncountfemale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
					rs.Open mswadvUnknownfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Undeclared","|",",")
                    b = Replace(rs("unknowncountfemale"),"|",",")
                    c = Replace("0","|",",")

    End If
    End If
     rs.close

    If mswadvtotalmales <> 0 Then
    mswadvUnknownmale_query = "SELECT Count (distinct UIN) unknowncountmale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalmales&" as float), 2) percentunknownmales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Unknown'"
    rs.Open mswadvUnknownmale_query,conn 
    
                    d= Replace(rs("unknowncountmale"),"|",",")
                    e = Replace(rs("percentunknownmales"),"|",",")
    
    Else
     mswadvUnknownmale_query = "SELECT Count (distinct UIN) unknowncountmale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswadvUnknownmale_query,conn 
    
                    d= Replace(rs("unknowncountmale"),"|",",")
                    e = Replace("0","|",",")
    
    End If
    rs.close
    
    If mswadvtotalna <> 0 Then
    mswadvUnknownna_query = "SELECT Count (distinct UIN) unknowncountna, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalna&" as float), 2) percentunknownna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F') and Race_ethinicity = 'Unknown'"
    rs.Open mswadvUnknownna_query,conn 
    
                    f = Replace(rs("unknowncountna"),"|",",")
                    g = Replace(rs("percentunknownna"),"|",",")



    Else
    mswadvUnknownna_query = "SELECT Count (distinct UIN) unknowncountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswadvUnknownna_query,conn 
    
                    f = Replace(rs("unknowncountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswadvTotal <> 0 Then
    mswadvUnknown_query = "SELECT Count (distinct UIN) unknowncount, round(Count(distinct UIN)* 100 /Cast("&mswadvTotal&" as float), 2) percentunknown FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Unknown'"
    rs.Open mswadvUnknown_query,conn 
    
                    h = Replace(rs("unknowncount") ,"|",",")
                    i = Replace(rs("percentunknown"),"|",",")
    
    Else
    mswadvUnknown_query = "SELECT Count (distinct UIN) unknowncount FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswadvUnknown_query,conn 
    
                    h = Replace(rs("unknowncount") ,"|",",")
                    i = Replace("0","|",",")
   End If
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
     '//Total//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswadvtotalfemales <> 0 Then
					mswadvTotalfemale_query="SELECT Count (distinct UIN) totalcountfemale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalfemales&" as float), 2) percenttotalfemales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'F' "
					rs.Open mswadvTotalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total","|",",")
                    b = Replace(rs("totalcountfemale"),"|",",")
                    c = Replace(rs("percenttotalfemales"),"|",",")
    End If
    Else
    mswadvTotalfemale_query="SELECT Count (distinct UIN) totalcountfemale FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' "
					rs.Open mswadvTotalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total","|",",")
                    b = Replace(rs("totalcountfemale"),"|",",")
                    c = Replace("0","|",",") 
    End If
    End If
    rs.close

    If mswadvtotalmales <> 0 Then
    mswadvTotalmale_query = "SELECT Count (distinct UIN) totalcountmale, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalmales&" as float), 2) percenttotalmales FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender = 'M' "
    rs.Open mswadvTotalmale_query,conn 
    
                    d= Replace(rs("totalcountmale"),"|",",")
                    e = Replace(rs("percenttotalmales"),"|",",")
    
    Else
    mswadvTotalmale_query = "SELECT Count (distinct UIN) totalcountmale FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'M' "
    rs.Open mswadvTotalmale_query,conn 
    
                    d= Replace(rs("totalcountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
     
    If mswadvtotalna <> 0 Then
    mswadvTotalna_query = "SELECT Count (distinct UIN) totalcountna, round(Count(distinct UIN)* 100 /Cast("&mswadvtotalna&" as float), 2) percenttotalna FROM Applicants where Degree_Program = 'MSW'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')"
    rs.Open mswadvTotalna_query,conn 
    
                    f = Replace(rs("totalcountna"),"|",",")
                    g = Replace(rs("percenttotalna"),"|",",")



    Else
    mswadvTotalna_query = "SELECT Count (distinct UIN) totalcountna FROM Applicants where Degree_Program = 'MSW' and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender not in ('M','F')"
    rs.Open mswadvTotalna_query,conn 
    
                    f = Replace(rs("totalcountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswadvTotal <> 0 Then
    mswADVTotal_query = "SELECT Count (distinct UIN) totalcount, round(Count(distinct UIN)* 100 /Cast("&mswadvTotal&" as float), 2) percenttotal FROM Applicants where Degree_Program = 'MSW'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Program_Type = 'ADV' and Term_CD='"&AdmitTerm&"' "
    rs.Open mswADVTotal_query,conn 
    
                    h = Replace(rs("totalcount") ,"|",",")
                    i = Replace(rs("percenttotal"),"|",",")
    
    Else

    mswADVTotal_query = "SELECT Count (distinct UIN) totalcount FROM Applicants where Degree_Program = 'MSW'  and Program_Type = 'ADV'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' "
    rs.Open mswADVTotal_query,conn 
    
                    h = Replace(rs("totalcount") ,"|",",")
                    i = Replace("0","|",",")
   End If
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                  
    pdf.Ln(10)

    
   ' //// PHD ////

    '///AfericanAmerican///

pdf.OrangeTitle("PhD")

totalphd_rows = 9
totalphd_cols = 9

Dim totalphd_col(9)

totalphd_col(1) = "Race"
totalphd_col(2) = "Total Female"
totalphd_col(3) = "% Female"
totalphd_col(4) = "Total Male"
totalphd_col(5) = "% Male"
totalphd_col(6) = "Total Not Available"
totalphd_col(7) = "% Not Available"
totalphd_col(8) = "Total"
totalphd_col(9) = "Percent"

pdf.Table.Border.Width = 0.1
pdf.Table.Border.Color="006699"
'pdf.Table.Fill.Color="C9C8C0"
pdf.Table.TextAlign = "L"
pdf.SetColumns 30,25,20,20,15,25,25,15,20
'pdf.SetAligns "C", "L", "L", "C"
pdf.SetFont "Arial","B",10
pdf.Row "Race","TotalFemale", "%Female", "Total Male", "%Male", "Total Not Available", "% Not Available", "Total", "Percent"
pdf.SetFont "Arial","",10

    set rs=Server.CreateObject("ADODB.recordset")
					mswphdTotalFemales_query="SELECT Count (distinct UIN) numberfemales FROM PHDApplicants where Degree_Program='PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' "
					rs.Open mswphdTotalFemales_query,conn 
    mswphdtotalfemales = rs("numberfemales")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
    set rs=Server.CreateObject("ADODB.recordset")
					mswphdTotalMales_query="SELECT Count (distinct UIN) numbermales FROM PHDApplicants where Degree_Program='PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M'"
					rs.Open mswphdTotalMales_query,conn 
    mswphdtotalmales = rs("numbermales")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
     set rs=Server.CreateObject("ADODB.recordset")
					mswphdTotalNA_query="SELECT Count (distinct UIN) numberna FROM PHDApplicants where Degree_Program='PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('M','F')"
					rs.Open mswphdTotalNA_query,conn 
    mswphdtotalna = rs("numberna")
                  
                    If rs.EOF Then
                      
                    Else
                    'if there are records then loop through the fields
               
                    End If

 rs.close
mswphdTotal = mswphdtotalfemales + mswphdtotalmales + mswphdtotalna

                    set rs=Server.CreateObject("ADODB.recordset")
    If mswphdtotalfemales <> 0 Then
					mswphdAfricanAmericanfemale_query="SELECT Count (distinct UIN) africanamericancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalfemales&" as float), 2) percentafricanamericanfemales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Black/African American'"
					rs.Open mswphdAfricanAmericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("African/American","|",",")
                    b = Replace(rs("africanamericancountfemale"),"|",",")
                    c = Replace(rs("percentafricanamericanfemales"),"|",",")
    End If
    Else
    mswphdAfricanAmericanfemale_query="SELECT Count (distinct UIN) africanamericancountfemale FROM PHDApplicants where Degree_Program = 'PHD'  and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
					rs.Open mswphdAfricanAmericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("African/American","|",",")
                    b = Replace(rs("africanamericancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswphdtotalmales <> 0 Then
    mswphdAfricanAmericanmale_query = "SELECT Count (distinct UIN) africanamericancountmales, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalmales&" as float), 2) percentafricanamericanmales FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'M' and Race_ethinicity = 'Black/African American'"
    rs.Open mswphdAfricanAmericanmale_query,conn 
    
                    d= Replace(rs("africanamericancountmales"),"|",",")
                    e = Replace(rs("percentafricanamericanmales"),"|",",")



    Else
    mswphdAfricanAmericanmale_query = "SELECT Count (distinct UIN) africanamericancountmales FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswphdAfricanAmericanmale_query,conn 
    
                    d= Replace(rs("africanamericancountmales"),"|",",")
                    e = Replace("0","|",",")

    End If
    rs.close
    
    If mswphdtotalna <> 0 Then
    mswphdAfricanAmericanna_query = "SELECT Count (distinct UIN) africanamericancountna, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalmales&" as float), 2) percentafricanamericanna FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M') and Race_ethinicity = 'Black/African American'"
    rs.Open mswphdAfricanAmericanna_query,conn 
    
                    f = Replace(rs("africanamericancountna"),"|",",")
                    g = Replace(rs("percentafricanamericanna"),"|",",")



    Else
    mswphdAfricanAmericanna_query = "SELECT Count (distinct UIN) africanamericancountna FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswphdAfricanAmericanna_query,conn 
    
                    f = Replace(rs("africanamericancountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswphdTotal <> 0 Then
    mswphdAfricanAmerican_query = "SELECT Count (distinct UIN) africanamericancount, round(Count(distinct UIN)* 100 /Cast("&mswphdTotal&" as float), 2) percentafricanamerican FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Black/African American'"
    rs.Open mswphdAfricanAmerican_query,conn 
    
                    h = Replace(rs("africanamericancount") ,"|",",")
                    i = Replace(rs("percentafricanamerican"),"|",",")
    
    Else 
    mswphdAfricanAmerican_query = "SELECT Count (distinct UIN) africanamericancount FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Black/African American'"
    rs.Open mswphdAfricanAmerican_query,conn 
    
                    h = Replace(rs("africanamericancount") ,"|",",")
                    i = Replace("0","|",",")
    
    End If
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
         

    '//Hispanic//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswphdtotalfemales <> 0 Then
					mswphdHispanicfemale_query="SELECT Count (distinct UIN) hispaniccountfemale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalfemales&" as float), 2) percenthispanicfemales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Hispanic'"
					rs.Open mswphdHispanicfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Hispanic","|",",")
                    b = Replace(rs("hispaniccountfemale"),"|",",")
                    c = Replace(rs("percenthispanicfemales"),"|",",")
    End If
    Else
    	mswphdHispanicfemale_query="SELECT Count (distinct UIN) hispaniccountfemale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
					rs.Open mswphdHispanicfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Hispanic","|",",")
                    b = Replace(rs("hispaniccountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswphdtotalmales <> 0 Then
    mswphdHispanicmale_query = "SELECT Count (distinct UIN) hispaniccountmales, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalmales&" as float), 2) percenthispanicmales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Hispanic'"
    rs.Open mswphdHispanicmale_query,conn 
    
                    d= Replace(rs("hispaniccountmales"),"|",",")
                    e = Replace(rs("percenthispanicmales"),"|",",")
    
    Else
    
    mswphdHispanicmale_query = "SELECT Count (distinct UIN) hispaniccountmales FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswphdHispanicmale_query,conn 
    
                    d= Replace(rs("hispaniccountmales"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswphdtotalna <> 0 Then
    mswphdHispanicna_query = "SELECT Count (distinct UIN) hispaniccountna, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalmales&" as float), 2) percenthispanicna FROM PHDApplicants where Degree_Program = 'PHD'   and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M') and Race_ethinicity = 'Hispanic'"
    rs.Open mswphdHispanicna_query,conn 
    
                    f = Replace(rs("hispaniccountna"),"|",",")
                    g = Replace(rs("percenthispanicna"),"|",",")



    Else
    mswphdHispanicna_query = "SELECT Count (distinct UIN) hispaniccountna FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswphdHispanicna_query,conn 
    
                    f = Replace(rs("hispaniccountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswphdTotal <> 0 Then
    mswphdHispanic_query = "SELECT Count (distinct UIN) hispaniccount, round(Count(distinct UIN)* 100 /Cast("&mswphdTotal&" as float), 2) percenthispanic FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Hispanic'"
    rs.Open mswphdHispanic_query,conn 
    
                    h = Replace(rs("hispaniccount") ,"|",",")
                    i = Replace(rs("percenthispanic"),"|",",")
    
    Else
    mswphdHispanic_query = "SELECT Count (distinct UIN) hispaniccount FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Hispanic'"
    rs.Open mswphdHispanic_query,conn 
    
                    h = Replace(rs("hispaniccount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
         
     '//Asian//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswphdtotalfemales <> 0 Then
					mswphdAsianfemale_query="SELECT Count (distinct UIN) asiancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalfemales&" as float), 2) percentasianfemales FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
					rs.Open mswphdAsianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Asian","|",",")
                    b = Replace(rs("asiancountfemale"),"|",",")
                    c = Replace(rs("percentasianfemales"),"|",",")
    End If
    Else
    mswphdAsianfemale_query="SELECT Count (distinct UIN) asiancountfemale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
					rs.Open mswphdAsianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Asian","|",",")
                    b = Replace(rs("asiancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If
     rs.close

    If mswphdtotalmales <> 0 Then
    mswphdAsianmale_query = "SELECT Count (distinct UIN) asiancountmales, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalmales&" as float), 2) percentasianmales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Asian'"
    rs.Open mswphdAsianmale_query,conn 
    
                    d= Replace(rs("asiancountmales"),"|",",")
                    e = Replace(rs("percentasianmales"),"|",",")
    
    Else
    mswphdAsianmale_query = "SELECT Count (distinct UIN) asiancountmales FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswphdAsianmale_query,conn 
    
                    d= Replace(rs("asiancountmales"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswphdtotalna <> 0 Then
    mswphdAsianna_query = "SELECT Count (distinct UIN) asiancountna, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalna&" as float), 2) percentasianna FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M') and Race_ethinicity = 'Asian'"
    rs.Open mswphdAsianna_query,conn 
    
                    f = Replace(rs("asiancountna"),"|",",")
                    g = Replace(rs("percentasianna"),"|",",")



    Else
    mswphdAsianna_query = "SELECT Count (distinct UIN) asiancountna FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswphdAsianna_query,conn 
    
                    f = Replace(rs("asiancountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswphdTotal <> 0 Then
    mswphdAsian_query = "SELECT Count (distinct UIN) asiancount, round(Count(distinct UIN)* 100 /Cast("&mswphdTotal&" as float), 2) percentasian FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Asian'"
    rs.Open mswphdAsian_query,conn 
    
                    h = Replace(rs("asiancount") ,"|",",")
                    i = Replace(rs("percentasian"),"|",",")
    
    Else
    mswphdAsian_query = "SELECT Count (distinct UIN) asiancount FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Asian'"
    rs.Open mswphdAsian_query,conn 
    
                    h = Replace(rs("asiancount") ,"|",",")
                    i = Replace("0","|",",")
   End If
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
         
     
    '//Native American//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswphdtotalfemales <> 0 Then
					mswphdNativeamericanfemale_query="SELECT Count (distinct UIN) nativeamericancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalfemales&" as float), 2) percentnativeamericanfemales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and NHPI_Race_Ind = 'Y'"
					rs.Open mswphdNativeamericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Native American","|",",")
                    b = Replace(rs("nativeamericancountfemale"),"|",",")
                    c = Replace(rs("percentnativeamericanfemales"),"|",",")
    End If
    Else
    mswphdNativeamericanfemale_query="SELECT Count (distinct UIN) nativeamericancountfemale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
					rs.Open mswphdNativeamericanfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Native American","|",",")
                    b = Replace(rs("nativeamericancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If

     rs.close

    If mswphdtotalmales <> 0 Then
    mswphdNativeamericanmale_query = "SELECT Count (distinct UIN) nativeamericancountmale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalmales&" as float), 2) percentnativeamericanmales FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'M' and NHPI_Race_Ind = 'Y'"
    rs.Open mswphdNativeamericanmale_query,conn 
    
                    d= Replace(rs("nativeamericancountmale"),"|",",")
                    e = Replace(rs("percentnativeamericanmales"),"|",",")
    
    Else
     mswphdNativeamericanmale_query = "SELECT Count (distinct UIN) nativeamericancountmale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswphdNativeamericanmale_query,conn 
    
                    d= Replace(rs("nativeamericancountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswphdtotalna <> 0 Then
    mswphdNativeamericanna_query = "SELECT Count (distinct UIN) nativeamericancountna, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalna&" as float), 2) percentnativeamericanna FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender not in ('F','M') and NHPI_Race_Ind = 'Y'"
    rs.Open mswphdNativeamericanna_query,conn 
    
                    f = Replace(rs("nativeamericancountna"),"|",",")
                    g = Replace(rs("percentnativeamericanna"),"|",",")



    Else
    mswphdNativeamericanna_query = "SELECT Count (distinct UIN) nativeamericancountna FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswphdNativeamericanna_query,conn 
    
                    f = Replace(rs("nativeamericancountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

     If mswphdTotal <> 0 Then
    mswphdNativeAmerican_query = "SELECT Count (distinct UIN) nativeamericancount, round(Count(distinct UIN)* 100 /Cast("&mswphdTotal&" as float), 2) percentnativeamerican FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and NHPI_Race_Ind = 'Y'"
    rs.Open mswphdNativeAmerican_query,conn 
    
                    h = Replace(rs("nativeamericancount") ,"|",",")
                    i = Replace(rs("percentnativeamerican"),"|",",")
    
    Else
     mswphdNativeAmerican_query = "SELECT Count (distinct UIN) nativeamericancount FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and NHPI_Race_Ind = 'Y'"
    rs.Open mswphdNativeAmerican_query,conn 
    
                    h = Replace(rs("nativeamericancount") ,"|",",")
                    i = Replace("0","|",",")
    End If
   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
   
     
    '//International//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswphdtotalfemales <> 0 Then
					mswphdinternationalfemale_query="SELECT Count (distinct UIN) internationalcountfemale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalfemales&" as float), 2) percentinternationalfemales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'International'"
					rs.Open mswphdinternationalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("International","|",",")
                    b = Replace(rs("internationalcountfemale"),"|",",")
                    c = Replace(rs("percentinternationalfemales"),"|",",")
    End If
    Else
    mswphdinternationalfemale_query="SELECT Count (distinct UIN) internationalcountfemale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
					rs.Open mswphdinternationalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("International","|",",")
                    b = Replace(rs("internationalcountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If 
    rs.close

    If mswphdtotalmales <> 0 Then
    mswphdinternationalmale_query = "SELECT Count (distinct UIN) internationalcountmale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalmales&" as float), 2) percentinternationalmales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'International'"
    rs.Open mswphdinternationalmale_query,conn 
    
                    d= Replace(rs("internationalcountmale"),"|",",")
                    e = Replace(rs("percentinternationalmales"),"|",",")
     
    Else
     mswphdinternationalmale_query = "SELECT Count (distinct UIN) internationalcountmale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswphdinternationalmale_query,conn 
    
                    d= Replace(rs("internationalcountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    
    rs.close
    
    If mswphdtotalna <> 0 Then
    mswphdinternationalna_query = "SELECT Count (distinct UIN) internationalcountna, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalna&" as float), 2) percentinternationalna FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender not in ('F','M') and Race_ethinicity = 'International'"
    rs.Open mswphdinternationalna_query,conn 
    
                    f = Replace(rs("internationalcountna"),"|",",")
                    g = Replace(rs("percentinternationalna"),"|",",")



    Else
    mswphdinternationalna_query = "SELECT Count (distinct UIN) internationalcountna FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswphdinternationalna_query,conn 
    
                    f = Replace(rs("internationalcountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswphdTotal <> 0 Then
    mswphdinternational_query = "SELECT Count (distinct UIN) internationalcount, round(Count(distinct UIN)* 100 /Cast("&mswphdTotal&" as float), 2) percentinternational FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'International'"
    rs.Open mswphdinternational_query,conn 
    
                    h = Replace(rs("internationalcount") ,"|",",")
                    i = Replace(rs("percentinternational"),"|",",")
    
    Else
    mswphdinternational_query = "SELECT Count (distinct UIN) internationalcount FROM PHDApplicants where Degree_Program = 'PHD'  and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'International'"
    rs.Open mswphdinternational_query,conn 
    
                    h = Replace(rs("internationalcount") ,"|",",")
                    i = Replace("0","|",",")
    End If


   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                 

     '//Multi-Race//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    
    If mswphdtotalfemales <> 0 Then
					mswphdmultiracefemale_query="SELECT Count (distinct UIN) multiracecountfemale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalfemales&" as float), 2) percentmultiracefemales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Multi-Race'"
					rs.Open mswphdmultiracefemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Multi-Race","|",",")
                    b = Replace(rs("multiracecountfemale"),"|",",")
                    c = Replace(rs("percentmultiracefemales"),"|",",")
    End If
    Else
    mswphdmultiracefemale_query="SELECT Count (distinct UIN) multiracecountfemale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
					rs.Open mswphdmultiracefemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Multi-Race","|",",")
                    b = Replace(rs("multiracecountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If

     rs.close

    If mswphdtotalmales <> 0 Then
    mswphdmultiracemale_query = "SELECT Count (distinct UIN) multiracecountmale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalmales&" as float), 2) percentmultiracemales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Multi-Race'"
    rs.Open mswphdmultiracemale_query,conn 
    
                    d= Replace(rs("multiracecountmale"),"|",",")
                    e = Replace(rs("percentmultiracemales"),"|",",")
    
    Else
    mswphdmultiracemale_query = "SELECT Count (distinct UIN) multiracecountmale FROM PHDApplicants where Degree_Program = 'PHD'  and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswphdmultiracemale_query,conn 
    
                    d= Replace(rs("multiracecountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswphdtotalna <> 0 Then
    mswphdmultiracena_query = "SELECT Count (distinct UIN) multiracecountna, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalna&" as float), 2) percentmultiracena FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender not in ('F','M') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswphdmultiracena_query,conn 
    
                    f = Replace(rs("multiracecountna"),"|",",")
                    g = Replace(rs("percentmultiracena"),"|",",")



    Else
    mswphdmultiracena_query = "SELECT Count (distinct UIN) multiracecountna FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswphdmultiracena_query,conn 
    
                    f = Replace(rs("multiracecountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswphdTotal <> 0 Then
    mswphdmultirace_query = "SELECT Count (distinct UIN) multiracecount, round(Count(distinct UIN)* 100 /Cast("&mswphdTotal&" as float), 2) percentmultirace FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Multi-Race'"
    rs.Open mswphdmultirace_query,conn 
    
                    h = Replace(rs("multiracecount") ,"|",",")
                    i = Replace(rs("percentmultirace"),"|",",")
    
    Else
    mswphdmultirace_query = "SELECT Count (distinct UIN) multiracecount FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Multi-Race'"
    rs.Open mswphdmultirace_query,conn 
    
                    h = Replace(rs("multiracecount") ,"|",",")
                    i = Replace("0","|",",")
   End If
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                
    
    '//Total Minority//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswphdtotalfemales <> 0 Then
					mswphdminorityfemale_query="SELECT Count (distinct UIN) minoritycountfemale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalfemales&" as float), 2) percentminorityfemales FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' and Race_ethinicity != 'White'"
					rs.Open mswphdminorityfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total Minority","|",",")
                    b = Replace(rs("minoritycountfemale"),"|",",")
                    c = Replace(rs("percentminorityfemales"),"|",",")
    End If
    Else
    mswphdminorityfemale_query="SELECT Count (distinct UIN) minoritycountfemale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
					rs.Open mswphdminorityfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total Minority","|",",")
                    b = Replace(rs("minoritycountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If 
    rs.close

    If mswphdtotalmales <> 0 Then
    mswphdminoritymale_query = "SELECT Count (distinct UIN) minoritycountmale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalmales&" as float), 2) percentminoritymales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity != 'White'"
    rs.Open mswphdminoritymale_query,conn 
    
                    d= Replace(rs("minoritycountmale"),"|",",")
                    e = Replace(rs("percentminoritymales"),"|",",")
     
    Else
     mswphdminoritymale_query = "SELECT Count (distinct UIN) minoritycountmale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswphdminoritymale_query,conn 
    
                    d= Replace(rs("minoritycountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    
    rs.close
    
    If mswphdtotalna <> 0 Then
    mswphdminorityna_query = "SELECT Count (distinct UIN) minoritycountna, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalna&" as float), 2) percentminorityna FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M') and Race_ethinicity != 'White'"
    rs.Open mswphdminorityna_query,conn 
    
                    f = Replace(rs("minoritycountna"),"|",",")
                    g = Replace(rs("percentminorityna"),"|",",")



    Else
    mswphdminorityna_query = "SELECT Count (distinct UIN) minoritycountna FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswphdminorityna_query,conn 
    
                    f = Replace(rs("minoritycountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswphdTotal <> 0 Then
    mswphdMinority_query = "SELECT Count (distinct UIN) minoritycount, round(Count(distinct UIN)* 100 /Cast("&mswphdTotal&" as float), 2) percentminority FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity != 'White'"
    rs.Open mswphdMinority_query,conn 
    
                    h = Replace(rs("minoritycount") ,"|",",")
                    i = Replace(rs("percentminority"),"|",",")
    
    Else
    mswphdMinority_query = "SELECT Count (distinct UIN) minoritycount FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity != 'White'"
    rs.Open mswphdMinority_query,conn 
    
                    h = Replace(rs("minoritycount") ,"|",",")
                    i = Replace("0","|",",")
    End If


   rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
                 

     '//Caucasian//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    
    If mswphdtotalfemales <> 0 Then
					mswphdcaucasianfemale_query="SELECT Count (distinct UIN) caucasiancountfemale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalfemales&" as float), 2) percentcaucasianfemales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'White'"
					rs.Open mswphdcaucasianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Caucasian","|",",")
                    b = Replace(rs("caucasiancountfemale"),"|",",")
                    c = Replace(rs("percentcaucasianfemales"),"|",",")
    End If
    Else
    mswphdcaucasianfemale_query="SELECT Count (distinct UIN) caucasiancountfemale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
					rs.Open mswphdcaucasianfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Caucasian","|",",")
                    b = Replace(rs("caucasiancountfemale"),"|",",")
                    c = Replace("0","|",",")
    End If
    End If

     rs.close

    If mswphdtotalmales <> 0 Then
    mswphdcaucasianmale_query = "SELECT Count (distinct UIN) caucasiancountmale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalmales&" as float), 2) percentcaucasianmales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'White'"
    rs.Open mswphdcaucasianmale_query,conn 
    
                    d= Replace(rs("caucasiancountmale"),"|",",")
                    e = Replace(rs("percentcaucasianmales"),"|",",")
    
    Else
    mswphdcaucasianmale_query = "SELECT Count (distinct UIN) caucasiancountmale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswphdcaucasianmale_query,conn 
    
                    d= Replace(rs("caucasiancountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
    
    If mswphdtotalna <> 0 Then
    mswphdcaucasianna_query = "SELECT Count (distinct UIN) caucasiancountna, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalna&" as float), 2) percentcaucasianna FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M') and Race_ethinicity = 'White'"
    rs.Open mswphdcaucasianna_query,conn 
    
                    f = Replace(rs("caucasiancountna"),"|",",")
                    g = Replace(rs("percentcaucasianna"),"|",",")



    Else
    mswphdcaucasianna_query = "SELECT Count (distinct UIN) caucasiancountna FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswphdcaucasianna_query,conn 
    
                    f = Replace(rs("caucasiancountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswphdTotal <> 0 Then
    mswphdCaucasian_query = "SELECT Count (distinct UIN) caucasiancount, round(Count(distinct UIN)* 100 /Cast("&mswphdTotal&" as float), 2) percentcaucasian FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'White'"
    rs.Open mswphdCaucasian_query,conn 
    
                    h = Replace(rs("caucasiancount") ,"|",",")
                    i = Replace(rs("percentcaucasian"),"|",",")
    
    Else
    mswphdCaucasian_query = "SELECT Count (distinct UIN) caucasiancount FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'White'"
    rs.Open mswphdCaucasian_query,conn 
    
                    h = Replace(rs("caucasiancount") ,"|",",")
                    i = Replace("0","|",",")
   End If
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   


    '//Unknown//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswphdtotalfemales <> 0 Then
					mswphdUnknownfemale_query="SELECT Count (distinct UIN) unknowncountfemale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalfemales&" as float), 2) percentunknownfemales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' and Race_ethinicity = 'Unknown'"
					rs.Open mswphdUnknownfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Undeclared","|",",")
                    b = Replace(rs("unknowncountfemale"),"|",",")
                    c = Replace(rs("percentunknownfemales"),"|",",")
    End If
    Else
    mswphdUnknownfemale_query="SELECT Count (distinct UIN) unknowncountfemale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'F'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
					rs.Open mswphdUnknownfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Undeclared","|",",")
                    b = Replace(rs("unknowncountfemale"),"|",",")
                    c = Replace("0","|",",")

    End If
    End If
     rs.close

    If mswphdtotalmales <> 0 Then
    mswphdUnknownmale_query = "SELECT Count (distinct UIN) unknowncountmale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalmales&" as float), 2) percentunknownmales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' and Race_ethinicity = 'Unknown'"
    rs.Open mswphdUnknownmale_query,conn 
    
                    d= Replace(rs("unknowncountmale"),"|",",")
                    e = Replace(rs("percentunknownmales"),"|",",")
    
    Else
     mswphdUnknownmale_query = "SELECT Count (distinct UIN) unknowncountmale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender = 'M'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswphdUnknownmale_query,conn 
    
                    d= Replace(rs("unknowncountmale"),"|",",")
                    e = Replace("0","|",",")
    
    End If
    rs.close
    
    If mswphdtotalna <> 0 Then
    mswphdUnknownna_query = "SELECT Count (distinct UIN) unknowncountna, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalna&" as float), 2) percentunknownna FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M') and Race_ethinicity = 'Unknown'"
    rs.Open mswphdUnknownna_query,conn 
    
                    f = Replace(rs("unknowncountna"),"|",",")
                    g = Replace(rs("percentunknownna"),"|",",")



    Else
    mswphdUnknownna_query = "SELECT Count (distinct UIN) unknowncountna FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"' and Gender not in ('F','M')  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswphdUnknownna_query,conn 
    
                    f = Replace(rs("unknowncountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswphdTotal <> 0 Then
    mswphdUnknown_query = "SELECT Count (distinct UIN) unknowncount, round(Count(distinct UIN)* 100 /Cast("&mswphdTotal&" as float), 2) percentunknown FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Race_ethinicity = 'Unknown'"
    rs.Open mswphdUnknown_query,conn 
    
                    h = Replace(rs("unknowncount") ,"|",",")
                    i = Replace(rs("percentunknown"),"|",",")
    
    Else
    mswphdUnknown_query = "SELECT Count (distinct UIN) unknowncount FROM PHDApplicants where Degree_Program = 'PHD'  and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Race_ethinicity = 'Unknown'"
    rs.Open mswphdUnknown_query,conn 
    
                    h = Replace(rs("unknowncount") ,"|",",")
                    i = Replace("0","|",",")
   End If
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
                   
     '//Total//
   
                    set rs=Server.CreateObject("ADODB.recordset")
    If mswphdtotalfemales <> 0 Then
					mswphdTotalfemale_query="SELECT Count (distinct UIN) totalcountfemale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalfemales&" as float), 2) percenttotalfemales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'F' "
					rs.Open mswphdTotalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total","|",",")
                    b = Replace(rs("totalcountfemale"),"|",",")
                    c = Replace(rs("percenttotalfemales"),"|",",")
    End If
    Else
    mswphdTotalfemale_query="SELECT Count (distinct UIN) totalcountfemale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'F' "
					rs.Open mswphdTotalfemale_query,conn 

      If rs.EOF Then
                      
                    Else

                    a= Replace("Total","|",",")
                    b = Replace(rs("totalcountfemale"),"|",",")
                    c = Replace("0","|",",") 
    End If
    End If
    rs.close

    If mswphdtotalmales <> 0 Then
    mswphdTotalmale_query = "SELECT Count (distinct UIN) totalcountmale, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalmales&" as float), 2) percenttotalmales FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' and Gender = 'M' "
    rs.Open mswphdTotalmale_query,conn 
    
                    d= Replace(rs("totalcountmale"),"|",",")
                    e = Replace(rs("percenttotalmales"),"|",",")
    
    Else
    mswphdTotalmale_query = "SELECT Count (distinct UIN) totalcountmale FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender = 'M' "
    rs.Open mswphdTotalmale_query,conn 
    
                    d= Replace(rs("totalcountmale"),"|",",")
                    e = Replace("0","|",",")
    End If
    rs.close
     
    If mswphdtotalna <> 0 Then
    mswphdTotalna_query = "SELECT Count (distinct UIN) totalcountna, round(Count(distinct UIN)* 100 /Cast("&mswphdtotalna&" as float), 2) percenttotalna FROM PHDApplicants where Degree_Program = 'PHD' and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender not in ('F','M')"
    rs.Open mswphdTotalna_query,conn 
    
                    f = Replace(rs("totalcountna"),"|",",")
                    g = Replace(rs("percenttotalna"),"|",",")



    Else
    mswphdTotalna_query = "SELECT Count (distinct UIN) totalcountna FROM PHDApplicants where Degree_Program = 'PHD'  and Term_CD='"&AdmitTerm&"'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Gender not in ('F','M')"
    rs.Open mswphdTotalna_query,conn 
    
                    f = Replace(rs("totalcountna"),"|",",")
                    g = Replace("0","|",",")

    End If
    rs.close

    If mswphdTotal <> 0 Then
    mswphdTotal_query = "SELECT Count (distinct UIN) totalcount, round(Count(distinct UIN)* 100 /Cast("&mswphdTotal&" as float), 2) percenttotal FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' "
    rs.Open mswphdTotal_query,conn 
    
                    h = Replace(rs("totalcount") ,"|",",")
                    i = Replace(rs("percenttotal"),"|",",")
    
    Else

    mswphdTotal_query = "SELECT Count (distinct UIN) totalcount FROM PHDApplicants where Degree_Program = 'PHD'  and Admission_decision IN ('A', 'S', 'ReAdmit') and Term_CD='"&AdmitTerm&"' "
    rs.Open mswphdTotal_query,conn 
    
                    h = Replace(rs("totalcount") ,"|",",")
                    i = Replace("0","|",",")
   End If
    rs.close
                    
                    pdf.Row a,b,c,d,e,f,g,h,i
    

    pdf.Ln(10)
pdf.ChapterTitle("                                                                         Thank you !")
pdf.Ln(5)

pdf.Close()
pdf.Output()
conn.close



    %>