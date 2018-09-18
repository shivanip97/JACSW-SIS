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

pdf.ChapterTitle2("             Race/Gender Ethinicity Report - Accepted - "&Termsel& "     "&LastUpdatedDt&" "&LastUpdatedTime)
pdf.SetFont "Helvetica","",12

pdf.SetTextColor(000)
'pdf.ChapterBody(userinfo_1)

pdf.Ln(10)

rs.close

   
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