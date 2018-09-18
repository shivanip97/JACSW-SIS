<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
UIN = Request("uin")

LastName = Request("lastname")
FirstName = Request("firstname")
MiddleName = Request("middlename")
MaidenName = Request("maidenname")
Gender = Request("gender")
OARdate= Request("oar_application_date")
SO_Name = Request("SO_Name")
Salutation=Request("Salutation")

application_status = Request("application_status")
Degree_Program = Request("Degree_Program")
fax = Request("fax")
InternationalPhoneNumber = Request("InternationalPhoneNumber")

admission_decision = Request("admission_decision")
decision_dt = Request("decision_dt")
Decision_Letter_Sent_Date = Request("Decision_Letter_Sent_Date")
Confirmed = Request("Confirmed")
Confirmed_Dt = Request("Confirmed_Dt")
Admit_Term = Request("Admit_Term")
Financial_Aid_Requested = Request("Financial_Aid_Requested")
LastUpdatedDt = Date()


RaceEthinicity = Request("Race_ethinicity")
RaceDesc = Request("Race_desc")
CurrentAddress1 = Request("currentAddress1")
CurrentAddress2 = Request("currentAddress2")
CurrentCity = Request("currentcity")
CurrentState = Request("currentstate")

CurrentZipCode = Request("currentzip")
CurrentCountry = Request("currentcountry")
HomePhone = Request("homephone")
WorkPhone = Request("workphone")
EMail = Request("email")
DateOfBirth = Request("dob")
Comments = Request("comments")
UGCollege = Request("ugcollege")
UGGPA = Request("uggpa")
UGMajor = Request("ugmajor")
GradCollege = Request("gradcollege")
GradMajor = Request("gradmajor")
GradGPA = Request("gradgpa")
GradDegree = Request("graddegree")
DateofInitialEntry=Request("DateofInitialEntry")
Reapplicant = Request("Reapplicant")
Orientation = Request("Orientation")
Registered = Request("Registered")
UIC_employee = Request("UIC_employee")
Open_house = Request("Open_house")
UIC_UG_Grad_Apps = Request("UIC_UG_Grad_Apps")
Application_Fee = Request("Application_Fee")
Jane_Addams_appln = Request("Jane_Addams_appln")
Transcripts = Request("Transcripts")
TOEFL_Score = Request("TOEFL_Score")
GRE_Quantitative = Request("GRE_Quantitative")
GRE_Verbal = Request("GRE_Verbal")
GRE_Analytical = Request("GRE_Analytical")
Field_of_Interest = Request("Field_of_Interest")
Dec_Cert_Finances_Sub = Request("Dec_Cert_Finances_Sub")

Citizenship_Status = Request("Citizenship_Status")
Country_of_Citizenship = Request("Country_of_Citizenship")


update_query="update PHDApplicants set UIN='"&UIN&"',LastName='"&LastName&"',Salutation='"&Salutation&"',Race_Desc='"&RaceDesc&"',FirstName='"&FirstName&"',MiddleName='"&MiddleName&"',MaidenName='"&MaidenName&"',Gender='"&Gender&"',OAR_Application_Date='"&OARdate&"', SO_Name='"&SO_Name&"',EnteredBy='"& Session("Username") &"',application_status='"&application_status&"',Degree_Program='"&Degree_Program&"',fax='"&fax&"',InternationalPhoneNumber='"&InternationalPhoneNumber&"',admission_decision='"&admission_decision&"',decision_dt='"&decision_dt&"',Decision_Letter_Sent_Date='"&Decision_Letter_Sent_Date&"',Reapplicant='"&Reapplicant&"',Confirmed='"&Confirmed&"',Confirmed_Dt='"&Confirmed_Dt&"',Admit_Term='"&Admit_Term&"',Financial_Aid_Requested='"&Financial_Aid_Requested&"',LastUpdatedDt='"&LastUpdatedDt&"', Race_ethinicity='"&RaceEthinicity&"',CurrentAddress1='"&CurrentAddress1&"',CurrentAddress2='"&CurrentAddress2&"',CurrentCity='"&CurrentCity&"',CurrentState='"&CurrentState&"',CurrentZipCode='"&CurrentZipCode&"',CurrentCountry='"&CurrentCountry&"',HomePhone='"&HomePhone&"',WorkPhone='"&WorkPhone&"',EMail='"&EMail&"',DateOfBirth='"&DateOfBirth&"',Comments='"&Comments&"',UGCollege='"&UGCollege&"',UGGPA='"&UGGPA&"',UGMajor='"&UGMajor&"',GradCollege='"&GradCollege&"',GradMajor='"&GradMajor&"',GradGPA='"&GradGPA&"',GradDegree='"&GradDegree&"',DateofInitialEntry='"&DateofInitialEntry&"',Orientation='"&Orientation&"',Registered='"&Registered&"',UIC_employee='"&UIC_employee&"',Open_house='"&Open_house&"',UIC_UG_Grad_Apps='"&UIC_UG_Grad_Apps&"',Application_Fee='"&Application_Fee&"',Jane_Addams_appln='"&Jane_Addams_appln&"',Transcripts='"&Transcripts&"',TOEFL_Score='"&TOEFL_Score&"',GRE_Quantitative='"&GRE_Quantitative&"',GRE_Verbal='"&GRE_Verbal&"',GRE_Analytical='"&GRE_Analytical&"',Field_of_Interest='"&Field_of_Interest&"',Dec_Cert_Finances_Sub='"&Dec_Cert_Finances_Sub&"',Country_of_Citizenship='"&Country_of_Citizenship&"',Citizenship_Status='"&Citizenship_Status&"' where UIN='"&UIN&"' and Admit_Term='"&Admit_Term&"'"


set rs=Server.CreateObject("ADODB.recordset")
isstudent = "select distinct uin from CurrentPHDStudents where UIN ='"&UIN&"'"
rs.Open isstudent,conn
if rs.eof = true then 
response.write(Confirmed)
response.write(StrComp(Confirmed,"Y"))
If ((StrComp(Confirmed,"Y"))= 0) Then
  
conn.Execute"insert into CurrentPHDStudents (UIN, LastName, FirstName, MiddleName, MaidenName, AdmitTerm, Gender, DateOfBirth) values ('"&UIN&"','"&LastName&"','"&FirstName&"','"&MiddleName&"','"&MaidenName&"','"&Admit_Term&"', '"&Gender&"', '"&DateOfBirth&"')"
End If 
End If
rs.Close
   Set objRS = Server.CreateObject("ADODB.recordset") 
   objRS.open update_query, conn
conn.close
conn1.close

        Response.Redirect "EditPHDStudent.asp?UIN="&UIN&"&ErrMsg='Student Information Updated to DB'"
%>
