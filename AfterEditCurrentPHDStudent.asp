<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
UIN = Request("UIN")
LastName = Replace(Request("lastname"),"'","''")
FirstName = Replace(Request("FirstName"),"'","''")
MiddleName = Replace(Request("middlename"),"'","''")
DateOfBirth = Request("dob")

Salutation = Request("Salutation")
MaidenName = Replace(Request("maidenname"),"'","''")
Gender = Request("gender")
CurrentAddress1 = Request("mailingAddress1")
CurrentAddress2 = Request("mailingAddress2")
CurrentCity = Request("mailingcity")
CurrentState = Request("mailingstate")
CurrentZipCode = Request("mailingzipcode")
CurrentCountry = Request("country")
HomePhone = Request("homephone")
SO_Name = Request("SO_Name")
Fax = Request("fax")
CellPhone = Request("cellphone")
InternationalPhone = Request("internationalphonenumber")
EMail = Request("email")
UGCollege = Request("ugcollege")
Race_ethinicity = Request("Race_ethinicity")
Race_desc=Request("Race_desc")
UGGPA=Request("uggpa")
UGMajor=Request("ugmajor")
GradCollege=Request("gradcollege")
GradGPA=Request("gradgpa")
GradMajor=Request("gradmajor")
GradDegree = Request("graddegree")
DateofDefense = Request("DateofDefense")
DateofPreliminaryExam = Request("DateofPreliminaryExam")
ProgramType = Request("ProgramType")
DateofComprehensiveExam = Request("DateofComprehensiveExam")
ReasonforRefusion = Request("ReasonforRefusion")
ApplyingForGraduation = Request("ApplyingForGraduation")
GraduationTermAppliedFor = Request("GraduationTermAppliedFor")
AdmitTerm = Request("AdmitTerm")
advisor = Request("advisor")
TermGraduated = Request("TermGraduated")
EnteredBy = Request("EnteredBy")
DateEntered = Request("DateEntered")
LastUpdatedBy = Request("LastUpdatedBy")



update_query="update CurrentPHDStudents set UIN='"&UIN&"',LastName = '"&LastName&"', FirstName = '"&FirstName&"', MiddleName = '"&MiddleName&"',Salutation='"&Salutation&"',MaidenName='"&MaidenName&"',Gender='"&Gender&"',DateOfBirth='"&DateOfBirth&"',Race_ethinicity='"&Race_ethinicity&"',Race_desc= '"&Race_Desc&"',SO_Name='"&SO_Name&"',MailingAddress1='"&CurrentAddress1&"',MailingAddress2='"&CurrentAddress2&"',MailingCity='"&CurrentCity&"',MailingState='"&CurrentState&"',MailingZipCode='"&CurrentZipCode&"',Country='"&CurrentCountry&"',HomePhone='"&HomePhone&"',Fax='"&Fax&"', CellPhone='"&CellPhone&"',InternationalPhoneNumber='"&InternationalPhoneNumber&"',EMail='"&EMail&"', UGCollege= '"&UGCollege&"',UGGPA= '"&UGGPA&"' ,UGMajor='"&UGMajor&"',GradCollege='"&GradCollege&"',GradGPA='"&GradGPA&"',Type='"&ProgramType&"',GradMajor='"&GradMajor&"',GradDegree='"&GradDegree&"',DateofDefense='"&DateofDefense&"',DateofPreliminaryExam='"&DateofPreliminaryExam&"',DateofComprehensiveExam = '"&DateofComprehensiveExam&"', AdmitTerm='"&AdmitTerm&"',Advisor='"&advisor&"',ReasonforRefusion='"&ReasonforRefusion&"',ApplyingForGraduation='"&ApplyingForGraduation&"',GraduationTermAppliedFor='"&GraduationTermAppliedFor&"',TermGraduated='"&TermGraduated&"',EnteredBy='"&EnteredBy&"', DateEntered= '"&DateEntered&"', LastUpdatedBy= '"&LastUpdatedBy&"' where UIN='"&UIN&"'"	
   
  
'set rs=Server.CreateObject("ADODB.recordset")
'isadvisor = "select distinct Advisor_Name from Adviser where Student_UIN ='"&UIN&"'"
'rs.Open isadvisor,conn
'if rs.eof = true then 
'response.write(advisor)
 '    if rs("Advisor_Name") is null then

  '  else
'response.write(StrComp(advisor,rs("Advisor_Name")))
'If ((StrComp(advisor,rs("Advisor_Name")))= 0) Then
  
'conn.Execute"insert into Adviser (Student_UIN, Advisor_Name, Student_FirstName, Student_LastName, Student_MiddleName) values ('"&UIN&"','"&advisor&"','"&FirstName&"', '"&LastName&"', '"&MiddleName&"')"
'End If 
'End If
 '   End If
'rs.Close
      
    Set objRS = Server.CreateObject("ADODB.recordset") 
   objRS.open update_query,conn
conn.close
conn1.close

        Response.Redirect "EditPHDCurrentStudent.asp?UIN="&UIN&"&ErrMsg='Student Information Updated to DB'"
%>
