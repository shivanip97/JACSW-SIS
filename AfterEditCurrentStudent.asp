<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
UIN = Request("UIN")
LastName = Replace(Request("lastname"),"'","''")
FirstName = Replace(Request("firstname"),"'","''")
MiddleName = Replace(Request("middlename"),"'","''")
DateOfBirth = Request("dob")
DegreeProgram = Request("DegreeProgram")
Salutation = Request("Salutation")
PreferredFirstName = Replace(Request("maidenname"),"'","''")
Gender = Request("gender")
CurrentAddress1 = Replace(Request("currentAddress1"),"'","''")
CurrentAddress2 = Replace(Request("currentAddress2"),"'","''")
CurrentCity = Request("currentcity")
CurrentState = Request("currentstate")
CurrentZipCode = Request("currentzipcode")
CurrentCountry = Request("currentcountry")
HomePhone = Request("homephone")
WorkPhone = Request("workphone")
CellPhone = Request("cellphone")
InternationalPhone = Request("internationalphonenumber")
EMail = Request("email")
Personalemail = Request("Personalemail")
RaceEthinicity = Request("Race_ethinicity")
Race_SubCategory=Request("Race_SubCategory")
ProbationStartTerm=Request("ProbationStartTerm")
    IBHE_Certificate=Replace(Request("IBHE_Certificate"),"'","''")
    Certificate_StartTerm = Request("Certificate_StartTerm")
     ChildWelfareTraineeshipProject=Replace(Request("ChildWelfareTraineeshipProject"),"'","''")
    ChildWelfareTraineeshipProjectStartTerm=Request("ChildWelfareTraineeshipProjectStartTerm")
ProbationEndTerm=Request("ProbationEndTerm")
LeaveofAbsenceStartTerm=Request("LeaveofAbsenceStartTerm")
LeaveofAbsenceEndTerm=Request("LeaveofAbsenceEndTerm")
ForwardtoField=Request("ForwardtoField")
    Withdrawn =Request("Withdrawn")
    WithdrawalReason =Request("WithdrawalReason")
    WithdrawnDate=Request("WithdrawnDate")
Field_Type = Request("Field_Type")
LimitedStatus = Request("LimitedStatus")
Comments = Request("comments")
ProgramType = Request("ProgramType")
Concentration = Request("Concentration")
Decision = Request("Decision")
Confirmed = Request("Confirmed")
     Status = Request("Status")
ConfirmedDate = Request("ConfirmedDate")
Admit_Term = Request("Admit_Term")
advisor = Replace(Request("advisor"),"'","''")
Track = Request("Track")
CurrentYear = Request("CurrentYear")
ApplyingForGraduation = Request("ApplyingForGraduation")
GraduationTermAppliedFor = Request("GraduationTermAppliedFor")
TermGraduated = Request("TermGraduated")
DegreeApplyingFor = Request("DegreeApplyingFor")
    Graduated = Request("Graduated")
    GraduatedDate = Request("GraduatedDate")
MailboxNumber = Request("MailboxNumber")
ModifiedPlan = Request("ModifiedPlan")
    WithdrawnTerm = Request("WithdrawnTerm")
    ConfirmedDueDate =Request("ConfirmedDueDate")

update_query="update CurrentStudents set BannerId='"&UIN&"', ModifiedPlan = '"&ModifiedPlan&"',WithdrawnTerm = '"&WithdrawnTerm&"',WithdrawalReason = '"&WithdrawalReason&"', WithdrawnDate = '"&WithdrawnDate&"',UIN = '"&UIN&"',GraduatedDate = '"&GraduatedDate&"', Status = '"&Status&"', Graduated = '"&Graduated&"',Field_Type = '"&Field_Type&"', IBHE_Certificate = '"&IBHE_Certificate&"',Certificate_StartTerm = '"&Certificate_StartTerm&"',ChildWelfareTraineeshipProject = '"&ChildWelfareTraineeshipProject&"',ChildWelfareTraineeshipProjectStartTerm = '"&ChildWelfareTraineeshipProjectStartTerm&"', ProbationStartTerm = '"&ProbationStartTerm&"', ProbationEndTerm = '"&ProbationEndTerm&"',LastName='"&LastName&"',FirstName='"&FirstName&"',MiddleName='"&MiddleName&"',PreferredFirstName='"&PreferredFirstName&"',Gender='"&Gender&"',DegreeProgram= '"&DegreeProgram&"',salutation='"&Salutation&"',CurrentAddress1='"&CurrentAddress1&"',CurrentAddress2='"&CurrentAddress2&"',CurrentCity='"&CurrentCity&"',CurrentState='"&CurrentState&"',CurrentZipCode='"&CurrentZipCode&"',CurrentCountry='"&CurrentCountry&"',HomePhone='"&HomePhone&"',WorkPhone='"&WorkPhone&"', CellPhone='"&CellPhone&"',InternationalPhoneNumber='"&InternationalPhoneNumber&"',EMail='"&EMail&"', Personalemail= '"&Personalemail&"',ForwardtoField= '"&ForwardtoField&"' ,DateOfBirth='"&DateOfBirth&"',Comments='"&Comments&"',LimitedStatus='"&LimitedStatus&"',ProgramType='"&ProgramType&"',Concentration='"&Concentration&"',Decision='"&Decision&"',Confirmed='"&Confirmed&"',ConfirmedDate='"&ConfirmedDate&"',AdmitTerm='"&Admit_Term&"',Advisor='"&advisor&"',Track='"&Track&"',CurrentYear='"&CurrentYear&"',ApplyingForGraduation='"&ApplyingForGraduation&"',GraduationTermAppliedFor='"&GraduationTermAppliedFor&"',TermGraduated='"&TermGraduated&"',DegreeApplyingFor='"&DegreeApplyingFor&"', LeaveofAbsenceStartTerm= '"&LeaveofAbsenceStartTerm&"', LeaveofAbsenceEndTerm= '"&LeaveofAbsenceEndTerm&"',Race_ethinicity= '"&RaceEthinicity&"', Race_SubCategory= '"&Race_SubCategory&"', Withdrawn='"&Withdrawn&"', ConfirmedDueDate='"&ConfirmedDueDate&"' where UIN='"&UIN&"'"	
Response.write(update_query)
     
set rs=Server.CreateObject("ADODB.recordset")
isstudent = "select distinct uin from Field1 where UIN ='"&UIN&"'"
 rs.Open isstudent,conn 
if rs.eof = true then 
response.write(ForwardtoField)
response.write(StrComp(ForwardtoField,"Y"))
If ((StrComp(ForwardtoField,"Y"))= 0) Then
  
conn.Execute"insert into Field1 (UIN, FieldTypeYear) values ('"&UIN&"','"&CurrentYear&"')"
End If 
End If
rs.Close
  
set rs=Server.CreateObject("ADODB.recordset")
isadvisor = "select distinct Advisor_Name from Adviser where Student_UIN ='"&UIN&"'"
rs.Open isadvisor,conn
if rs.eof = true then 
response.write(advisor)  
conn.Execute"insert into Adviser (Student_UIN, Advisor_Name, Student_FirstName, Student_LastName, Student_MiddleName) values ('"&UIN&"','"&advisor&"','"&FirstName&"', '"&LastName&"', '"&MiddleName&"')"
    End If
rs.Close
      
    Set objRS = Server.CreateObject("ADODB.recordset") 
   objRS.open update_query,conn
conn.close
conn1.close

        Response.Redirect "EditCurrentStudent.asp?UIN="&UIN&"&ErrMsg='Student Information Updated to DB'"
%>
