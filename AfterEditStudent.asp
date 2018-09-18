<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
UIN = Request("uin")
LastName = Replace(Request("lastname"), "'","''")


FirstName = Replace(Request("firstname"),"'","''")
MiddleName = Replace(Request("middlename"), "'","''")
MaidenName = Replace(Request("maidenname"), "'","''")
Gender = Request("gender")
OARdate= Request("oar_application_date")
ReadyforReviewDate = Request("readyforreviewdate")


application_status = Request("application_status")
Degree_Program = Request("Degree_Program")
admitted_to_school = Request("admitted_to_school")
credit_in_ba_bs = Request("credit_in_ba_bs")
credit_in_english = Request("credit_in_english")
requesting_schools = Request("requesting_schools")
reapplicant = Request("reapplicant")
program_type = Request("program_type")
concentration = Request("Concentration")
admission_decision = Request("admission_decision")
decision_dt = Request("decision_dt")
Decision_Letter_Sent_Date = Request("Decision_Letter_Sent_Date")
Limited_status = Request("Limited_status")
Confirmed = Request("Confirmed")
Confirmed_Dt = Request("Confirmed_Dt")
Admit_Term = Request("Admit_Term")
Credit_in_Statistics = Request("Credit_in_Statistics")
Financial_Aid_Request = Request("Financial_Aid_Request")
Basic_Skill_Test = Request("Basic_Skill_Test")
ACT_SAT = Request("ACT_SAT")
Passed_Test = Request("Passed_Test")
DateCommentsEntered = Request("DateCommentsEntered")
ConfirmedDueDate = Request("ConfirmedDueDate")
LastUpdatedDt = Date()


RaceEthinicity = Request("Race_ethinicity")
CurrentAddress1 = Replace(Request("currentAddress1"), "'","''")
CurrentAddress2 = Replace(Request("currentAddress2"), "'","''")
CurrentCity = Request("currentcity")
CurrentState = Request("currentstate")

CurrentZipCode = Request("currentzip")
CurrentCountry = Request("currentcountry")
HomePhone = Request("homephone")
WorkPhone = Request("workphone")
EMail = Request("email")
DateOfBirth = Request("dob")
Comments = Request("comments")
UGCollege = Replace(Request("ugcollege"),"'","''")
UGGPA = Request("uggpa")
UGMajor = Replace(Request("ugmajor"),"'","''")
GradCollege = Replace(Request("gradcollege"),"'","''")
GradMajor = Replace(Request("gradmajor"),"'","''")
GradGPA = Request("gradgpa")
GradDegree = Request("graddegree")
Award_Type=Request("award_type1")
Award_Amount=Request("award_amount1")
Award_Date=Request("award_date1")
    Award_Type2=Request("award_type2")
Award_Amount2=Request("award_amount2")
Award_Date2=Request("award_date2")
    Award_Type3=Request("award_type3")
Award_Amount3=Request("award_amount3")
Award_Date3=Request("award_date3")
Race_SubCategory=Request("race_subcategory")
Withdrawn=Request("withdrawn")
Withdraw_Reason=Request("withdraw_reason")
WithdrawnDate = Request("WithdrawnDate")
Field_Type=Request("Field_Type")  
Received_Deposit=Request("received_deposit")
Forward_to_Field=Request("forward_to_field")
DeferredFrom=Request("DeferredFrom")
DeferredTo=Request("DeferredTo")
Adv_Verification=Request("Adv_Verification")
International=Request("International")

update_query="update Applicants set UIN='"&UIN&"',WithdrawnDate='"&WithdrawnDate&"',Adv_Verification='"&Adv_Verification&"',International='"&International&"', DeferredFrom='"&DeferredFrom&"',DeferredTo='"&DeferredTo&"',LastName='"&LastName&"',FirstName='"&FirstName&"',MiddleName='"&MiddleName&"',MaidenName='"&MaidenName&"',Gender='"&Gender&"',OAR_Application_Date='"&OARdate&"', ReadyforReviewDate='"&ReadyforReviewDate&"',EnteredBy='"& Session("Username") &"',application_status='"&application_status&"',Degree_Program='"&Degree_Program&"',admitted_to_school='"&admitted_to_school&"',credit_in_ba_bs='"&credit_in_ba_bs&"',credit_in_english='"&credit_in_english&"',requesting_schools='"&requesting_schools&"',reapplicant='"&reapplicant&"',program_type='"&program_type&"',concentration='"&concentration&"',admission_decision='"&admission_decision&"',decision_dt='"&decision_dt&"',Decision_Letter_Sent_Date='"&Decision_Letter_Sent_Date&"',Limited_status='"&Limited_status&"',Confirmed='"&Confirmed&"',Confirmed_Dt='"&Confirmed_Dt&"',ConfirmedDueDate='"&ConfirmedDueDate&"',Admit_Term='"&Admit_Term&"',Credit_in_Statistics='"&Credit_in_Statistics&"',Financial_Aid_Request='"&Financial_Aid_Request&"',Basic_Skill_Test='"&Basic_Skill_Test&"',ACT_SAT='"&ACT_SAT&"',Passed_Test='"&Passed_Test&"',DateCommentsEntered='"&DateCommentsEntered&"',LastUpdatedDt='"&LastUpdatedDt&"', Race_ethinicity='"&RaceEthinicity&"',CurrentAddress1='"&CurrentAddress1&"',CurrentAddress2='"&CurrentAddress2&"',CurrentCity='"&CurrentCity&"',CurrentState='"&CurrentState&"',CurrentZipCode='"&CurrentZipCode&"',CurrentCountry='"&CurrentCountry&"',HomePhone='"&HomePhone&"',WorkPhone='"&WorkPhone&"',InternationalPhoneNumber='"&InternationalPhoneNumber&"',EMail='"&EMail&"',DateOfBirth='"&DateOfBirth&"',Comments='"&Comments&"',UGCollege='"&UGCollege&"',UGGPA='"&UGGPA&"',UGMajor='"&UGMajor&"',GradCollege='"&GradCollege&"',GradMajor='"&GradMajor&"',GradGPA='"&GradGPA&"',GradDegree='"&GradDegree&"',Award_Type='"&Award_Type&"' ,Award_Amount='"&Award_Amount&"',Award_Date='"&Award_Date&"',Award_Type2='"&Award_Type2&"' ,Award_Amount2='"&Award_Amount2&"',Award_Date2='"&Award_Date2&"',Award_Type3='"&Award_Type3&"' ,Award_Amount3='"&Award_Amount3&"',Award_Date3='"&Award_Date3&"',Race_SubCategory='"&Race_SubCategory&"',Withdrawn='"&Withdrawn&"',Withdraw_Reason='"&Withdraw_Reason&"',Field_Type='"&Field_Type&"',forward_to_field='"&Forward_to_Field&"',received_deposit='"&Received_Deposit&"' where UIN='"&UIN&"' and Admit_Term='"&Admit_Term&"' and Degree_Program='"&Degree_Program&"' " 

set rs=Server.CreateObject("ADODB.recordset")
isstudent = "select distinct uin from CurrentStudents where UIN ='"&UIN&"'"
rs.Open isstudent,conn
if rs.eof = true then 
response.write(Confirmed)
response.write(StrComp(Confirmed,"Y"))
If ((StrComp(Confirmed,"Y"))= 0) Then

conn.Execute"insert into Adviser (Student_UIN) values ('"&UIN&"')" 
  
conn.Execute"insert into CurrentStudents (Graduated,UIN, Race_ethinicity, WithdrawnDate, CurrentYear, Race_SubCategory, LastName, FirstName, MiddleName, MaidenName, DegreeProgram, AdmitTerm, Personalemail, ProgramType, DateOfBirth, Gender, CurrentAddress1, CurrentAddress2, CurrentCity, CurrentState, CurrentZipCode, CurrentCountry, HomePhone, WorkPhone, InternationalPhoneNumber, LimitedStatus, Comments, Concentration, Decision, Confirmed, ConfirmedDate, DegreeApplyingFor,ForwardtoField,Advisor, Track,EMail,Withdrawn) values ('""','"&UIN&"','"&RaceEthinicity&"','"&WithdrawnDate&"','','"&Race_SubCategory&"','"&LastName&"','"&FirstName&"','"&MiddleName&"','"&MaidenName&"','"&Degree_Program&"','"&Admit_Term&"', '"&EMail&"', '"&program_type&"', '"&DateOfBirth&"', '"&Gender&"', '"&CurrentAddress1&"', '"&CurrentAddress2&"', '"&CurrentCity&"', '"&CurrentState&"', '"&CurrentZipCode&"', '"&CurrentCountry&"', '"&HomePhone&"', '"&WorkPhone&"', '"&InternationalPhoneNumber&"', '"&Limited_status&"', '"&Comments&"', '"&concentration&"', '"&admission_decision&"', '"&Confirmed&"', '"&Confirmed_Dt&"', '"&Degree_Program&"','"&Forward_to_Field&"','','','','"&Withdrawn&"')"
End If 

    If ((StrComp(Withdrawn,"Y"))= 0) Then
conn.Execute"delete from CurrentStudents where UIN='"&UIN&"' and Withdrawn = 'Y'"
End If 

Else
    If ((StrComp(Confirmed,"N"))= 0) Then
    conn.Execute"delete from CurrentStudents where UIN='"&UIN&"' "
     Else
    update_current = "update CurrentStudents set UIN='"&UIN&"',WithdrawnDate='"&WithdrawnDate&"',LastName='"&LastName&"',FirstName='"&FirstName&"',MiddleName='"&MiddleName&"',MaidenName='"&MaidenName&"',DegreeProgram ='"&Degree_Program&"', AdmitTerm='"&Admit_Term&"', Personalemail='"&EMail&"', ProgramType='"&program_type&"', DateOfBirth='"&DateOfBirth&"', Gender='"&Gender&"',CurrentAddress1= '"&CurrentAddress1&"', CurrentAddress2='"&CurrentAddress2&"', CurrentCity='"&CurrentCity&"', CurrentState='"&CurrentState&"', Race_ethinicity='"&RaceEthinicity&"', Race_SubCategory='"&Race_SubCategory&"', CurrentZipCode='"&CurrentZipCode&"', CurrentCountry='"&CurrentCountry&"', HomePhone='"&HomePhone&"', WorkPhone='"&WorkPhone&"', InternationalPhoneNumber='"&InternationalPhoneNumber&"', LimitedStatus='"&Limited_status&"', Comments='"&Comments&"', Concentration= '"&concentration&"', Decision='"&admission_decision&"', Confirmed='"&Confirmed&"', ConfirmedDate='"&Confirmed_Dt&"', DegreeApplyingFor='"&Degree_Program&"', ForwardtoField='"&Forward_to_Field&"',Withdrawn = '"&Withdrawn&"' where UIN='"&UIN&"' and AdmitTerm='"&Admit_Term&"' and DegreeProgram='"&Degree_Program&"' " 
    Set objRS1 = Server.CreateObject("ADODB.recordset") 
   objRS1.open update_current, conn
    End If
End If

    rs.Close

   Set objRS = Server.CreateObject("ADODB.recordset") 
     Response.Write "strsql1: " & update_query 
   objRS.open update_query, conn
   
  
conn.close
conn1.close

        Response.Redirect "EditStudent.asp?UIN="&UIN&"&Admit_Term="&Admit_Term&"&ErrMsg='Student Information Updated to DB'"
%>
