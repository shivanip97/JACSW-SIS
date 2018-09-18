<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
uin = Session("uin")
DegreeProgram = Request("DegreeProgram")
LimitedStatus = Request("LimitedStatus")
Confirmed = Request("Confirmed")
Decision = Request("Decision")
Concentration = Request("Concentration")
ProgramType = Request("ProgramType")
Track = Request("Track")
ConfirmedDate = Request("date")
AdmitTerm = Request("AdmitTerm")
Advisor = Request("Advisor")
CurrentYear = Request("CurrentYear")
ApplyingForGraduation = Request("ApplyingForGraduation")
GraduationTermAppliedFor = Request("GraduationTermAppliedFor")
TermGraduated = Request("TermGraduated")
DegreeApplyingFor = Request("DegreeApplyingFor")
MailboxNumber = Request("MailboxNumber")

set rs=Server.CreateObject("ADODB.recordset")
insert_query="insert into CurrentStudent(UIN,DegreeProgram,LimitedStatus,ProgramType,Concentration,Decision,Confirmed,ConfirmedDate,AdmitTerm,Advisor,Track,CurrentYear,ApplyingForGraduation,GraduationTermAppliedFor,TermGraduated,DegreeApplyingFor,MailboxNumber) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

		Set objCommand = Server.CreateObject("ADODB.Command") 
	                objCommand.ActiveConnection = conn
	                objCommand.CommandText = insert_query
	                objCommand.Parameters(0).value = uin
	                objCommand.Parameters(1).value = DegreeProgram
	                objCommand.Parameters(2).value = LimitedStatus
objCommand.Parameters(3).value = ProgramType
objCommand.Parameters(4).value = Concentration
objCommand.Parameters(5).value = Decision
objCommand.Parameters(6).value = Confirmed
objCommand.Parameters(7).value = ConfirmedDate
objCommand.Parameters(8).value = AdmitTerm
objCommand.Parameters(9).value = Advisor
objCommand.Parameters(10).value = Track
objCommand.Parameters(11).value = CurrentYear
objCommand.Parameters(12).value = ApplyingForGraduation
objCommand.Parameters(13).value = GraduationTermAppliedFor
objCommand.Parameters(14).value = TermGraduated
objCommand.Parameters(15).value = DegreeApplyingFor
objCommand.Parameters(16).value = MailboxNumber

Set objRS = objCommand.Execute()
        Response.Redirect "ShowCurrentStudents.asp?ErrMsg='Record was successfully added to DB'"	

rs.close
conn.close
%>
