<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
AgencyID = Request("AgencyID")


Agency = Replace(Request("Agency"),"'","''")
AgencyContactEmail = Request("Email")
SchoolDistrict = Request("SchoolDistrict")


InUseFoundation = Request("InUseFoundation")
InUseMH = Request("InUseMH")
InUseCF = Request("InUseCF")
InUseCHUD = Request("InUseCHUD")
InUseSCH = Request("InUseSCH")
WebsiteAddress = Request("WebsiteAddress")
AgencyPhone = Request("AgencyPhone")
AgencyContact = Request("Person")
Description = Replace(Request("Description"),"'","''")

AddressL1 = Request("AddressL1")
AddressL2 = Request("AddressL2")
City = Request("City")
State = Request("State")
Zip = Request("Zip")
AddressId = Request("AddressId")
AgencyContactPhone=Request("AgencyContactPhone")
Note = Request("Note")
 
    
    Set objRS = Server.CreateObject("ADODB.recordset") 
      updatesaddress_query="update AgencyAddress1 set AddressL1='"&AddressL1&"',Phone = '"&AgencyPhone&"',AgencyContactPhone = '"&AgencyContactPhone&"',AddressL2= '"&AddressL2&"',Email='"&AgencyContactEmail&"',City='"&City&"',State='"&State&"',Zip='"&Zip&"' where AddressId='"&AddressId&"' " 
        Response.Write "strsql1: " & updatesaddress_query 
        objRS.open updatesaddress_query, conn1
    


SupervisorFullName = Replace(Request("SupervisorFullName"),"'","''")
SupervisorFullName = Split(SupervisorFullName,",")

EmailAddress = Replace(Request("EmailAddress"),"'","''")
EmailAddress = Split(Request("EmailAddress"),",")
FPhone = Split(Request("FPhone"),",")
CellPhone = Split(Request("CellPhone"),",")

SupervisorID = Split(Request("SupervisorID"),",")

    Set objRS = Server.CreateObject("ADODB.recordset") 
    if (UBound(SupervisorFullName) = 0) then
    Response.Write "Value " & Request("SupervisorFullName")
    SupervisorFullName1 = Request("SupervisorFullName")
EmailAddress1 = Request("EmailAddress")
FPhone1 = Request("FPhone")
CellPhone1 = Request("CellPhone")
FPhoneExt1 = Request("FPhoneExt")
SupervisorID1 = Request("SupervisorID")
   
        updatesupervisor_query="update Supervisor1 set SupervisorFullName='"&SupervisorFullName1&"',EmailAddress='"&EmailAddress1&"',Phone='"&FPhone1&"',Fax='"&FPhoneExt1&"',CellPhone='"&CellPhone1&"' where SupervisorID='"&SupervisorID1&"' " 
        
   
        Response.Write "strsql1: " & updatesupervisor_query 
        objRS.open updatesupervisor_query, conn1
    Else
    for	index = 0 to UBound(SupervisorFullName)
     Response.Write "Value " & index 
        updatesupervisor_query="update Supervisor1 set SupervisorFullName='"&SupervisorFullName(index)&"',EmailAddress='"&EmailAddress(index)&"',Phone='"&FPhone(index)&"',CellPhone='"&CellPhone(index)&"' where SupervisorID='"&SupervisorID(index)&"' " 
        
   
        Response.Write "strsql1: " & updatesupervisor_query 
        objRS.open updatesupervisor_query, conn1
       
    next
    End If
       
if InUseFoundation = "1" then
updateagency_query="update Agency1 set AgencyID='"&AgencyID&"',ContactFoundation = '"&AgencyContact&"',Description = '"&Description&"', Agency='"&Agency&"',SchoolDistrict='"&SchoolDistrict&"',InUseFoundation='"&InUseFoundation&"',InUseMH='"&InUseMH&"',InUseCF='"&InUseCF&"',InUseCHUD='"&InUseCHUD&"', InUseSCH='"&InUseSCH&"',WebsiteAddress='"&WebsiteAddress&"',Phone='"&AgencyPhone&"',AgencyContactPhone='"&AgencyContactPhone&"' where AgencyID='"&AgencyID&"' " 
elseif InUseMH = "1" then
updateagency_query="update Agency1 set AgencyID='"&AgencyID&"',ContactFoundation = '"&AgencyContact&"',ContactMH = '"&AgencyContact&"',Description = '"&Description&"', Agency='"&Agency&"',SchoolDistrict='"&SchoolDistrict&"',InUseFoundation='"&InUseFoundation&"',InUseMH='"&InUseMH&"',InUseCF='"&InUseCF&"',InUseCHUD='"&InUseCHUD&"', InUseSCH='"&InUseSCH&"',WebsiteAddress='"&WebsiteAddress&"',Phone='"&AgencyPhone&"',AgencyContactPhone='"&AgencyContactPhone&"' where AgencyID='"&AgencyID&"' " 
elseif InUseCF = "1" then
updateagency_query="update Agency1 set AgencyID='"&AgencyID&"',ContactFoundation = '"&AgencyContact&"',ContactCF = '"&AgencyContact&"',Description = '"&Description&"', Agency='"&Agency&"',SchoolDistrict='"&SchoolDistrict&"',InUseFoundation='"&InUseFoundation&"',InUseMH='"&InUseMH&"',InUseCF='"&InUseCF&"',InUseCHUD='"&InUseCHUD&"', InUseSCH='"&InUseSCH&"',WebsiteAddress='"&WebsiteAddress&"',Phone='"&AgencyPhone&"',AgencyContactPhone='"&AgencyContactPhone&"' where AgencyID='"&AgencyID&"' " 
elseif InUseCHUD = "1" then
updateagency_query="update Agency1 set AgencyID='"&AgencyID&"',ContactFoundation = '"&AgencyContact&"',ContactCHUD = '"&AgencyContact&"',Description = '"&Description&"', Agency='"&Agency&"',SchoolDistrict='"&SchoolDistrict&"',InUseFoundation='"&InUseFoundation&"',InUseMH='"&InUseMH&"',InUseCF='"&InUseCF&"',InUseCHUD='"&InUseCHUD&"', InUseSCH='"&InUseSCH&"',WebsiteAddress='"&WebsiteAddress&"',Phone='"&AgencyPhone&"',AgencyContactPhone='"&AgencyContactPhone&"' where AgencyID='"&AgencyID&"' " 
else
updateagency_query="update Agency1 set AgencyID='"&AgencyID&"',ContactFoundation = '"&AgencyContact&"',ContactSCH = '"&AgencyContact&"',Description = '"&Description&"', Agency='"&Agency&"',SchoolDistrict='"&SchoolDistrict&"',InUseFoundation='"&InUseFoundation&"',InUseMH='"&InUseMH&"',InUseCF='"&InUseCF&"',InUseCHUD='"&InUseCHUD&"', InUseSCH='"&InUseSCH&"',WebsiteAddress='"&WebsiteAddress&"',Phone='"&AgencyPhone&"',AgencyContactPhone='"&AgencyContactPhone&"' where AgencyID='"&AgencyID&"' " 
end if

     Set objRS = Server.CreateObject("ADODB.recordset") 
     objRS.open updateagency_query, conn1

        set rs4 = Server.CreateObject("ADODB.recordset")	
    selectmaxnoteid_query="select max(AgencyNotesID)+1 as MaxAgencyNoteID from AgencyNotes1"
    rs4.Open selectmaxnoteid_query,conn1
    AgencyNotesID = rs4("MaxAgencyNoteID")

         set rs5 = Server.CreateObject("ADODB.recordset")	
    select_query="select * from AgencyNotes1 where AgencyID= '"&AgencyID&"' " 
    rs5.Open select_query,conn1

   

    if rs5.EOF then
    insertagencynotes_query="insert into AgencyNotes1 (Note,AgencyID,AgencyNotesID) values('"&Note&"', '"&AgencyID&"','"&AgencyNotesID&"') " 
    Set objRS2 = Server.CreateObject("ADODB.recordset") 
        Response.Write "strsql1: " & insertagencynotes_query 
        objRS2.open insertagencynotes_query, conn1
    else
    
      updateagencynotes_query="update AgencyNotes1 set Note = '"&Note&"' where AgencyID='"&AgencyID&"' " 
    Set objRS1 = Server.CreateObject("ADODB.recordset") 
        Response.Write "strsql1: " & updateagencynotes_query 
        objRS1.open updateagencynotes_query, conn1 
    
   end if
conn.close
conn1.close

        Response.Redirect "EditSISAgency.asp?AgencyID="&AgencyID&"&ErrMsg='Student Information Updated to DB'"
%>
