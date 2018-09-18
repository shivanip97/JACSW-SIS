<!--#include file="md5.asp"-->
<!--#include file="DBconn.asp"-->
<%
Agency =Replace(Request("Agency"),"'","''") 
FieldType = Request("FieldType")
EmailAddress = Request("Email")
SchoolDistrict=Request("SchoolDistrict")
CHFConcentration=Request("CHFConcentration")
SCHConcentration = Request("SCHConcentration")
MHConcentration=Request("MHConcentration")
CHUDConcentration=Request("CHUDConcentration")
AddressL1=Request("AddressL1")
AddressL2=Request("AddressL2")
City = Request("City")
State=Request("State")
Zip=Request("Zip")
AgencyPhone=Request("AgencyPhone")
AgencyContactPhone=Request("AgencyContactPhone")
InUseFoundation = Request("InUseFoundation")
InUseMH=Request("InUseMH")
InUseCF=Request("InUseCF")
InUseCHUD = Request("InUseCHUD")
InUseSCH=Request("InUseSCH")
WebsiteAddress=Request("WebsiteAddress")
Active=Request("Active")
Description=Replace(Request("Description"),"'","''")
Person=Request("Person")
Note=Request("Note")

set rs = Server.CreateObject("ADODB.recordset")	
	selectmaxid_query="select max(AgencyID)+1 as MaxID from Agency1"
	rs.Open selectmaxid_query,conn1
AgencyID = rs("MaxID")

set rs3 = Server.CreateObject("ADODB.recordset")	
	selectmaxaddid_query="select max(AddressId)+1 as MaxAddressID from AgencyAddress1"
	rs3.Open selectmaxaddid_query,conn1
	AddressId = rs3("MaxAddressID")

	set rs4 = Server.CreateObject("ADODB.recordset")	
	selectmaxnoteid_query="select max(AgencyNotesID)+1 as MaxAgencyNoteID from AgencyNotes1"
	rs4.Open selectmaxnoteid_query,conn1
	AgencyNotesID = rs4("MaxAgencyNoteID")

set rs1=Server.CreateObject("ADODB.recordset")	
	insertagencydataonline_query="insert into AgencyDataforOnline1(AgencyID,Agency,AddressL1,City,State,Zip,Description,Phone,FieldType,InUseFoundation,InUseCF,InUseCHUD,InUseMH,InUseSCH,SchoolDistrict,WebsiteAddress) values ('"&AgencyID&"','"&Agency&"','"&AddressL1&"','"&City&"','"&State&"','"&Zip&"','"&Description&"','"&AgencyPhone&"','"&FieldType&"','"&InUseFoundation&"','"&InUseCF&"','"&InUseCHUD&"','"&InUseMH&"','"&InUseSCH&"','"&SchoolDistrict&"','"&WebsiteAddress&"')"
	rs1.Open insertagencydataonline_query,conn1

set rs2=Server.CreateObject("ADODB.recordset")	
	if InUseFoundation = "1" then
	insertagency_query="insert into Agency1(AgencyContactPhone,ContactFoundation,AgencyID,Agency,AgencyAddress,AgencyCity,State,Zipcode,Description,Phone,FieldType,InUseFoundation,InUseCF,InUseCHUD,InUseMH,InUseSCH,SchoolDistrict,WebsiteAddress,CHFConcentration,SCHConcentration,MHConcentration,CHUDConcentration,FT,Evening,Weekend,JanStart,SmrBlk,Stipend,Active,CAPConcentration,HLTConcentration) values ('"&AgencyContactPhone&"','"&Person&"','"&AgencyID&"','"&Agency&"','"&AddressL1&"','"&City&"','"&State&"','"&Zip&"','"&Description&"','"&AgencyPhone&"','"&FieldType&"','"&InUseFoundation&"','"&InUseCF&"','"&InUseCHUD&"','"&InUseMH&"','"&InUseSCH&"','"&SchoolDistrict&"','"&WebsiteAddress&"','"&CHFConcentration&"','"&SCHConcentration&"','"&MHConcentration&"','"&CHUDConcentration&"','','','','','','','','','')"
	elseif InUseMH ="1" then
	insertagency_query="insert into Agency1(AgencyContactPhone,ContactFoundation,ContactMH,AgencyID,Agency,AgencyAddress,AgencyCity,State,Zipcode,Description,Phone,FieldType,InUseFoundation,InUseCF,InUseCHUD,InUseMH,InUseSCH,SchoolDistrict,WebsiteAddress,CHFConcentration,SCHConcentration,MHConcentration,CHUDConcentration,FT,Evening,Weekend,JanStart,SmrBlk,Stipend,Active,CAPConcentration,HLTConcentration) values ('"&AgencyContactPhone&"','"&Person&"','"&Person&"','"&AgencyID&"','"&Agency&"','"&AddressL1&"','"&City&"','"&State&"','"&Zip&"','"&Description&"','"&AgencyPhone&"','"&FieldType&"','"&InUseFoundation&"','"&InUseCF&"','"&InUseCHUD&"','"&InUseMH&"','"&InUseSCH&"','"&SchoolDistrict&"','"&WebsiteAddress&"','"&CHFConcentration&"','"&SCHConcentration&"','"&MHConcentration&"','"&CHUDConcentration&"','','','','','','','','','')"
	elseif InUseCF ="1" then
	insertagency_query="insert into Agency1(AgencyContactPhone,ContactFoundation,ContactCF,AgencyID,Agency,AgencyAddress,AgencyCity,State,Zipcode,Description,Phone,FieldType,InUseFoundation,InUseCF,InUseCHUD,InUseMH,InUseSCH,SchoolDistrict,WebsiteAddress,CHFConcentration,SCHConcentration,MHConcentration,CHUDConcentration,FT,Evening,Weekend,JanStart,SmrBlk,Stipend,Active,CAPConcentration,HLTConcentration) values ('"&AgencyContactPhone&"','"&Person&"','"&Person&"','"&AgencyID&"','"&Agency&"','"&AddressL1&"','"&City&"','"&State&"','"&Zip&"','"&Description&"','"&AgencyPhone&"','"&FieldType&"','"&InUseFoundation&"','"&InUseCF&"','"&InUseCHUD&"','"&InUseMH&"','"&InUseSCH&"','"&SchoolDistrict&"','"&WebsiteAddress&"','"&CHFConcentration&"','"&SCHConcentration&"','"&MHConcentration&"','"&CHUDConcentration&"','','','','','','','','','')"
	elseif InUseCHUD ="1" then
	insertagency_query="insert into Agency1(AgencyContactPhone,ContactFoundation,ContactCHUD,AgencyID,Agency,AgencyAddress,AgencyCity,State,Zipcode,Description,Phone,FieldType,InUseFoundation,InUseCF,InUseCHUD,InUseMH,InUseSCH,SchoolDistrict,WebsiteAddress,CHFConcentration,SCHConcentration,MHConcentration,CHUDConcentration,FT,Evening,Weekend,JanStart,SmrBlk,Stipend,Active,CAPConcentration,HLTConcentration) values ('"&AgencyContactPhone&"','"&Person&"','"&Person&"','"&AgencyID&"','"&Agency&"','"&AddressL1&"','"&City&"','"&State&"','"&Zip&"','"&Description&"','"&AgencyPhone&"','"&FieldType&"','"&InUseFoundation&"','"&InUseCF&"','"&InUseCHUD&"','"&InUseMH&"','"&InUseSCH&"','"&SchoolDistrict&"','"&WebsiteAddress&"','"&CHFConcentration&"','"&SCHConcentration&"','"&MHConcentration&"','"&CHUDConcentration&"','','','','','','','','','')"
	 else 
	 insertagency_query="insert into Agency1(AgencyContactPhone,ContactFoundation,ContactSCH,AgencyID,Agency,AgencyAddress,AgencyCity,State,Zipcode,Description,Phone,FieldType,InUseFoundation,InUseCF,InUseCHUD,InUseMH,InUseSCH,SchoolDistrict,WebsiteAddress,CHFConcentration,SCHConcentration,MHConcentration,CHUDConcentration,FT,Evening,Weekend,JanStart,SmrBlk,Stipend,Active,CAPConcentration,HLTConcentration) values ('"&AgencyContactPhone&"','"&Person&"','"&Person&"','"&AgencyID&"','"&Agency&"','"&AddressL1&"','"&City&"','"&State&"','"&Zip&"','"&Description&"','"&AgencyPhone&"','"&FieldType&"','"&InUseFoundation&"','"&InUseCF&"','"&InUseCHUD&"','"&InUseMH&"','"&InUseSCH&"','"&SchoolDistrict&"','"&WebsiteAddress&"','"&CHFConcentration&"','"&SCHConcentration&"','"&MHConcentration&"','"&CHUDConcentration&"','','','','','','','','','')"
 end if
		
	rs2.Open insertagency_query,conn1

	set rs5=Server.CreateObject("ADODB.recordset")	
	insertagencyadd_query="insert into AgencyAddress1(AgencyContactPhone,AddressId,AgencyID,AddressL1,AddressL2,City,State,Zip,Phone,Email,PrimaryAddress,MailingAddress) values ('"&AgencyContactPhone&"','"&AddressId&"','"&AgencyID&"','"&AddressL1&"','"&AddressL2&"','"&City&"','"&State&"','"&Zip&"','"&AgencyPhone&"','"&EmailAddress&"','','')"
	rs5.Open insertagencyadd_query,conn1

	set rs6=Server.CreateObject("ADODB.recordset")	
	insertagencynote_query="insert into AgencyNotes1(AgencyId,AgencyNotesID,Note) values ('"&AgencyID&"','"&AgencyNotesID&"','"&Note&"')"
	rs6.Open insertagencynote_query,conn1

		Response.Redirect "AddInstructorforNewAgency.asp?AgencyID="&AgencyID
	

	rs.close
	rs1.close
	rs2.close
	 rs3.close
	rs4.close
	 rs5.close
	rs6.close
conn.close
conn1.close
%>
