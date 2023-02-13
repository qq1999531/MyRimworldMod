<!--#include virtual="/inc/config.inc" -->
<!--#include  virtual="/public/Agent_check.asp" -->
<META HTTP-EQUIV="Content-Type" CONTENT="TEXT/HTML; CHARSET=BIG5">
<%
su_no=request("su_no")
webflag=request("webflag")
currentuser=request("currentuser")
cclist=request("cclist")
autocclist=request("autocclist")
cclist=autocclist&Trim(cclist)
CEVP=request("CEVP")
Shipper=request("Shipper")
CEVPflag=request("CEVPflag")
result_E2=request("result_E2")
Assignflag=md5(year(date())&month(date())&day(date())&hour(time())&minute(time())&second(time()),32)
strcurrentuser="select * from EIPUser where userid='"&currentuser&"'"
set RSCurrentUser=objConn.execute(strcurrentuser)
strSample="select * from Sample where su_no='"&Trim(su_no)&"'"
set RSSample=objConn.execute(strSample)
if not RSSample.eof then
    createUserID=RSSample("createUserID")
    Applicant=RSSample("a_name")
    RequestDate=RSSample("a_date")
	BU=RSSample("BU")
	S_Customer=RSSample("S_Customer")
	Product=RSSample("Product")
	product_no=RSSample("product_no")
	PM=RSSample("PM")
	ME1=RSSample("ME")
	PC=RSSample("PC")
	QA=RSSample("QA")
	times=RSSample("times")
	fail=RSSample("fail")
	systemeip=RSSample("systemeip")
end if

if not RSCurrentUser.eof then
   currentusername=RSCurrentUser("userForeignName")
   currentuseremail=RSCurrentUser("userEmail")
end if
currentrole=request("currentrole")
'BUHead=request("BUHead")
comment=XMLEncode(request("comment"),"comment")


if comment="" then
    Response.Write "<script>alert(""Please input your comments !"");</script>"
    Response.Write "<script>history.back()</script>"
    response.end
end if

'strSamplepn="select * from SamplePart where su_no='"&Trim(su_no)&"' order by cast(itemno as int) "
'set RSSamplepn=objConn.execute(strSamplepn)
'while not RSSamplepn.eof

'strSamplefail="select * from SampleHistory where S_Customer='"&S_Customer&"' and product_no='"&product_no&"'  "
'set RSSamplefail=objConn.execute(strSamplefail)
'if not RSSamplefail.eof  then
'	fail=RSSamplefail("fail")
'end if
'strupdatemodifydate="update Sample set fail='"&fail&"',lastmodifydate=getdate() where su_no='"&su_no&"'"
'objConn.execute(strupdatemodifydate)

'RSSamplepn.movenext
'wend


if CEVPflag="Y" then

if CEVP="" then
    Response.Write "<script>alert(""Pleasen select SVP!"");</script>"
    Response.Write "<script>history.back()</script>"
    response.end
end if
   strCEVP="select distinct userID,userName from UserList where  userID='"&CEVP&"'"
  set rsCEVP=objconn.execute(strCEVP)
    if not rsCEVP.eof then
   CEVPName1=rsCEVP("userName")

    end if

 sendtoid=trim(CEVP)
   ToRole="SVP"
'response.end

strstatusname="select * from StatusCode where statuscode='VP01'"
    set RSStatusName=objConn.execute(strstatusname)
   if not RSStatusName.eof then
       statusname=Trim(RSStatusName("statusDesc"))
   end if
strUpdate="update Sample set   CEVP='"&CEVP&"',CEVPName='"&CEVPName1&"',result_E2='"&result_E2&"',opinion_E2='"&comment&"',BUMGR_date=getdate(),currentstatus='VP01',currentStatusName='"&statusname&"', currentUser='"&CEVP&"',currentrole='SVP',TEXT4='"&Assignflag&"' where su_no='"&su_no&"'"


else




if Shipper="" then
    Response.Write "<script>alert(""Pleasen select Shipper!"");</script>"
    Response.Write "<script>history.back()</script>"
    response.end
end if
   strShipper="select distinct userID,userName from UserList where  userID='"&Shipper&"'"
    set rsShipper=objconn.execute(strShipper)
    if not rsShipper.eof then
    ShipperName=rsShipper("userName")
     end if


 sendtoid=trim(Shipper)
 ToRole="Shipper"
'Response.Write sendtoid
'esponse.end
'strcc="select * from CCList where BU='"&Trim(BU)&"' and flag='Y' and CCType='BU'"
'set RSCC=objConn.execute(strcc)
'if not RSCC.eof then
 '   While Not RSCC.EOF
 '   cclist=Trim(cclist)&RSCC("UserID")&","
'
'    RSCC.MoveNext
'    Wend
'end if



strstatusname="select * from StatusCode where statuscode='SH01'"
    set RSStatusName=objConn.execute(strstatusname)
   if not RSStatusName.eof then
       statusname=Trim(RSStatusName("statusDesc"))
   end if
strUpdate="update Sample set   Shipper='"&Shipper&"',ShipperName='"&ShipperName&"',result_E2='"&result_E2&"',opinion_E2='"&comment&"',BUMGR_date=getdate(),currentstatus='SH01',currentStatusName='"&statusname&"', currentUser='"&Shipper&"',currentrole='Shipper',TEXT4='"&Assignflag&"' where su_no='"&su_no&"'"


end if




'Response.Write strUpdate
'response.end
objConn.execute(strUpdate)
strupdatemodifydate="update Sample set lastmodifydate=getdate() where su_no='"&su_no&"'"
objConn.execute(strupdatemodifydate)
strLog="Insert into userAction(userid,username,occurtime,su_no,actiondesc) values('"&currentuser&"','"&currentusername&"',getdate(),'"&su_no&"','Approved Sample')"
objConn.execute(strLog)
'Response.write strLog

                 cctoid=Trim(cclist)
                  if Trim(cctoid)<>"" then
                       cctoid=split(cctoid,",")
                       ccemail=""
                       ccname=""
                       for i=0 to ubound(cctoid)
                             ccid=cctoid(i)
                             sql_approver="select * from EIPUser   where userid='"&ccid&"' and userIDActive='Y'"
                             set rs_approver=objconn.execute(sql_approver)
                             if not rs_approver.eof  then
                                 ccemail=ccemail&rs_approver("userEmail")&","
                                 ccname=ccname&rs_approver("userForeignname")&","
                             end if
                       next
                 end if
                strComment="insert into ProcessRecord(su_no,userID,userName,userRole,action,actiontime,comment,copyto) values('"&su_no&"','"&currentuser&"','"&currentusername&"','"&currentrole&"','Approved',getdate(),'"&comment&"','"&ccname&"')"
                objConn.execute(strcomment)

	'================EIP代理使用======================
	StrGetAgent=CheckAgent(sendtoid,systemeip)
	StrGetAgentarray=Split(StrGetAgent,";")
                  	If StrGetAgentarray(0)="Y" Then
		agentflag="Y"
		agentname=trim(StrGetAgentarray(3))
		toname=agentname
         sql_approver="select * from userList a,EIPUser b  where a.userid='"&sendtoid&"'  and a.userID=b.userID"
         set rs_approver=objconn.execute(sql_approver)
		'flow_c=flow_c&rs_approver("userForeignname")&"("&agentname&"代)"&"-->"
		toemail=trim(StrGetAgentarray(4))
'		'formStatus=agentname&"審核中"
	else

        sql_approver="select * from EIPUser   where userid='"&sendtoid&"'"
        set rs_approver=objconn.execute(sql_approver)
		toname=rs_approver("userForeignname")
		toemail=rs_approver("userEmail")
'		'formStatus=toname&"審核中"

	End if
	'================================================

                 'toname="SSD"
                 'toemail="SSD@foxconn.com"


    femail =currentuseremail
    fname=currentusername
ccsubject=""
ccbody=""
strSamplepn="select * from SamplePart where su_no='"&Trim(su_no)&"' order by cast(itemno as int) "
set RSSamplepn=objConn.execute(strSamplepn)
while not RSSamplepn.eof
product_no=RSSamplepn("product_no")
times=RSSamplepn("times")
fail=RSSamplepn("fail")
ccsubject=ccsubject&" ["& product_no &" Sample :  " & trim(times) & " times, Fail :  " & trim(fail) & " times]"
ccbody=ccbody&"HH P/N: " & trim(product_no) & ",  Sample :  " & trim(times) & " times, Fail :  " & trim(fail) & " times <br>"
RSSamplepn.movenext
wend




				 '為0表示不發郵寄至WebAdmin信箱
	strHTML="<A HREF='" & ServerFulladdr & "/Management/Sample_Approve.asp?su_no="& su_no & "&currentrole=" & ToRole & "&currentuser=" & sendtoid  & "&systemid=" & systemeip & "&Assignflag=" & Assignflag & "&webflag=0" &"'>"&su_no&"</A>"
    strccHTML="<A HREF='" & ServerFulladdr & "/Management/Sample_View.asp?su_no="& su_no &"&webflag=0" &"'>"&su_no&"</A>"
    session("mailfromemail")=femail
	session("mailtoemail")=toemail
	session("mailccemail")=ccemail
    session("mailccstrhtml")=strccHTML

if CEVPflag="Y" then
	session("mailccsubject")="Sample C.C. Notice : "&fname&" [ SBU Manager  ] approve Sample  " & su_no &" to "&toname&" [ SVP ] !"
	session("mailccbody")=fname&" [ SBU Manager  ] approve this Sample. " &"<br><br>"&ccbody &" Customer :  " & trim(S_Customer) & "<br>Product Series :  " & trim(Product) & "<br>Request Date : " & trim(RequestDate) & "<br>Request By : "& Applicant
    session("mailfname")=fname
	session("mailstrhtml")=strHTML
	session("mailsubject")="Sample approve Notice : Sample "& su_no  &"  need your approve !"
	session("mailbody")=fname&" [ SBU Manager] approve this Sample, Please approve it !<br><br>"&ccbody &" Customer :  " & trim(S_Customer) & "<br>Product Series :  " & trim(Product) & "<br>Request Date : " & trim(RequestDate) & "<br>Request By : "& Applicant
	session("memo1")="This Sample is submitted to"							   '核準後顯示的畫面
	session("memo2")=toname								'核準後顯示的畫面
	session("memo3")="for approve"								 '核準後顯示的畫面
else
'	session("mailccsubject")="Sample C.C. Notice : Sample request has been approved  " & su_no &" ["& trim(S_Customer)  &"]- "& product_no  &" ("& product  &") [Sample " & trim(times) & " times, Fail :  " & trim(fail) & " times]"
    session("mailccsubject")="Sample C.C. Notice : Sample request has been approved  " & su_no &" ["& trim(S_Customer)  &"]- "& ccsubject
	session("mailccbody")=fname&" [ SBU Manager  ] approve this Sample. " &"<br><br>"&ccbody &"Customer :  " & trim(S_Customer) & "<br>Product Series :  " & trim(Product) & "<br>Request Date : " & trim(RequestDate) & "<br>Request By : "& Applicant
'    session("mailccbody")=fname&" [ SBU Manager  ] approve this Sample. " &"<br><br>Sample :  " & trim(times) & " times, Fail :  " & trim(fail) & " times <br>Customer :  " & trim(S_Customer) & "<br>Product Series :  " & trim(Product) & "<br>Request Date : " & trim(RequestDate) & "<br>Request By : "& Applicant
    session("mailfname")=fname
	session("mailstrhtml")=strHTML
	session("mailsubject")="Sample approve Notice : Sample request "& su_no  &"  has been approved , Please preparing sample and sending.!"
	session("mailbody")=fname&" [ SBU Manager] approve this Sample, Please preparing sample and sending. !<br><br>"&ccbody &" Customer :  " & trim(S_Customer) & "<br>Product Series :  " & trim(Product) & "<br>Request Date : " & trim(RequestDate) & "<br>Request By : "& Applicant
	session("memo1")="This Sample is submitted to"							   '核準後顯示的畫面
	session("memo2")=toname								'核準後顯示的畫面
	session("memo3")="for sending"
end if

	session("turnurl")="../Management/Sample_View.asp?su_no="&su_no&"&currentrole="&ToRole&"&currentuser="&sendtoid&"&sys01=" & systemid&"&webflag=1"     'mail後要返回的網址





	if webflag=0 then
                  session("sendflag")=1				 '為1的話表示在notes中作業
	else
                   session("sendflag")=0
                  end if				 '為1的話表示在notes中作業
	session("debugflag")=0					'為1的話表示顯示郵寄的各項參數﹐Debug時使用
	session("adminflag")=0					 '為0表示不發郵寄至WebAdmin信箱

	response.redirect"/public/Pub_MailSend.asp"
    Response.end

%>