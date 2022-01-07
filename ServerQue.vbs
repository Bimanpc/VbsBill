<HTML>  
<BODY>  
<H1>Opening a public queue on a Web server</H1>  
<%  
    Set qinfo = Server.CreateObject ("MSMQ.MSMQQueueInfo")  
    qinfo.PathName = ".\PCQueue"  
    qinfo.Label = "PV Queue"  
    On Error Resume Next      'Disregard this error if the queue exists.  
    qinfo.Create  
    On Error Goto 0  
    Set qPub = qinfo.Open (2, 0)  
  
%>  
</BODY>  
</HTML>