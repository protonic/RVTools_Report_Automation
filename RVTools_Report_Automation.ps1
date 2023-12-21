start-transcript C:\Automation\logs\rvtool_logs$(((get-date).ToUniversalTime()).ToString("yyyyMMddThhmmss")).txt

##########################
# Name	RVTools_Report_Automation
# By		Techwanderers youtube channel: 
# Date	December 2023
# Version	2.0
##########################


cd "c:\program files (x86)\robware\rvtools"

#rem =========================
#rem Set environment variables
#rem =========================
#set $VCServer=VCSERVER
#set $SMTPserver=mailexchange.company.net
#set $SMTPport=25
#set $Mailto=client@company.com
#set $Mailfrom=vcadmin@company.com
#set $Mailsubject="RVTools Report for last month"
#set $AttachmentDir="C:\Automation\RVToolsReporting"
#set $AttachmentFile="RVTools_export_all_Report$(((get-date).ToUniversalTime()).ToString("yyyyMMddThhmmss")).xlsx"

#rem ===================
#rem Start RVTools batch 
#rem ===================
.\rvtools.exe -passthroughAuth -s VCSERVER.company.net -c ExportAll2xls -d "C:\Automation\RVToolsReporting" -f "RVTools_export_all_Report$(((get-date).ToUniversalTime()).ToString("yyyyMMddThhmmss")).xlsx"

#rem =========
#rem Send mail
#rem =========
#rvtoolssendmail.exe /smtpserver m /smtpport %$SMTPport% /mailto %$Mailto% /mailfrom %$Mailfrom% /mailsubject %$Mailsubject% /attachment %$AttachmentDir%\%$AttachmentFile%

Send-MailMessage -To client@company.com -From vcadmin@company.com  -Subject "Rvtools Report Monthly" -Body "Hi Team, please find the report attahed for this month" -SmtpServer mailexchange.company.net -Attachments "\\VCSERVER\C$\Automation\RVToolsReporting\*.xlsx"


stop-transcript