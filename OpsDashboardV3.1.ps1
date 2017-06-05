#############################################################################################################
######################## CHANGE LOG #########################################################################
#############################################################################################################
#############################################################################################################
#############################################################################################################
# Call .Net objects for forms and drawing.
#
#AD Module
import-module activedirectory
Import-Module ActiveDirectory -Cmdlet Get-aduser
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$Version = "v3.1"
# Create form
$OpsDashForm = New-Object System.Windows.Forms.Form
$OpsDashForm.Size = New-Object System.Drawing.Size(1200,600)
$OpsDashForm.StartPosition = "CenterScreen"
$OpsDashForm.Text = "Operation Support Dashboard $Version"
$TechEmail = $(Get-WMIObject -class Win32_ComputerSystem | select username).username
# Gets the current logged in user
$Technician = $TechEmail.Substring(6)

######## Auth Code Button function
function Auth_Code
	{
	# Actual function for Auth codes
    # Gets the date for logging
    $Today = get-date
    $Time2 = Get-Date -Format t
    $filedate = get-date -format "MMddyyyy"
    $filedateexcel = get-date -format "MM/dd/yyyy"
    $SAMAccount = $SAMBox.text
    #Change @yahoo to your own domain
    $SA = $SAMAccount+"@yahoo.com"
    $ML= $comboBoxCompany.Text+$listBoxClub.text
    $MarketD = $comboBoxCompany.Text
    $LocationD = $listBoxClub.text
    $ST = $TicketStatus.Text
    $Time = $Timer.Text
    $AC = $Additionalcomments.text
    $PCName = $inputboxPC.Text
    $Formoutput = "You ran Auth Code at $Today"
    $outputbox.text = [string]$formoutput
    #Change email adderss to what ever users send an email normally to to generate a ticket
    $recipients = "Ticksummit@yahoo.com"
       $subject = "Auth code for $SAMAccount at $ML for PC - $PCName"
       $Sender = "$SA"
       #This is the body of the email being sent to itgroup
       $mailbody = "
Priority=3-Medium
Status=$ST
Subcategory=Auth
Location Code=$ML
Submission Tracking=Phone
assignees=+$Technician
Category=Software
Computer Name = $PCName
Email = $SA
$Time
We have provided an authentication for $SAMAccount at $ML for PC - $PCName
$AC

Ticket summitted by Ops Dashboard
"
#########
# This command sends the email
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
        #This creates the name of the log file
		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_AuthCode" + "_$filedate" + ".txt"
        #This puts the body of the email into the log file
        $mailbody | out-file $filename
        #This sends the ticket information to a spreadsheet for reporting
        $msg="$SA,$MarketD,$LocationD,$PCName,$Technician,AuthCode,$filedateexcel,$Time2,$Time,Software,Auth,$Version"
        #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
	}

########  Set Time Button function
function Set_Time
	{
	# Actual function:
    $Today = get-date
    $Time2 = Get-Date -Format t
    $filedate = get-date -format "MMddyyyy"
    $filedateexcel = get-date -format "MM/dd/yyyy"
    $SAMAccount = $SAMBox.text
    $SA = $SAMAccount+"@yahoo.com"
    $ML= $comboBoxCompany.Text+$listBoxClub.text
    $MarketD = $comboBoxCompany.Text
    $LocationD = $listBoxClub.text
    $PCName = $inputboxPC.Text
    $ST = $TicketStatus.Text
    $AC = $Additionalcomments.text
    $Time = $Timer.Text
    $Formoutput = "You ran Set Time at $Today"
    $outputbox.text = [string]$formoutput
    $recipients = "Ticksummit@yahoo.com"
       $subject = "Issue at $ML for PC $PCName"
       $Sender = "$SA"
       $mailbody = "
Priority=3-Medium
Status=$ST
Subcategory=eROES Token
Location Code=$ML
Submission Tracking=Phone
assignees=+$Technician
Category=Software
Computer Name = $PCName
Email = $SA
$Time
Changed time on $PCName to allow login.
$AC

Ticket summitted by Ops Dashboard
"
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_SetTime" + "_$filedate" + ".txt"
        $mailbody | out-file $filename
		$msg="$SA,$MarketD,$LocationD,$PCName,$Technician,SetTime,$filedateexcel,$Time2,$Time,Software,eROES Token,$Version"
        #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
	}

########  Term Button Button function
function Term_button
    {


    }

######## Tech Dispatch function
function Tech_Dispatch
	{
	# Actual function:
    $Today = get-date
    $Time2 = Get-Date -Format t
    $filedate = get-date -format "MMddyyyy"
    $filedateexcel = get-date -format "MM/dd/yyyy"
    $SAMAccount = $SAMBox.text
    $SA = $SAMAccount+"@yahoo.com"
    $ML= $comboBoxCompany.Text+$listBoxClub.text
    $MarketD = $comboBoxCompany.Text
    $LocationD = $listBoxClub.text
    $ST = $TicketStatus.Text
    $AC = $Additionalcomments.text
    $Time = $Timer.Text
    $Formoutput = "You ran Tech Dispatch at $Today"
    $outputbox.text = [string]$formoutput
    $recipients = "Ticksummit@yahoo.com"
       $subject = "Please dispatch a field tech at $ML"
       $Sender = "$TechEmail"
       $mailbody = "
Priority=3-Medium
Status=$ST
assignees=+Test.User
Category=Software
Computer Name = $PCName
Email = $SA
$Time


Please dispatch a field to $ML detail are below
$AC

Ticket summitted by Ops Dashboard

Location Code=$ML
Submission Tracking=Phone"
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_Dispatch" + "_$filedate" + ".txt"
        $mailbody | out-file $filename
        $msg="$SA,$MarketD,$LocationD,$PCName,$Technician,FieldTech,$filedateexcel,$Time2,$Time,$Version"
        #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
	function exit-form {
	$OpsDashForm.close()
	}

}

######## Re-Token function
function Re_Token
    {
    	# Actual function:
    $Today = get-date
    $Time2 = Get-Date -Format t
    $filedate = get-date -format "MMddyyyy"
    $filedateexcel = get-date -format "MM/dd/yyyy"
    $SAMAccount = $SAMBox.text
    $SA = $SAMAccount+"@yahoo.com"
    $ML= $comboBoxCompany.Text+$listBoxClub.text
    $MarketD = $comboBoxCompany.Text
    $LocationD = $listBoxClub.text
    $ST = $TicketStatus.Text
    $AC = $Additionalcomments.text
    $PCName = $inputboxPC.Text
    $Time = $Timer.Text
    $Formoutput = "You ran Re-Token at $Today"
    $outputbox.text = [string]$formoutput
    $recipients = "Ticksummit@yahoo.com"
       $subject = "Installed Token for $PCName at $ML"
       $Sender = "$SA"
       $mailbody = "
Priority=3-Medium
Status=$ST
assignees=+$Technician
Category=Software
Computer Name = $PCName
Email = $SA

We have installed a token for $PCName at $ML
$SAMAccount
$AC

Ticket summitted by Ops Dashboard
Subcategory=eROES Token
Location Code=$ML
Submission Tracking=Phone"
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_ReToken" + "_$filedate" + ".txt"
        $mailbody | out-file $filename
        $msg="$SA,$MarketD,$LocationD,$PCName,$Technician,OMNI_Token,$filedateexcel,$Time2,$Time,Software,eROES Token,$Version"
        #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
    }

######## Other function
function Hardware
    {
    # Actual function:
    $Today = get-date
    $Time2 = Get-Date -Format t
    $filedate = get-date -format "MMddyyyy"
    $filedateexcel = get-date -format "MM/dd/yyyy"
    $SAMAccount = $SAMBox.text
    $SA = $SAMAccount+"@yahoo.com"
    $ML= $comboBoxCompany.Text+$listBoxClub.text
    $MarketD = $comboBoxCompany.Text
    $LocationD = $listBoxClub.text
    $ST = $TicketStatus.Text
    $AC = $Additionalcomments.text
    $PCName = $inputboxPC.Text
    $Time = $Timer.Text
    $Formoutput = "You ran Hardware at $Today"
    $outputbox.text = [string]$formoutput
    $Category = $comboBoxCat.Text
    $Subcategory = $listBoxSubCat.Text
    $recipients = "Ticksummit@yahoo.com"
       $subject = "$SAMAccount at $ML is having an Hardware issues"
       $Sender = "$SA"
       $mailbody = "
Priority=3-Medium
Status=$ST
Subcategory=$Subcategory
Location Code=$ML
Submission Tracking=Phone
assignees=+$Technician
Category=Hardware
Computer Name = $PCName
Email = $SA
$Time
$AC

Ticket summitted by Ops Dashboard

"
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_Hardware" + "_$filedate" + ".txt"
        $mailbody | out-file $filename
        $msg="$SA,$MarketD,$LocationD,$PCName,$Technician,Hardware,$filedateexcel,$Time2,$Time,Hardware,$Subcategory,$Version"
        #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
    }

######## Other function
function Software
    {
    # Actual function:
    $Today = get-date
    $Time2 = Get-Date -Format t
    $filedate = get-date -format "MMddyyyy"
    $filedateexcel = get-date -format "MM/dd/yyyy"
    $SAMAccount = $SAMBox.text
    $SA = $SAMAccount+"@yahoo.com"
    $ML= $comboBoxCompany.Text+$listBoxClub.text
    $MarketD = $comboBoxCompany.Text
    $LocationD = $listBoxClub.text
    $ST = $TicketStatus.Text
    $AC = $Additionalcomments.text
    $PCName = $inputboxPC.Text
    $Time = $Timer.Text
    $Formoutput = "You ran Software at $Today"
    $outputbox.text = [string]$formoutput
    $Category = $comboBoxCat.Text
    $Subcategory = $listBoxSubCat.Text
    $recipients = "Ticksummit@yahoo.com"
       $subject = "$SAMAccount at $ML is having an Software issues"
       $Sender = "$SA"
       $mailbody = "
Priority=3-Medium
Status=$ST
Subcategory=$Subcategory
Location Code=$ML
Submission Tracking=Phone
assignees=+$Technician
Category=Software
Computer Name = $PCName
Email = $SA
$Time
$AC

Ticket summitted by Ops Dashboard

"
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_Software" + "_$filedate" + ".txt"
        $mailbody | out-file $filename
        $msg="$SA,$MarketD,$LocationD,$PCName,$Technician,Software,$filedateexcel,$Time2,$Time,Software,$Subcategory,$Version"
        #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
    }


######## AD unlock function
function AD_Unlock
    {
    # Actual function:
    $Today = get-date
    $Time2 = Get-Date -Format t
    $filedate = get-date -format "MMddyyyy"
    $filedateexcel = get-date -format "MM/dd/yyyy"
    $SAMAccount = $SAMBox.text
    $SA = $SAMAccount+"@yahoo.com"
    $ML= $comboBoxCompany.Text+$listBoxClub.text
    $MarketD = $comboBoxCompany.Text
    $LocationD = $listBoxClub.text
    $ST = $TicketStatus.Text
    $AC = $Additionalcomments.text
    $PCName = $inputboxPC.Text
    $Time = $Timer.Text
    $Formoutput = "You ran AD at $Today"
    $outputbox.text = [string]$formoutput
    $recipients = "Ticksummit@yahoo.com"
       $subject = "Password or account was unlock for $SAMAccount at $ML"
       $Sender = "$SA"
       $mailbody = "
Priority=3-Medium
Status=$ST
Subcategory=AD
Location Code=$ML
Submission Tracking=Phone
assignees=+$Technician
Category=Software
$Time
$PCName
Password or account was unlock for $SA at $ML
$AC

Ticket summitted by Ops Dashboard
"
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_ADUnlock" + "_$filedate" + ".txt"
        $mailbody | out-file $filename
        $msg="$SA,$MarketD,$LocationD,$PCName,$Technician,AD_Unlock,$filedateexcel,$Time2,$Time,Software,AD,$Version"
        #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
    }

######## Internet funtion
function Internet
    {
    # Actual function:
    $Today = get-date
    $Time2 = Get-Date -Format t
    $filedate = get-date -format "MMddyyyy"
    $filedateexcel = get-date -format "MM/dd/yyyy"
    $SAMAccount = $SAMBox.text
    $SA = $SAMAccount+"@yahoo.com"
    $ML= $comboBoxCompany.Text+$listBoxClub.text
    $MarketD = $comboBoxCompany.Text
    $LocationD = $listBoxClub.text
    $ST = $TicketStatus.Text
    $AC = $Additionalcomments.text
    $PCName = $inputboxPC.Text
    $Time = $Timer.Text
    $Formoutput = "You ran Internet at $Today"
    $outputbox.text = [string]$formoutput
    $recipients = "Ticksummit@yahoo.com"
       $subject = "Internet issue at $ML"
       $Sender = "$SA"
       $mailbody = "
Priority=3-Medium
Status=$ST
Category=Network
Location Code=$ML
Submission Tracking=Phone
Subcategory=ISP Outage
assignees=+$Technician


$Time
$SAMAccount called in to report an Internet issue at $ML
$AC


Ticket summitted by Ops Dashboard
"
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_Internet" + "_$filedate" + ".txt"
        $mailbody | out-file $filename
        $msg="$SA,$MarketD,$LocationD,$PCName,$Technician,Internet,$filedateexcel,$Time2,$Time,Network,ISP Outage,$Version"
        #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
    }

######## RQ funtion
function RQ_Code
	{
	# Actual function for Auth codes
    # Gets the date for logging
    $Today = get-date
    $Time2 = Get-Date -Format t
    $filedate = get-date -format "MMddyyyy"
    $filedateexcel = get-date -format "MM/dd/yyyy"
    $SAMAccount = $SAMBox.text
    $SA = $SAMAccount+"@yahoo.com"
    $ML= $comboBoxCompany.Text+$listBoxClub.text
    $MarketD = $comboBoxCompany.Text
    $LocationD = $listBoxClub.text
    $ST = $TicketStatus.Text
    $Time = $Timer.Text
    $AC = $Additionalcomments.text
    $PCName = $inputboxPC.Text
    $Formoutput = "You ran RQ at $Today"
    $outputbox.text = [string]$formoutput
    $recipients = "Ticksummit@yahoo.com"
    $subject = "RQ issue at $ML for $SAMAccount using PC - $PCName"
    $Sender = "$SA"
       #This is the body of the email being sent to itgroup
       $mailbody = "
Priority=3-Medium
Status=$ST
assignees=+$Technician
Category=RQ
Location Code=$ML
Email = $SA
$Time
RQ issue at $ML with PC $PCName has been recorded
$AC

Ticket summitted by Ops Dashboard

Location Code=$ML
Submission Tracking=Phone"
#########
# This command sends the email
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
        #This creates the name of the log file
		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_RQ" + "_$filedate" + ".txt"
        #This puts the body of the email into the log file
        $mailbody | out-file $filename
        #This sends the ticket information to a spreadsheet for reporting
        $msg="$SA,$MarketD,$LocationD,$PCName,$Technician,RQ,$filedateexcel,$Time2,$Time,RQ,,$Version"
        #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
	}

######## Verifone funtion
function Verifone
	{
	# Actual function for Auth codes
    # Gets the date for logging
    $Today = get-date
    $Time2 = Get-Date -Format t
    $filedate = get-date -format "MMddyyyy"
    $filedateexcel = get-date -format "MM/dd/yyyy"
    $SAMAccount = $SAMBox.text
    $SA = $SAMAccount+"@yahoo.com"
    $ML= $comboBoxCompany.Text+$listBoxClub.text
    $MarketD = $comboBoxCompany.Text
    $LocationD = $listBoxClub.text
    $ST = $TicketStatus.Text
    $Time = $Timer.Text
    $AC = $Additionalcomments.text
    $PCName = $inputboxPC.Text
    $Formoutput = "You ran Verifone at $Today"
    $outputbox.text = [string]$formoutput
    $recipients = "Ticksummit@yahoo.com"
    $subject = "Verifone issue at $ML for $SAMAccount using PC - $PCName"
    $Sender = "$SA"
       #This is the body of the email being sent to itgroup
       $mailbody = "
Priority=3-Medium
Status=$ST
assignees=+$Technician
Category=Hardware
Computer Name = $PCName
Email = $SA
$Time
Verifone was not working at $ML at PC $PCName
$AC

Ticket summitted by Ops Dashboard
Subcategory=Verifone
Location Code=$ML
Submission Tracking=Phone"
#########
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
        #This creates the name of the log file
		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_Verifone" + "_$filedate" + ".txt"
        #This puts the body of the email into the log file
        $mailbody | out-file $filename
        #This sends the ticket information to a spreadsheet for reporting
        $msg="$SA,$MarketD,$LocationD,$PCName,$Technician,Verifone,$filedateexcel,$Time2,$Time,Hardware,Verifone,$Version"
        #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
	}

######## Printer funtion
function Printer
	{
	# Actual function for Auth codes
    # Gets the date for logging
    $Today = get-date
    $Time2 = Get-Date -Format t
    $filedate = get-date -format "MMddyyyy"
    $filedateexcel = get-date -format "MM/dd/yyyy"
    $SAMAccount = $SAMBox.text
    $SA = $SAMAccount+"@yahoo.com"
    $ML= $comboBoxCompany.Text+$listBoxClub.text
    $MarketD = $comboBoxCompany.Text
    $LocationD = $listBoxClub.text
    $ST = $TicketStatus.Text
    $Time = $Timer.Text
    $AC = $Additionalcomments.text
    $PCName = $inputboxPC.Text
    $Formoutput = "You ran Printer at $Today"
    $outputbox.text = [string]$formoutput
    $recipients = "Ticksummit@yahoo.com"
    $subject = "Printer issue at $ML for $SAMAccount using PC - $PCName"
    $Sender = "$SA"
       #This is the body of the email being sent to itgroup
       $mailbody = "
Priority=3-Medium
Status=$ST
Subcategory=Printer
Location Code=$ML
Submission Tracking=Phone
assignees=+$Technician
Category=Printing
Computer Name = $PCName
Email = $SA
$Time
Printer was not working at $ML at PC $PCName
$AC

Ticket summitted by Ops Dashboard
"
#########
# This command sends the email
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
        #This creates the name of the log file
		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_Printer" + "_$filedate" + ".txt"
        #This puts the body of the email into the log file
        $mailbody | out-file $filename
        #This sends the ticket information to a spreadsheet for reporting
        $msg="$SA,$MarketD,$LocationD,$PCName,$Technician,Printer,$filedateexcel,$Time2,$Time,Printing,Printer,$Version"
        #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
	}

######## Global funtion
function Global
	{
	# Actual function for Auth codes
    # Gets the date for logging
    $Today = get-date
    $Time2 = Get-Date -Format t
    $filedate = get-date -format "MMddyyyy"
    $filedateexcel = get-date -format "MM/dd/yyyy"
    $SAMAccount = $SAMBox.text
    $SA = $SAMAccount+"@yahoo.com"
    $ML= $comboBoxCompany.Text+$listBoxClub.text
    $MarketD = $comboBoxCompany.Text
    $LocationD = $listBoxClub.text
    $ST = $TicketStatus.Text
    $Time = $Timer.Text
    $AC = $Additionalcomments.text
    $PCName = $inputboxPC.Text
    $Formoutput = "You ran Global at $Today"
    $Category = $comboBoxCat.Text
    $Subcategory = $listBoxSubCat.Text
    $outputbox.text = [string]$formoutput
    $recipients = "Ticksummit@yahoo.com"
    $subject = "Global issue for $ML $SAMAccount"
    $Sender = "$SA"
       #This is the body of the email being sent to itgroup
       $mailbody = "
Priority=3-Medium
Status=$ST
Subcategory=$Subcategory
Location Code=$ML
Submission Tracking=Phone
assignees=+$Technician
Category=$Category
Computer Name = $PCName
Email = $SA
$Time
We have reported a Global issue at $ML
$AC

Ticket summitted by Ops Dashboard
"
#########
# This command sends the email
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
        #This creates the name of the log file
		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_Global" + "_$filedate" + ".txt"
        #This puts the body of the email into the log file
        $mailbody | out-file $filename
        #This sends the ticket information to a spreadsheet for reporting
        $msg="$SA,$MarketD,$LocationD,$PCName,$Technician,Global,$filedateexcel,$Time2,$Time,$Category,$Subcategory,$Version"
        #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
	}

######## Global funtion
function Events
	{
	# Actual function for Auth codes
    # Gets the date for logging
    $Today = get-date
    $Time2 = Get-Date -Format t
    $filedate = get-date -format "MMddyyyy"
    $filedateexcel = get-date -format "MM/dd/yyyy"
    $SAMAccount = $SAMBox.text
    $SA = $SAMAccount+"@yahoo.com"
    $ML= $comboBoxCompany.Text+$listBoxClub.text
    $MarketD = $comboBoxCompany.Text
    $LocationD = $listBoxClub.text
    $ST = $TicketStatus.Text
    $Time = $Timer.Text
    $AC = $Additionalcomments.text
    $PCName = $inputboxPC.Text
    $Formoutput = "You ran Events at $Today"
    $outputbox.text = [string]$formoutput
    $Category = $comboBoxCat.Text
    $Subcategory = $listBoxSubCat.Text
    $recipients = "Ticksummit@yahoo.com"
    $subject = "Event $ML had an issue/$SAMAccount"
    $Sender = "$SA"
       #This is the body of the email being sent to itgroup
       $mailbody = "
Priority=3-Medium
Status=$ST
Subcategory=$Subcategory
Location Code=$ML
Submission Tracking=Phone
assignees=+$Technician
Category=$Category
Computer Name = $PCName
Email = $SA
$Time
We have reported a issue at an event $ML
$AC

Ticket summitted by Ops Dashboard
"
#########
# This command sends the email
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
        #This creates the name of the log file
		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_Event" + "_$filedate" + ".txt"
        #This puts the body of the email into the log file
        $mailbody | out-file $filename
        #This sends the ticket information to a spreadsheet for reporting
        $msg="$SA,$MarketD,$LocationD,$PCName,$Technician,Event,$filedateexcel,$Time2,$Time,$Category,$Subcategory,$Version"
        #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
	}

  ######## Information Request Function
  function Information
  	{
  	# Actual function for Information Requests
      # Gets the date for logging
      $Today = get-date
      $Time2 = Get-Date -Format t
      $filedate = get-date -format "MMddyyyy"
      $filedateexcel = get-date -format "MM/dd/yyyy"
      $SAMAccount = $SAMBox.text
      $SA = $SAMAccount+"@yahoo.com"
      $ML= $comboBoxCompany.Text+$listBoxClub.text
      $MarketD = $comboBoxCompany.Text
      $LocationD = $listBoxClub.text
      $ST = $TicketStatus.Text
      $Time = $Timer.Text
      $AC = $Additionalcomments.text
      $PCName = $inputboxPC.Text
      $Formoutput = "You ran Information Request at $Today"
      $outputbox.text = [string]$formoutput
      $recipients = "Ticksummit@yahoo.com"
      $subject = "Information Request at $ML for $SAMAccount"
      $Sender = "$SA"
         #This is the body of the email being sent to itgroup
         $mailbody = "
  Priority=3-Medium
  Status=$ST
  assignees=+$Technician
  Category=Information Request
  Computer Name = $PCName
  Email = $SA
  $Time
  A Request for information was submitted at $ML for $SAMAccount
  $AC

  Ticket summitted by Ops Dashboard V3.0
  Subcategory=$Subcategory
  Location Code=$ML
  Submission Tracking=Phone"
  #########
  # This command sends the email
#change smtp.yahoo.local to whatever your local mail relay
send-mailmessage -smtpserver smtp.yahoo.local -to $recipients -Subject $subject -from $sender -body $mailbody -Encoding ASCII
          #This creates the name of the log file
  		$filename = "c:\OpsDashboard\$ML" + "_$SAMAccount" + "_Information" + "_$filedate" + ".txt"
          #This puts the body of the email into the log file
          $mailbody | out-file $filename
          #This sends the ticket information to a spreadsheet for reporting
          $msg="$SA,$MarketD,$LocationD,$PCName,$Technician,Information,$filedateexcel,$Time2,$Time,Information Request,$Subcategory,$Version"
          #Change the file path to what ever path you want as a csv
        $msg | Out-File -Append -FilePath "\\testdomain\test\Tickets.csv" -Encoding ASCII
  	}

# Create the label control and set text, size and location
$label_lastname_ad = New-Object Windows.Forms.Label
$label_lastname_ad.Location = New-Object Drawing.Point 10,40
$label_lastname_ad.Size = New-Object Drawing.Point 100,15
$label_lastname_ad.text = "Search lastname"

# Create TextBox and set text, size and location
$textfield_lastname_ad = New-Object Windows.Forms.TextBox
$textfield_lastname_ad.Location = New-Object Drawing.Point 10,15
$textfield_lastname_ad.Size = New-Object Drawing.Point 100,15


$OpsDashform.add_Load($OnLoadForm_UpdateGrid)

function search_contact_ad([string]$lastname_str){
if ($lastname_str) {
$array_ad = New-Object System.Collections.ArrayList
#change cd=yahoo to your domain
$Script:procInfo = @(Get-ADUser -Filter {sn -like $lastname_str} -Properties sn,givenname,mail,displayname -SearchBase "dc=yahoo,dc=local" |sort-object -property sn |Select-Object samaccountname)
$array_ad.AddRange($procInfo)
$grid.DataSource = $array_ad
$OpsDashform.refresh()
}
else {
[windows.forms.messagebox]::show('Please enter a lastname to search before clicking the button','Warning','OK',[Windows.Forms.MessageBoxIcon]::Warning)
}
}

######################################


function Ping {
$ML = $InputBox.text
$storepart1 = "$ML".Substring(0,2)
$storepart2 = "$ML".Substring(2)
$store =$storepart1 + "-" + $storepart2
$StoreDNS = "CS-" + "$store" + "-LAN"
$Connection = ping $StoreDNS
Write-Output $Connection
$Formoutput = "$Connection"
$PingBox.text = [string]$formoutput
}

######## Clears attributes in fields
function Clear_Fields
    {
    $SAMBox.text = ""
    $TicketStatus.Text = "In Progress"
    $Additionalcomments.text = ""
    $inputboxPC.Text = ""
}

######## Clears info in status box
function Clear_Status
    {
    $outputbox.text = ""
    $FieldTechInfo.text = ""
    }

function exit-form {
	$OpsDashForm.close()
	}
# Creating the form

function FieldTech {
$ML= $comboBoxCompany.Text+$listBoxClub.text
if ($ML -eq "TTUN")
{$FieldTechStatus = "John Doe
123-456-7891"
$FieldTechInfo.text = [string]$FieldTechStatus}
}


	# end functions
	# Start text fields./.
    # Input text box for SAM account
	$SAMBox = New-Object System.Windows.Forms.TextBox
	$SAMBox.Location = New-Object System.Drawing.Size(20,50)
	$SAMBox.Size = New-Object System.Drawing.Size(150,20)
    $OpsDashForm.controls.Add($SAMBox)
    $SAMBox.add_KeyPress({
    If ($_.KeyChar -eq 13) {
    $grid = $ad_grid
    $lnquery = "*"+$SAMBox.Text.ToString()+"*"
    search_contact_ad($lnquery)
    }
    })
#################################################################################################################################################################################################### End CheckBox

#################################################################################################################################################################################################### Start ListBox
# This create the dropdown for the Location field
$listBoxClub = New-Object System.Windows.Forms.ComboBox
$listBoxClub.Location = New-Object System.Drawing.Size(275,50)
$listBoxClub.Size = New-Object System.Drawing.Size(50,40)
$listBoxClub.DropDownHeight = 200
$listBoxClub.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$listBoxClub.TabIndex = 3
# If you need to add another market to the "Market" drop down do so here; i.e. ,"CT" at the correct alphabetizing
$arrayCompany=@("WT","AP")
# This creates the dropdown box for the Market field
$comboBoxCompany = New-Object System.Windows.Forms.ComboBox
$comboBoxCompany.Location = New-Object System.Drawing.Size(220,50)
$comboBoxCompany.Size = New-Object System.Drawing.Size(50,20)
$comboBoxCompany.DropDownHeight = 200
$comboBoxCompany.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboBoxCompany.TabIndex = 2
$OpsDashForm.Controls.Add($comboBoxCompany)

foreach ($company in $arraycompany) {
                      $comboBoxCompany.Items.Add($company)
                              }

$comboBoxCompany_SelectedIndexChanged=
{
   If ($comboBoxCompany.text -eq "WT")
   {
   $listBoxClub.Items.Clear()
# If you need to add a new location for the AL market please add the field here in the correct aphabetizing
[void] $listBoxClub.Items.Add("1B")
[void] $listBoxClub.Items.Add("2B")
 }

   ElseIf ($comboBoxCompany.text -eq "AP")
   {
   $listBoxClub.Items.Clear()
   # If you need to add a new location for the AZ market please add the field here in the correct aphabetizing
[void] $listBoxClub.Items.Add("TE")
[void] $listBoxClub.Items.Add("GG")
[void] $listBoxClub.Items.Add("IC")
[void] $listBoxClub.Items.Add("CW")
[void] $listBoxClub.Items.Add("4V")
[void] $listBoxClub.Items.Add("11")
   }

$OpsDashForm.Controls.Add($listBoxClub)
}
$comboBoxCompany.add_SelectedIndexChanged($comboBoxCompany_SelectedIndexChanged)

$ML= $comboBoxCompany.Text+$listBoxClub.text
if ($ML -eq "WT1B" -or $ML -eq "APTE")
{$FieldTechStatus = "Joe Doe
123-456-7891"
$FieldTechInfo.text = [string]$FieldTechStatus}
#################################################################################################################################################################################################### End ComboBox

#################################################################################################################################################################################################### Start Text Fields

# This create the dropdown for the Location field
$listBoxSubCat = New-Object System.Windows.Forms.ComboBox
$listBoxSubCat.Location = New-Object System.Drawing.Size(670,253)
$listBoxSubCat.Size = New-Object System.Drawing.Size(90,20)
$listBoxSubCat.DropDownHeight = 200
$listBoxSubCat.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$listBoxSubCat.TabIndex = 6
# If you need to add another Category to the "Category" drop down do so here; i.e. ,"Network" at the correct alphabetizing
$arrayCat=@("Events","Hardware","Information Request","Network","Printing","RQ","Software","System")
# This creates the dropdown box for the Market field
$comboBoxCat = New-Object System.Windows.Forms.ComboBox
$comboBoxCat.Location = New-Object System.Drawing.Size(540,253)
$comboBoxCat.Size = New-Object System.Drawing.Size(90,30)
$comboBoxCat.DropDownHeight = 200
$comboBoxCat.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboBoxCat.TabIndex = 5
$OpsDashForm.Controls.Add($comboBoxCat)


foreach ($Cats in $arrayCat) {
                      $comboBoxCat.Items.Add($Cats)
                              }

$comboBoxCategory_SelectedIndexChanged=
{
   If ($comboBoxCat.text -eq "Events")
   {
   $listBoxSubCat.Items.Clear()
# If you need to add a new location for the AL market please add the field here in the correct aphabetizing
[void] $listBoxSubCat.Items.Add("Network")
[void] $listBoxSubCat.Items.Add("Verifone")
[void] $listBoxSubCat.Items.Add("Tokens")
[void] $listBoxSubCat.Items.Add("AuthCode")
 }

   ElseIf ($comboBoxCat.text -eq "Hardware")
   {
   $listBoxSubCat.Items.Clear()
# If you need to add a new location for the AL market please add the field here in the correct aphabetizing
[void] $listBoxSubCat.Items.Add("Brightsign")
[void] $listBoxSubCat.Items.Add("Cash Safe")
[void] $listBoxSubCat.Items.Add("Cellebrite")
[void] $listBoxSubCat.Items.Add("Desktop")
[void] $listBoxSubCat.Items.Add("DMX")
[void] $listBoxSubCat.Items.Add("DVR")
[void] $listBoxSubCat.Items.Add("Hard Drive")
[void] $listBoxSubCat.Items.Add("iPad")
[void] $listBoxSubCat.Items.Add("Laptop")
[void] $listBoxSubCat.Items.Add("Monitor")
[void] $listBoxSubCat.Items.Add("Mouse")
[void] $listBoxSubCat.Items.Add("Phone")
[void] $listBoxSubCat.Items.Add("Scanner")
[void] $listBoxSubCat.Items.Add("Topaz")
[void] $listBoxSubCat.Items.Add("Verifone")
 }

   ElseIf ($comboBoxCat.text -eq "Network")
   {
   $listBoxSubCat.Items.Clear()
# If you need to add a new location for the AL market please add the field here in the correct aphabetizing
[void] $listBoxSubCat.Items.Add("Access Point")
[void] $listBoxSubCat.Items.Add("Avigilon")
[void] $listBoxSubCat.Items.Add("Cisco Anyconnect (VPN)")
[void] $listBoxSubCat.Items.Add("IP Phone")
[void] $listBoxSubCat.Items.Add("ISP Outage")
[void] $listBoxSubCat.Items.Add("Modem")
[void] $listBoxSubCat.Items.Add("Router")
[void] $listBoxSubCat.Items.Add("Source Fire")
[void] $listBoxSubCat.Items.Add("Switch")
}

   ElseIf ($comboBoxCat.text -eq "Printing")
   {
   $listBoxSubCat.Items.Clear()
# If you need to add a new location for the AL market please add the field here in the correct aphabetizing
[void] $listBoxSubCat.Items.Add("Paper Jam")
[void] $listBoxSubCat.Items.Add("Power")
[void] $listBoxSubCat.Items.Add("Print Quality")
[void] $listBoxSubCat.Items.Add("Printer")
[void] $listBoxSubCat.Items.Add("Printer Offline")
[void] $listBoxSubCat.Items.Add("Toner")
}

   ElseIf ($comboBoxCat.text -eq "RQ")
   {
   $listBoxSubCat.Items.Clear()
# If you need to add a new location for the AL market please add the field here in the correct aphabetizing
[void] $listBoxSubCat.Items.Add("Commisions")
[void] $listBoxSubCat.Items.Add("Datascape")
[void] $listBoxSubCat.Items.Add("Inventory")
[void] $listBoxSubCat.Items.Add("Lockouts")
[void] $listBoxSubCat.Items.Add("Permissions")
[void] $listBoxSubCat.Items.Add("Reporting")
[void] $listBoxSubCat.Items.Add("Carrier Integration")
}

   ElseIf ($comboBoxCat.text -eq "Software")
   {
   $listBoxSubCat.Items.Clear()
# If you need to add a new location for the AL market please add the field here in the correct aphabetizing
[void] $listBoxSubCat.Items.Add("Adobe Acrobat")
[void] $listBoxSubCat.Items.Add("Adobe Creative Suite")
[void] $listBoxSubCat.Items.Add("Airwatch")
[void] $listBoxSubCat.Items.Add("Android")
[void] $listBoxSubCat.Items.Add("Avamar")
[void] $listBoxSubCat.Items.Add("Brightsign")
[void] $listBoxSubCat.Items.Add("Cisco Anyconnect")
[void] $listBoxSubCat.Items.Add("CSOKI")
[void] $listBoxSubCat.Items.Add("Desktop")
[void] $listBoxSubCat.Items.Add("eROES Token")
[void] $listBoxSubCat.Items.Add("Exchange Management Console")
[void] $listBoxSubCat.Items.Add("Google Chrome")
[void] $listBoxSubCat.Items.Add("Internet Explorer")
[void] $listBoxSubCat.Items.Add("iOS")
[void] $listBoxSubCat.Items.Add("Java")
[void] $listBoxSubCat.Items.Add("Microsoft Excel")
[void] $listBoxSubCat.Items.Add("Microsoft Office")
[void] $listBoxSubCat.Items.Add("Mozilla Firefox")
[void] $listBoxSubCat.Items.Add("Sophos AV")
[void] $listBoxSubCat.Items.Add("Visual Studio")
[void] $listBoxSubCat.Items.Add("WebEx")
[void] $listBoxSubCat.Items.Add("Windows 7")
[void] $listBoxSubCat.Items.Add("Windows 10")
}

   ElseIf ($comboBoxCat.text -eq "System")
   {
   $listBoxSubCat.Items.Clear()
# If you need to add a new location for the AL market please add the field here in the correct aphabetizing
[void] $listBoxSubCat.Items.Add("AD")
[void] $listBoxSubCat.Items.Add("Avamar")
[void] $listBoxSubCat.Items.Add("BMC Service Core")
[void] $listBoxSubCat.Items.Add("IQ/RQ")
[void] $listBoxSubCat.Items.Add("Sharepoint")
[void] $listBoxSubCat.Items.Add("SolarWinds")
[void] $listBoxSubCat.Items.Add("Windows 7")
[void] $listBoxSubCat.Items.Add("Windows Server 2008 R2")
[void] $listBoxSubCat.Items.Add("Windows Server 2012 R2")
}

ElseIf ($comboBoxCat.text -eq "Information Request")
{
$listBoxSubCat.Items.Clear()
# If you need to add a new location for the AL market please add the field here in the correct aphabetizing
[void] $listBoxSubCat.Items.Add("AD")
[void] $listBoxSubCat.Items.Add("Email")
[void] $listBoxSubCat.Items.Add("Legal")
[void] $listBoxSubCat.Items.Add("SQL data")
[void] $listBoxSubCat.Items.Add("Video Surveillance")

}

 $OpsDashForm.Controls.Add($listBoxSubCat)
}
$comboBoxCat.add_SelectedIndexChanged($comboBoxCategory_SelectedIndexChanged)

    # Add input Label for Ticket number
	$inputboxlabelSAM = new-object System.Windows.Forms.Label
	$inputboxlabelSAM.Location = new-object System.Drawing.Size(20,35)
	$inputboxlabelSAM.size = new-object System.Drawing.Size(90,20)
	$inputboxlabelSAM.Text = "SAM account"
	$OpsDashForm.controls.Add($inputboxlabelSAM)
	# Add input Label
	$inputboxlabelMarketlocation = new-object System.Windows.Forms.Label
	$inputboxlabelMarketlocation.Location = new-object System.Drawing.Size(220,30)
	$inputboxlabelMarketlocation.size = new-object System.Drawing.Size(150,20)
	$inputboxlabelMarketlocation.Text = "Market/Location (KXBH)"
	$OpsDashForm.controls.Add($inputboxlabelMarketlocation)
    # Add Computer textbox
    $inputboxPC = New-Object System.Windows.Forms.TextBox
	$inputboxPC.Location = New-Object System.Drawing.Size(420,50)
	$inputboxPC.Size = New-Object System.Drawing.Size(150,20)
    $inputboxPC.TabIndex = 4
	$OpsDashForm.Controls.Add($inputboxPC)
	# Add Additional Comments textbox
	$Additionalcomments = New-Object System.Windows.Forms.TextBox
	$Additionalcomments.Location = New-Object System.Drawing.Size(28,340)
	$Additionalcomments.Size = New-Object System.Drawing.Size(300,200)
    $Additionalcomments.ScrollBars = "Vertical"
	$Additionalcomments.MultiLine = $True
    $Additionalcomments.TabIndex = 4
	$OpsDashForm.Controls.Add($Additionalcomments)
	# Add Additional Comments Label
	$Additionalcommentslabel = new-object System.Windows.Forms.Label
	$Additionalcommentslabel.Location = new-object System.Drawing.Size(28,315)
	$Additionalcommentslabel.size = new-object System.Drawing.Size(300,25)
	$Additionalcommentslabel.Text = "Ticket Journal"
	$OpsDashForm.controls.Add($Additionalcommentslabel)
    # Add Note text box
    $NoteTextbox = New-Object System.Windows.Forms.TextBox
    $NoteTextbox.Location = New-Object System.Drawing.Size(345,340)
    $NoteTextbox.Size = New-Object System.Drawing.Size(425,200)
    $NoteTextbox.ScrollBars = "Vertical"
    $NoteTextbox.MultiLine = $True
    $OpsDashForm.Controls.Add($NoteTextbox)
    # Add Lable for Notes
	$NoteTextbox = new-object System.Windows.Forms.Label
	$NoteTextbox.Location = new-object System.Drawing.Size(345,315)
	$NoteTextbox.size = new-object System.Drawing.Size(300,25)
	$NoteTextbox.Text = "Notes only (Will not be in Ticket)"
	$OpsDashForm.controls.Add($NoteTextbox)
    # Add PC name label
	$inputboxPCLabel = new-object System.Windows.Forms.Label
	$inputboxPCLabel.Location = new-object System.Drawing.Size(420,35)
	$inputboxPCLabel.size = new-object System.Drawing.Size(100,20)
	$inputboxPCLabel.Text = "PC Name"
	$OpsDashForm.controls.Add($inputboxPCLabel)
	# Add Dropdown box for status
    $TicketStatus  = new-object System.Windows.Forms.ComboBox
    $TicketStatus.Location = new-object System.Drawing.Size(620,50)
    $TicketStatus.Size = new-object System.Drawing.Size(150,20)
    $TicketStatus.Text = "In Progress"
    $OpsDashForm.controls.Add($TicketStatus)
    # Arrow label
    $AarrowLabel  = new-object System.Windows.Forms.Label
    $AarrowLabel.Location = new-object System.Drawing.Size(500,253)
    $AarrowLabel.Size = new-object System.Drawing.Size(150,20)
    $AarrowLabel.Font = "System.Drawing,12"
    $AarrowLabel.Text = "<---"
    $OpsDashForm.controls.Add($AarrowLabel)
    # copyright label
    $copyrightLabel  = new-object System.Windows.Forms.Label
    $copyrightLabel.Location = new-object System.Drawing.Size(35,545)
    $copyrightLabel.Size = new-object System.Drawing.Size(200,200)
    $copyrightLabel.Text = ""
    $OpsDashForm.controls.Add($copyrightLabel)
    # Category Label
    $CategoryLabel  = new-object System.Windows.Forms.Label
    $CategoryLabel.Location = new-object System.Drawing.Size(560,235)
    $CategoryLabel.Size = new-object System.Drawing.Size(75,20)
    $CategoryLabel.Text = "Category"
    $OpsDashForm.controls.Add($CategoryLabel)
    # Subcategory Label
    $SubcategoryLabel  = new-object System.Windows.Forms.Label
    $SubcategoryLabel.Location = new-object System.Drawing.Size(680,235)
    $SubcategoryLabel.Size = new-object System.Drawing.Size(75,18)
    $SubcategoryLabel.Text = "Subcategory"
    $OpsDashForm.controls.Add($SubcategoryLabel)
    # Line Label
    $lineLabel  = new-object System.Windows.Forms.Label
    $lineLabel.Location = new-object System.Drawing.Size(20,230)
    $lineLabel.Size = new-object System.Drawing.Size(495,18)
    $lineLabel.Text = "__________________________________________________________________________________________________"
    $OpsDashForm.controls.Add($lineLabel)
    # Add Ticket Status Comment Label
    $TicketStatusLabel = new-object System.Windows.Forms.Label
	$TicketStatusLabel.Location = new-object System.Drawing.Size(620,35)
	$TicketStatusLabel.size = new-object System.Drawing.Size(75,25)
	$TicketStatusLabel.Text = "Ticket Status"
	$OpsDashForm.controls.Add($TicketStatusLabel)
    [array]$DropDownArray = "In Progress", "Completed"
     ForEach ($Item in $DropDownArray) {
     [void] $TicketStatus.Items.Add($Item)
    }
    # output textbox
	$outputBox = New-Object System.Windows.Forms.Label
	$outputBox.Location = New-Object System.Drawing.Size(785,50)
	$outputBox.Size = New-Object System.Drawing.Size(175,50)
    $outputBox.ForeColor = "Green"
    $outputBox.Font = "Microsoft Sans Serif, 10pt, style=Bold"
	$OpsDashForm.Controls.Add($outputBox)

    # Add Label for info
    $ContactLabel = new-object System.Windows.Forms.Label
	$ContactLabel.Location = new-object System.Drawing.Size(800,100)
	$ContactLabel.size = new-object System.Drawing.Size(160,100)
	$ContactLabel.Text = ""
	$OpsDashForm.controls.Add($ContactLabel)
    # end text fields
    # FieldTech textbox
	$FieldTechInfo = New-Object System.Windows.Forms.Label
	$FieldTechInfo.Location = New-Object System.Drawing.Size(800,250)
	$FieldTechInfo.Size = New-Object System.Drawing.Size(150,50)
    $FieldTechInfo.ForeColor = "Green"
    $FieldTechInfo.Font = "Microsoft Sans Serif, 10pt, style=Bold"
	$OpsDashForm.Controls.Add($FieldTechInfo)


    $PrinterLink = New-Object System.Windows.Forms.LinkLabel
    $PrinterLink.Location = New-Object System.Drawing.Size(415,208)
    $PrinterLink.Size = New-Object System.Drawing.Size(80,20)
    $PrinterLink.LinkColor = "BLUE"
    $PrinterLink.ActiveLinkColor = "RED"
    $PrinterLink.Text = "Documentation"
    $PrinterLink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($PrinterLink)
    #
    $Authcodelink = New-Object System.Windows.Forms.LinkLabel
    $Authcodelink.Location = New-Object System.Drawing.Size(25,132)
    $Authcodelink.Size = New-Object System.Drawing.Size(80,20)
    $Authcodelink.LinkColor = "BLUE"
    $Authcodelink.ActiveLinkColor = "RED"
    $Authcodelink.Text = "Documentation"
    $Authcodelink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($Authcodelink)
    #
    $SetTimeLink = New-Object System.Windows.Forms.LinkLabel
    $SetTimeLink.Location = New-Object System.Drawing.Size(155,132)
    $SetTimeLink.Size = New-Object System.Drawing.Size(80,20)
    $SetTimeLink.LinkColor = "BLUE"
    $SetTimeLink.ActiveLinkColor = "RED"
    $SetTimeLink.Text = "Documentation"
    $SetTimeLink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($SetTimeLink)
    #
    $TermLink = New-Object System.Windows.Forms.LinkLabel
    $TermLink.Location = New-Object System.Drawing.Size(285,132)
    $TermLink.Size = New-Object System.Drawing.Size(80,20)
    $TermLink.LinkColor = "BLUE"
    $TermLink.ActiveLinkColor = "RED"
    $TermLink.Text = "Documentation"
    $TermLink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($TermLink)
    #
    $TechDispatchLink = New-Object System.Windows.Forms.LinkLabel
    $TechDispatchLink.Location = New-Object System.Drawing.Size(415,132)
    $TechDispatchLink.Size = New-Object System.Drawing.Size(80,20)
    $TechDispatchLink.LinkColor = "BLUE"
    $TechDispatchLink.ActiveLinkColor = "RED"
    $TechDispatchLink.Text = "Documentation"
    $TechDispatchLink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($TechDispatchLink)
    #
    $InternetLink = New-Object System.Windows.Forms.LinkLabel
    $InternetLink.Location = New-Object System.Drawing.Size(545,132)
    $InternetLink.Size = New-Object System.Drawing.Size(80,20)
    $InternetLink.LinkColor = "BLUE"
    $InternetLink.ActiveLinkColor = "RED"
    $InternetLink.Text = "Documentation"
    $InternetLink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($InternetLink)
    #
    $ReTokenLink = New-Object System.Windows.Forms.LinkLabel
    $ReTokenLink.Location = New-Object System.Drawing.Size(675,132)
    $ReTokenLink.Size = New-Object System.Drawing.Size(80,20)
    $ReTokenLink.LinkColor = "BLUE"
    $ReTokenLink.ActiveLinkColor = "RED"
    $ReTokenLink.Text = "Documentation"
    $ReTokenLink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($ReTokenLink)
    #
    $ADUnlockLink = New-Object System.Windows.Forms.LinkLabel
    $ADUnlockLink.Location = New-Object System.Drawing.Size(25,207)
    $ADUnlockLink.Size = New-Object System.Drawing.Size(80,20)
    $ADUnlockLink.LinkColor = "BLUE"
    $ADUnlockLink.ActiveLinkColor = "RED"
    $ADUnlockLink.Text = "Documentation"
    $ADUnlockLink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($ADUnlockLink)
    #Other maybe
    $OtherLink = New-Object System.Windows.Forms.LinkLabel
    $OtherLink.Location = New-Object System.Drawing.Size(25,280)
    $OtherLink.Size = New-Object System.Drawing.Size(80,20)
    $OtherLink.LinkColor = "BLUE"
    $OtherLink.ActiveLinkColor = "RED"
    $OtherLink.Text = "Documentation"
    $OtherLink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($OtherLink)
    #
    $RQLink = New-Object System.Windows.Forms.LinkLabel
    $RQLink.Location = New-Object System.Drawing.Size(155,207)
    $RQLink.Size = New-Object System.Drawing.Size(80,20)
    $RQLink.LinkColor = "BLUE"
    $RQLink.ActiveLinkColor = "RED"
    $RQLink.Text = "Documentation"
    $RQLink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($RQLink)
    #
    $VerifoneLink = New-Object System.Windows.Forms.LinkLabel
    $VerifoneLink.Location = New-Object System.Drawing.Size(285,207)
    $VerifoneLink.Size = New-Object System.Drawing.Size(80,20)
    $VerifoneLink.LinkColor = "BLUE"
    $VerifoneLink.ActiveLinkColor = "RED"
    $VerifoneLink.Text = "Documentation"
    $VerifoneLink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($VerifoneLink)
    #
    $GlobalLink = New-Object System.Windows.Forms.LinkLabel
    $GlobalLink.Location = New-Object System.Drawing.Size(415,280)
    $GlobalLink.Size = New-Object System.Drawing.Size(80,20)
    $GlobalLink.LinkColor = "BLUE"
    $GlobalLink.ActiveLinkColor = "RED"
    $GlobalLink.Text = "Documentation"
    $GlobalLink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($GlobalLink)
    #
    $EventsLink = New-Object System.Windows.Forms.LinkLabel
    $EventsLink.Location = New-Object System.Drawing.Size(285,280)
    $EventsLink.Size = New-Object System.Drawing.Size(80,20)
    $EventsLink.LinkColor = "BLUE"
    $EventsLink.ActiveLinkColor = "RED"
    $EventsLink.Text = "Documentation"
    $EventsLink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($EventsLink)
    #
    $SoftLink = New-Object System.Windows.Forms.LinkLabel
    $SoftLink.Location = New-Object System.Drawing.Size(155,280)
    $SoftLink.Size = New-Object System.Drawing.Size(80,20)
    $SoftLink.LinkColor = "BLUE"
    $SoftLink.ActiveLinkColor = "RED"
    $SoftLink.Text = "Documentation"
    $SoftLink.add_Click({[system.Diagnostics.Process]::start("")})
    $OpsDashForm.controls.Add($SoftLink)

    ##### Web Browser


    $webbrowser2 = New-Object 'System.Windows.Forms.WebBrowser'
    $webbrowser2.IsWebBrowserContextMenuEnabled = $false
    $webbrowser2.Location = '960, 50'
	$webbrowser2.Name = "webbrowser2"
	$webbrowser2.Size = '315, 700'
	$webbrowser2.navigate("")
	$OpsDashForm.Controls.Add($webbrowser2)



    # Start button
    $StartButton = New-Object System.Windows.Forms.Button
	$StartButton.Location = New-Object System.Drawing.Size(800,10)
	$StartButton.Size = New-Object System.Drawing.Size(90,30)
	$StartButton.Text = "Start"
	$OpsDashForm.Controls.Add($StartButton)
    # Start Timer
    $Stopwatch = New-Object System.Diagnostics.Stopwatch
    $Formtimer = New-Object System.Windows.Forms.Timer -Property @{Interval = 2}
    $Timer = New-Object System.Windows.Forms.TextBox
    $Timer.Location = New-Object System.Drawing.Size(960,1)
    $Timer.Size = New-Object System.Drawing.Size(220,30)
    $Timer.Text = $stopwatch.Elapsed
    $Timer.Font = New-Object System.Drawing.Font("Verdana",23.6,[System.Drawing.FontStyle]::Bold)
    $Timer.ReadOnly = $true
    $OpsDashForm.Controls.Add($Timer)
        $SB_Button2 = {
        $Stopwatch.Reset()
        $Stopwatch.Start()
        $Formtimer.Start()}
        $SB_OnTick = {
        $Timer.Text = "{0:hh}:{0:mm}:{0:ss}.{0:ff}" -f $stopwatch.Elapsed
    }
        $StartButton.Add_Click($SB_Button2)

        $SB_Button4 = {
        $Stopwatch.Stop()
        $Formtimer.Stop() }

        $Formtimer.Add_tick($SB_OnTick)

	# Auth Code buttons
	$AuthCodeButton = New-Object System.Windows.Forms.Button
	$AuthCodeButton.Location = New-Object System.Drawing.Size(20,100)
	$AuthCodeButton.Size = New-Object System.Drawing.Size(90,30)
	$AuthCodeButton.Text = "Auth Code"
    $AuthCodeButton.add_Click({[system.Diagnostics.Process]::start("")})
    $AuthCodeButton.Add_Click({Clear_Status})
    $AuthCodeButton.Add_Click({FieldTech})
	$AuthCodeButton.Add_Click({Auth_Code})
    $AuthCodeButton.Add_Click({Clear_Fields})
    $AuthCodeButton.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($AuthCodeButton)
    #
	$GetField = New-Object System.Windows.Forms.Button
	$GetField.Location = New-Object System.Drawing.Size(800,210)
	$GetField.Size = New-Object System.Drawing.Size(90,30)
	$GetField.Text = "Get Field Tech"
    $GetField.Add_Click({FieldTech})
    $GetField.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($GetField)
    # Set time on PC button
	$SetTime = New-Object System.Windows.Forms.Button
	$SetTime.Location = New-Object System.Drawing.Size(150,100)
	$SetTime.Size = New-Object System.Drawing.Size(90,30)
	$SetTime.Text = "Set Time"
    $SetTime.Add_Click({Clear_Status})
    $SetTime.Add_Click({FieldTech})
	$SetTime.Add_Click({Set_Time})
    $SetTime.Add_Click({Clear_Fields})
    $SetTime.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($SetTime)
    # RQ button
	$RQButton = New-Object System.Windows.Forms.Button
	$RQButton.Location = New-Object System.Drawing.Size(150,175)
	$RQButton.Size = New-Object System.Drawing.Size(90,30)
	$RQButton.Text = "RQ"
    $RQButton.Add_Click({Clear_Status})
    $RQButton.Add_Click({FieldTech})
	$RQButton.Add_Click({RQ_Code})
    $RQButton.Add_Click({Clear_Fields})
    $RQButton.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($RQButton)
    # Verifone button
	$VerifoneButton = New-Object System.Windows.Forms.Button
	$VerifoneButton.Location = New-Object System.Drawing.Size(280,175)
	$VerifoneButton.Size = New-Object System.Drawing.Size(90,30)
	$VerifoneButton.Text = "Verifone"
    $VerifoneButton.Add_Click({Clear_Status})
    $VerifoneButton.Add_Click({FieldTech})
	$VerifoneButton.Add_Click({Verifone})
    $VerifoneButton.Add_Click({Clear_Fields})
    $VerifoneButton.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($VerifoneButton)
	# Printer button
$PartyButton = New-Object System.Windows.Forms.Button
$PartyButton.Location = New-Object System.Drawing.Size(870,520)
$PartyButton.Size = New-Object System.Drawing.Size(15,15)
$PartyButton.Text = ":balloon:"
	$PartyButton.add_Click({[system.Diagnostics.Process]::start("https://www.youtube.com/watch?v=dQw4w9WgXcQ")})
$OpsDashForm.Controls.Add($PartyButton)
    # Printer button
	$PrinterButton = New-Object System.Windows.Forms.Button
	$PrinterButton.Location = New-Object System.Drawing.Size(410,175)
	$PrinterButton.Size = New-Object System.Drawing.Size(90,30)
	$PrinterButton.Text = "Printer"
    $PrinterButton.Add_Click({Clear_Status})
    $PrinterButton.Add_Click({FieldTech})
	$PrinterButton.Add_Click({Printer})
    $PrinterButton.Add_Click({Clear_Fields})
    $PrinterButton.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($PrinterButton)
  # Information button
$InformationButton = New-Object System.Windows.Forms.Button
$InformationButton.Location = New-Object System.Drawing.Size(540,175)
$InformationButton.Size = New-Object System.Drawing.Size(90,30)
$InformationButton.Text = "Info Request"
  $InformationButton.Add_Click({Clear_Status})
  $InformationButton.Add_Click({FieldTech})
$InformationButton.Add_Click({Information})
  $InformationButton.Add_Click({Clear_Fields})
  $InformationButton.Add_Click($SB_Button4)
$OpsDashForm.Controls.Add($InformationButton)
    # Termination button
	$TermButton = New-Object System.Windows.Forms.Button
	$TermButton.Location = New-Object System.Drawing.Size(280,100)
	$TermButton.Size = New-Object System.Drawing.Size(90,30)
	$TermButton.Text = "Termination"
    $TermButton.Add_Click({Clear_Status})
    $TermButton.Add_Click({FieldTech})
	$TermButton.Add_Click({Term_button})
    $TermButton.Add_Click({Clear_Fields})
    $TermButton.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($TermButton)
    # Tech Dispatch button
	$TechDispatch = New-Object System.Windows.Forms.Button
	$TechDispatch.Location = New-Object System.Drawing.Size(410,100)
	$TechDispatch.Size = New-Object System.Drawing.Size(90,30)
	$TechDispatch.Text = "Tech Dispatch"
    $TechDispatch.Add_Click({Clear_Status})
    $TechDispatch.Add_Click({FieldTech})
	$TechDispatch.Add_Click({Tech_Dispatch})
    $TechDispatch.Add_Click({Clear_Fields})
    $TechDispatch.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($TechDispatch)
    # Hardware button
	$HardwareButton = New-Object System.Windows.Forms.Button
	$HardwareButton.Location = New-Object System.Drawing.Size(20,250)
	$HardwareButton.Size = New-Object System.Drawing.Size(90,30)
	$HardwareButton.Text = "Hardware"
    $HardwareButton.Add_Click({Clear_Status})
    $HardwareButton.Add_Click({FieldTech})
	$HardwareButton.Add_Click({Hardware})
    $HardwareButton.Add_Click({Clear_Fields})
    $HardwareButton.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($HardwareButton)
    # Software button
	$SoftwareButton = New-Object System.Windows.Forms.Button
	$SoftwareButton.Location = New-Object System.Drawing.Size(150,250)
	$SoftwareButton.Size = New-Object System.Drawing.Size(90,30)
	$SoftwareButton.Text = "Software"
    $SoftwareButton.Add_Click({Clear_Status})
    $SoftwareButton.Add_Click({FieldTech})
	$SoftwareButton.Add_Click({Software})
    $SoftwareButton.Add_Click({Clear_Fields})
    $SoftwareButton.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($SoftwareButton)
    # Re-Token button
	$ReTokenButton = New-Object System.Windows.Forms.Button
	$ReTokenButton.Location = New-Object System.Drawing.Size(670,100)
	$ReTokenButton.Size = New-Object System.Drawing.Size(90,30)
	$ReTokenButton.Text = "Re-Token"
    $ReTokenButton.Add_Click({Clear_Status})
    $ReTokenButton.Add_Click({FieldTech})
	$ReTokenButton.Add_Click({Re_Token})
    $ReTokenButton.Add_Click({Clear_Fields})
    $ReTokenButton.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($ReTokenButton)
    # Unlock AD account
	$ADUnlockButton = New-Object System.Windows.Forms.Button
	$ADUnlockButton.Location = New-Object System.Drawing.Size(20,175)
	$ADUnlockButton.Size = New-Object System.Drawing.Size(90,30)
	$ADUnlockButton.Text = "AD"
    $ADUnlockButton.Add_Click({Clear_Status})
    $AuthCodeButton.Add_Click({FieldTech})
	$ADUnlockButton.Add_Click({AD_Unlock})
    $ADUnlockButton.Add_Click({Clear_Fields})
    $ADUnlockButton.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($ADUnlockButton)
    # Internet button
	$InternetButton = New-Object System.Windows.Forms.Button
	$InternetButton.Location = New-Object System.Drawing.Size(540,100)
	$InternetButton.Size = New-Object System.Drawing.Size(90,30)
	$InternetButton.Text = "Internet"
    $InternetButton.Add_Click({Clear_Status})
    $InternetButton.Add_Click({FieldTech})
	$InternetButton.Add_Click({Internet})
    $InternetButton.Add_Click({Clear_Fields})
    $InternetButton.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($InternetButton)
    # Global button
	$GlobalButton = New-Object System.Windows.Forms.Button
	$GlobalButton.Location = New-Object System.Drawing.Size(410,250)
	$GlobalButton.Size = New-Object System.Drawing.Size(90,30)
	$GlobalButton.Text = "Global"
    $GlobalButton.Add_Click({Clear_Status})
    $GlobalButton.Add_Click({FieldTech})
	$GlobalButton.Add_Click({Global})
    $GlobalButton.Add_Click({Clear_Fields})
    $GlobalButton.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($GlobalButton)
    # Events button
	$EventButton = New-Object System.Windows.Forms.Button
	$EventButton.Location = New-Object System.Drawing.Size(280,250)
	$EventButton.Size = New-Object System.Drawing.Size(90,30)
	$EventButton.Text = "Events"
    $EventButton.Add_Click({Clear_Status})
    $EventButton.Add_Click({FieldTech})
	$EventButton.Add_Click({Events})
    $EventButton.Add_Click({Clear_Fields})
    $EventButton.Add_Click($SB_Button4)
	$OpsDashForm.Controls.Add($EventButton)
    # check button
	$CheckButton = New-Object System.Windows.Forms.Button
	$CheckButton.Location = New-Object System.Drawing.Size(110,25)
	$CheckButton.Size = New-Object System.Drawing.Size(50,20)
	$CheckButton.Text = "Check"
    $CheckButton.TabIndex = 1
    $CheckButton.add_click({
    $grid = $ad_grid
    $lnquery = "*"+$SAMBox.Text.ToString()+"*"
    search_contact_ad($lnquery)
    })

$OpsDashForm.Controls.Add($CheckButton)
    # Grid for AD info
    $ad_grid = New-Object Windows.Forms.DataGridview
    $ad_grid.DataBindings.DefaultDataSourceUpdateMode = 0
    $ad_grid.Name = "grouplist"
    $ad_grid.DataMember = ""
    $ad_grid.TabIndex = 6
    $ad_grid.Location = New-Object Drawing.Point 790,300
    $ad_grid.Size = New-Object Drawing.Point 150,200
    $ad_grid.readonly = $true
    $ad_grid.AutoSizeColumnsMode = 'AllCells'
    $ad_grid.SelectionMode = 'FullRowSelect'
    $ad_grid.MultiSelect = $false
    $ad_grid.RowHeadersVisible = $false
    $ad_grid.allowusertoordercolumns = $true
    $OpsDashForm.Controls.Add($ad_grid)

    # Ping button
	$PingButton = New-Object System.Windows.Forms.Button
	$PingButton.Location = New-Object System.Drawing.Size(150,175)
	$PingButton.Size = New-Object System.Drawing.Size(90,30)
	$PingButton.Text = "Ping"
    $PingButton.Add_Click({Ping})
    # Add Exit button
	$ExitButton = New-Object System.Windows.Forms.Button
	$ExitButton.Location = New-Object System.Drawing.Size(790,505)
	$ExitButton.Size = New-Object System.Drawing.Size(80,30)
	$ExitButton.Text = "Quit"
	$ExitButton.Add_Click({exit-form})
	$OpsDashForm.Controls.Add($ExitButton)
	# Set keys to call button pushes.
	$OpsDashForm.KeyPreview = $True
	$OpsDashForm.Add_KeyDown({if ($_.KeyCode -eq "Escape")
		{exit-form}})

	# end buttons
	# Activate form and controls.
	$OpsDashForm.Add_Shown({$OpsDashForm.Select()})
	[void] $OpsDashForm.ShowDialog()
