param (
    # Who to send the email to.
    [Array]$Addresses = $( throw "You have to supply target email addresses."; ),

    # The SMTP server to use for sending.
    [string]$SmtpServerAddress = $( throw "You have to supply an SMTP server."; ),
    # The username to use to log into the SMTP server.
    [string]$SmtpUsername,
    # If username is given, also ask for a password.
    [System.Security.SecureString]$SmtpPassword = $( if( $SmtpUsername ) { Read-Host -AsSecureString "Input SMTP password for $SmtpUsername, please"; } ),

    # The address of the sender. For example "Your Company Bulk Mailer <bulk@yourcompany.com>"
    [string]$SenderAddress = $( throw "You have to supply a sender address."; ),
    # The optional reply-to address.
    [string]$ReplyToAddress = $SenderAddress,
    # The subject of the mail.
    [string]$Subject = $( throw "You have to supply a subject."; ),

    # Is the mail body in HTML?
    [switch]$Html,
    # How many seconds to count down before actually starting to send mail.
    [int]$Countdown = 5,
    # If enabled, the script will run through everything excepting actually sending mail.
    [switch]$DryRun
)

Write-Host "HARTWIG Bulk Emailer";
Write-Host "====================";

# Construct the body of the email
Write-Host -NoNewLine "> Constructing email body...";
$body = "";
foreach( $line in $input ) {
    $body += $line + "`n";
}
Write-Host "Done.  ";

if( $Html ) {
    Write-Host -NoNewline "> Determining embedded content...";
    $cidRegex = [regex]"cid:([^`"]+)";
    $Matches = $cidRegex.Matches( $body );
    Write-Host "Done. Found" $Matches.Count "elements";
    Write-Host ">";

    if( $Matches.Count -gt 0 ) {
        $alternateView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString( $body, $null, "text/html" );
    
        foreach( $contentId in $Matches ) {
            $id = $contentId.Groups[1].Value;
            if( Test-Path $id ) {
                $filePath =  $(get-location).Path + "\" + $id;
                
                Write-Host ">    - ``"$id"`` **found** (will be used in email)  ";
                Write-Host ">      → ``$filePath``";
                
                $contentElement = New-Object System.Net.Mail.LinkedResource( $filePath );
                $contentElement.ContentId = $id;

                $alternateView.LinkedResources.Add( $contentElement );

            } else {
                Write-Host ">    - ``"$id"`` missing";
            }
        }   
    }
}
Write-Host "";

Write-Host "Email Summary";
Write-Host "-------------";
Write-Host "- FROM    : **$SenderAddress**";
Write-Host "- REPLY-TO: **$ReplyToAddress**";
Write-Host "- SUBJECT : **$Subject**";
Write-Host "- BODY    :" $body.Length "characters" @{$true="(HTML)";$false="(Text)"}[$Html -eq $True];
Write-Host "";

# Start a countdown to allow for canceling of operation.
Write-Host "> Sending to" $addresses.Count "addresses in $Countdown seconds. Press Ctrl+C to cancel operation.";
for( $percentComplete = $Countdown; $percentComplete -ge 0; --$percentComplete ) {
  Write-Progress -Activity “Countdown” -PercentComplete $percentComplete -Status “Counting down, press CTRL+C to kill this script now if required!”;
  Sleep -Seconds 1;
}
Write-Host "";

# Set up counter
$sendingFailed = @();
$sendingSucceeded = @();

# Prepare the client
$SmtpClient = new-object Net.Mail.SmtpClient( $SmtpServerAddress );
if( $SmtpUsername -and $SmtpPassword ) {
    # Construct credentials to log on to the server.
    $SmtpClient.Credentials = 
        New-Object System.Net.NetworkCredential( 
            $SmtpUsername, 
            [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $SmtpPassword ) )
        );
}

# Send mail to each address
Write-Host "Operation Log";
Write-Host "-------------";
foreach( $address in $addresses ) {
    Write-Host -NoNewline "- Sending email to **$address**...";

    $mailMessage = New-Object Net.Mail.MailMessage
    $mailMessage.From = $SenderAddress;

    Try {
        $mailMessage.To.Add( $address );
    } Catch [System.FormatException] {
        Write-Host "**Failed!** The email address is malformed. _($_.Exception.Message)_";
        $sendingFailed += $address;
        Continue;
    }

    Try {
        $mailMessage.ReplyToList.Add( $ReplyToAddress );
    } Catch [System.FormatException] {
        Write-Host "**Failed!** The reply-to email address is malformed. _($_.Exception.Message)_";
        Write-Host "**Aborting operation!**"
        $sendingFailed += $address;
        Break;
    }
    
    $mailMessage.Subject = $Subject;
    $mailMessage.Body = $body;
    if( $alternateView ) {
        $mailMessage.AlternateViews.Add( $alternateView );
    }
    $mailMessage.IsBodyHtml = $( $Html );

    Try {
        if( !$DryRun ) {
            $SmtpClient.Send( $mailMessage );
        }
    } Catch [System.Net.Mail.SmtpFailedRecipientException] {
        Write-Host "**Failed!** Sending email failed. _($_.Exception.Message)_";
        $sendingFailed += $address;
        Continue;
    }

    $sendingSucceeded += $address;
    Write-Host "Done.";
}

Write-Host "";
Write-Host "> Operation completed." $sendingSucceeded.Count "successful," $sendingFailed.Count "failed";

if( $sendingFailed.Count -gt 0 ) {
    Write-Host "";
    Write-Host "Failed Addresses";
    Write-Host "----------------";

    foreach( $address in $sendingFailed ) {
        Write-Host "- $address";
    }
}
