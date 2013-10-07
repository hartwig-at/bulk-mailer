bulk-mailer
===========

How to use
----------

Grab the content of `email.html` and pipe it into the script.

    Get-Content .\email.html | .\Send-Emails.ps1 

Grab the addresses from `addresses.txt`. One address per line. Format should be either `user@host.com` or `User Name <user@host.com>`.

    -Addresses (Get-Content "addresses.txt") 
  
Declare the SMTP server to use.
  
    -SmtpServerAddress "localhost"

Set the SMTP username to log into the SMTP server. This parameter is optional. If it is given, you'll be prompted for a password at runtime.

    -SmtpUsername "OliverSalzburg" 

The address to use for the sender.

    -SenderAddress "HARTWIG Bulk Mailer <noreply@hartwig-at.de>"

Who should mails go to when someone replies to a received mail? If omitted, the `SenderAddress` will be used.

    -ReplyToAddress "Customer INC. <info@customer.com>"

The subject of the mail.

    -Subject "A very special email for you"

The email is in HTML format.

    -Html
