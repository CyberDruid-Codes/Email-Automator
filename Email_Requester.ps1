Add-Type -AssemblyName PresentationFramework

#Creating The XML commands for the UI 
[xml]$Form  = @"
  <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
   Title="Email Automation" Height="500" Width="500">

    <Canvas Background="DarkBlue">

     <Button Name="TheButton" Height="40" Width="100" Content='Click to Start' Margin="200,250,0,0" VerticalAlignment="Top">
       <Button.Effect>
         <DropShadowEffect/>
       </Button.Effect>
     </Button>

     <Label Name="Title" Content="Access Automation" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="130,30,0,0" Height="50" Width="250" FontSize="20" Foreground="White"/>

     <Label Name="NameLabel" Content="Write the Requester name here:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="150,165,0,0" Height="30" Width="210" FontSize="14" Foreground="White"/>

     <TextBox Name="TextName" Height="23" Margin="150,200,0,0" Width="200" TextWrapping="Wrap" Text="" VerticalAlignment="Top" BorderBrush="#FFF96816"/>

     <Label Name="CommentName" Content="Add Extra Comments:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="165,105,0,0" Height="30" Width="210" FontSize="14" Foreground="White"/>

     <TextBox Name="Comment" Height="23" Margin="150,140,0,0" Width="200" TextWrapping="Wrap" Text="" VerticalAlignment="Top" BorderBrush="#FFF96816"/>

    </Canvas>
  </Window>
"@

#Converting The XML format to be readable by powershell
$NR=(New-Object System.Xml.XmlNodeReader $Form)
$Win=[Windows.Markup.XamlReader]::Load( $NR )

#Used to connect the XML UI with the script by finding it after the ID
$UserID = $Win.FindName("TextName")
$Start = $Win.FindName("TheButton")
$ticketNr = $Win.FindName("Comment")

#Making an instance of the variable so it retrieves the textbox input
$script:UserName=" "
$script:Info=" "
 
#Adding an on-click that parses the input in the box
$Start.Add_Click({
$script:UserName=$UserID.text
$script:Info=$ticketNr.text
$Win.Close()})

#Displays the dialog box
$Win.Showdialog()

###
#Email Script. This will be the actual email, which can be modified. 
$From = "<From Email Address>"
$To = "<To Email Address>"
$Subject = "Access For User $UserName"
$Body = "Hello,
 
I need Access for this person:
 
           $UserName
          -Comments: $Info

 
Thank you,
"
$SMTPServer = "smtp-mail.outlook.com"
$SMTPPort = "587"
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential (Get-Credential)
###
