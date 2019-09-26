
#+-------------------------------------------------------------------+
#| = : = : = : = : = : = : = : = : = : = : = : = : = : = : = : = : = |
#|{>/-------------------------------------------------------------\<}|
#|: | Author:  Philippe-Alexandre Munch                           | :|
#| :| Email:   --------------------------------                   |: |
#|: | Purpose: Create_User :)  in GUI Version                     | :|
#| :|                                                             |: |
#|: |                      					  | :|
#| :|                                                             |: |
#|: |         		Date:15-08-2018                           | :|
#|: |                            16:48                            |: |
#| :| 	/^(o.o)^\    Version: 1.0           	                  | :|
#|{>\-------------------------------------------------------------/<}|
#| = : = : = : = : = : = : = : = : = : = : = : = : = : = : = : = : = |
#+-------------------------------------------------------------------+


#################
# LOAD ASSEMBLY #
#################
Add-Type -AssemblyName Microsoft.VisualBasic
[void][reflection.assembly]::LoadWithPartialName("System.Windows.Forms")
[reflection.assembly]::LoadWithPartialName( "System.Windows.Forms")
[reflection.assembly]::LoadWithPartialName( "System.Drawing")
[System.Windows.Forms.Application]::EnableVisualStyles();

########
# FORM #
########
$Icon = New-Object system.drawing.icon ("")
$form = New-Object System.Windows.Forms.Form
$form.text = "AD_USER_CREATION_GUI"
$form.Size = New-Object System.Drawing.Size(500,500)
$form.Icon = $Icon

##########
# Object #
##########
$button1 = New-Object System.Windows.Forms.Button
$Okbutton = New-Object System.Windows.Forms.Button
$label1 = New-Object System.Windows.Forms.Label
$label2 = New-Object System.Windows.Forms.Label
$label3 = New-Object System.Windows.Forms.Label
$label4 = New-Object System.Windows.Forms.Label
$label5 = New-Object System.Windows.Forms.Label
$label6 = New-Object System.Windows.Forms.Label
$label7 = New-Object System.Windows.Forms.Label
$label8 = New-Object System.Windows.Forms.Label
$textbox1 = New-Object System.Windows.Forms.TextBox
$textbox2 = New-Object System.Windows.Forms.TextBox
$textbox3 = New-Object System.Windows.Forms.TextBox
$textbox4 = New-Object System.Windows.Forms.TextBox
$textbox5 = New-Object System.Windows.Forms.TextBox
$textbox6 = New-Object System.Windows.Forms.TextBox
$textbox7 = New-Object System.Windows.Forms.TextBox
$combobox = New-Object System.Windows.Forms.ComboBox

########################
# Visual Basic Request #
########################
$Givename = [Microsoft.VisualBasic.Interaction]::InputBox("Enter User First Name","First Name")
$Surname = [Microsoft.VisualBasic.Interaction]::InputBox("Enter User Last Name","Last Name")
$Name = [Microsoft.VisualBasic.Interaction]::InputBox("Complet User Name","Complet User Name",$Givename + " " + $Surname)
$UserPrincipalName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter User Email Address","E-Mail",$Givename + "." + $Surname)
$SamAccountName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter User Trigramme ","TRIGRAMME")
$password = [Microsoft.VisualBasic.Interaction]::InputBox("Enter User Password","Password") 
$Password2 =[Microsoft.VisualBasic.Interaction]::InputBox("Check User Password","Password") 
$securepassword = $Password2 | ConvertTo-SecureString -AsPlainText -Force

############
# GiveName #
############
$label1.Text = "First Name"
$label1.size = New-Object System.Drawing.Size(50,30)
$label1.Location = New-Object System.Drawing.Size(05,20)
$textbox1.Text = $Givename
$textbox1.Size = New-Object System.Drawing.Size(100,100)
$textbox1.Location = new-object System.Drawing.Size(70,20)
$textbox1.ReadOnly = $true

############
# SurName #
############
$label2.Text = "Last Name"
$label2.size = New-Object System.Drawing.Size(50,30)
$label2.Location = New-Object System.Drawing.Size(05,70)
$textbox2.Text = $Surname
$textbox2.Size = New-Object System.Drawing.Size(100,100)
$textbox2.Location = new-object System.Drawing.Size(70,70)
$textbox2.ReadOnly = $true

############
#   Name   #
############
$label3.Text = "Complet Name"
$label3.size = New-Object System.Drawing.Size(60,50)
$label3.Location = New-Object System.Drawing.Size(05,120)
$textbox3.Text = $Name
$textbox3.Size = New-Object System.Drawing.Size(100,100)
$textbox3.Location = new-object System.Drawing.Size(70,120)
$textbox3.ReadOnly = $true

########################
# UserPrincipaleName   #
########################
$label4.Text = "E-Mail Address"
$label4.size = New-Object System.Drawing.Size(60,50)
$label4.Location = New-Object System.Drawing.Size(05,170)
$textbox4.Text = $UserPrincipalName
$textbox4.Size = New-Object System.Drawing.Size(100,100)
$textbox4.Location = new-object System.Drawing.Size(70,170)
$textbox4.ReadOnly = $true

########################
#     SamAccountName   #
########################
$label5.Text = "Trigramme"
$label5.size = New-Object System.Drawing.Size(60,50)
$label5.Location = New-Object System.Drawing.Size(05,220)
$textbox5.Text = $SamAccountName
$textbox5.Size = New-Object System.Drawing.Size(100,100)
$textbox5.Location = new-object System.Drawing.Size(70,220)
$textbox5.ReadOnly = $true

########################
#     Pa$$w0Rd         #
########################
$label6.Text = "Password"
$label6.size = New-Object System.Drawing.Size(60,50)
$label6.Location = New-Object System.Drawing.Size(05,270)
$textbox6.Text = $Password2
$textbox6.Size = New-Object System.Drawing.Size(100,100)
$textbox6.Location = new-object System.Drawing.Size(70,270)
$textbox6.ReadOnly = $true

#################
# COpy USer     #
#################
$array = @('PRE-SALES','MANAGER','ACCOUNTING','COM','RECRUTING')

foreach($arr in $array)
{
    $combobox.Items.Add($arr)
}
$label7.Text = "Select a service"
$label7.size = New-Object System.Drawing.Size(150,11)
$label7.Location = New-Object System.Drawing.Point(250,20)
$combobox.Size = New-Object System.Drawing.Size(150,50)
$combobox.Location = New-Object System.Drawing.Size(250,40)

##################
# SetUP Variable #
##################
$ADVvalue = Get-ADUser -Identity "PRE-SALES" -Properties MemberOf | select -ExpandProperty MemberOf
$ManagerValue = Get-ADUser -Identity "MANAGER" -Properties MemberOf | select -ExpandProperty MemberOf
$ComptaValue = Get-ADUser -Identity "ACCOUNTING" -Properties MemberOf | select -ExcludeProperty Memberof
$ComValue = Get-ADUser -Identity "COM" -Properties MemberOf | select -ExcludeProperty MemberOf
$RecrutementValue = Get-ADUser -Identity "RECRUTING" -Properties MemberOf | select -ExcludeProperty MemberOf

###########################
# Validation of Copy User #
###########################
$button1.Text = "Validation"
$button1.size = New-Object System.Drawing.Size(62,20)
$button1.Location = New-Object System.Drawing.Size(410,40)
$button1.add_click({if ($combobox.Text -eq "PRE-SALES") {$textbox7.Text = $ADVvalue}
    if ($combobox.Text -eq "MANAGER") {$textbox7.Text = $ManagerValue} 
    if ($combobox.Text -eq "ACCOUNTING") {$textbox7.Text = $ComptaValue}
    if ($combobox.Text -eq "COM") {$textbox7.Text = $ComValue}
    if ($combobox.Text -eq "RECRUTING") {$textbox7.Text = $RecrutementValue}})

############################
#  User Domain Information #
############################
$label8.text = "User Domain Info"
$label8.Size = New-Object System.Drawing.Size(100,15)
$label8.Location = New-Object System.Drawing.Size(250,80)
$textbox7.Size = New-Object System.Drawing.Size(150,150)
$textbox7.Location = New-Object System.Drawing.Size(250,100)
$textbox7.ReadOnly = $true
$textbox7.ScrollBars = "Vertical"
$textbox7.Multiline = $true


#############
# OK BUTTON #
#############
$Okbutton.text = "OK"
$Okbutton.Size = New-Object System.Drawing.Size(30,30)
$Okbutton.Location = New-Object System.Drawing.Size(300,350)
$Okbutton.add_click({
    New-ADUser -GivenName $Givename -Surname $Surname -Name $Name -UserPrincipalName $UserPrincipalName -SamAccountName $SamAccountName -Path "CN=Users,DC=test,DC=com" -AccountPassword $securepassword
    Enable-ADAccount -Identity $SamAccountName
    if ($combobox.Text -eq "PRE-SALES"){$ADVvalue | Add-ADGroupMember -Members $SamAccountName}
    if ($combobox.Text -eq "MANAGER"){$ManagerValue | Add-ADGroupMember -Members $SamAccountName}
    if ($combobox.Text -eq "ACCOUNTING"){$ComptaValue | Add-ADGroupMember -Members $SamAccountName}
    if ($combobox.Text -eq "COM"){$ComValue | Add-ADGroupMember -Members $SamAccountName}
    if ($combobox.Text -eq "RECRUTING"){$RecrutementValue | Add-ADGroupMember -Members $SamAccountName}})


##############
# Show GUI   #
##############
$form.Controls.Add($button1)
$form.Controls.Add($Okbutton)
$form.Controls.Add($label1)
$form.Controls.Add($label2)
$form.Controls.Add($label3)
$form.Controls.Add($label4)
$form.Controls.Add($label5)
$form.Controls.Add($label6)
$form.Controls.add($label7)
$form.Controls.Add($label8)
$form.Controls.add($textbox1)
$form.Controls.Add($textbox2)
$form.Controls.Add($textbox3)
$form.Controls.Add($textbox4)
$form.Controls.Add($textbox5)
$form.Controls.Add($textbox6)
$form.controls.add($textbox7)
$form.Controls.Add($combobox)
$form.Add_Shown( { $form.Activate() } )
$form.ShowDialog()
