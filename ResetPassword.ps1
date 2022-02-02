function Show_Notification([string]$s){
    $r=[System.Windows.Forms.MessageBox]::Show($s, "Notification")
}

Function Show_Error([string]$e){
    $r=[System.Windows.Forms.MessageBox]::Show($e, "Error")
}

function Check_Plugin($m){
    if (-Not (Get-Module -ListAvailable -Name $m)) {
        Show_Notification "Attention!!!!`nT$m will be installed.`nClick OK to start install.`nPlease wait for a while`nIf any notification comes out, please click ""Yes to All"""       
        $LabelMsg.Text = "Installing $m ..."
        try{
            install-module $m
        }
        catch{
             return $_.Exception.Message
        }
    }
    return $true
}

# This is the function for login only.
function Login_Core([string]$mngrusername,[SecureString]$mngrpassword){
    try{
        Connect-MsolService -Credential (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $mngrusername, $mngrpassword) -ErrorAction Stop
        return $true
    }
    catch{
        return $_.Exception.Message
    }
}

# This is the function changing user password only.
# Login_Core should be called to ensure login successfully before call this function
# Check Plugin is suggested Before the first time call this function
function ChangePW_Core([string]$username,[string]$password){
    # Change Password
    import-module MSOnline  
    $LabelMsg.Text = "Changing Password"
    try{
        Set-MsolUserPassword -UserPrincipalName $username -NewPassword $password -ForceChangePassword:$False -ErrorAction Stop
        return "is the new password. Update Success"
        # return $true
    }
    catch{
        return $_.Exception.Message
    }

}

function Update_Password([string]$mngrusername,[SecureString]$mngrpassword,[string]$username,[string]$password){
    
    $NotificationForm.Show() 

    # Check Plugin
    $Plugin= Check_Plugin "MSOnline"
    if(-Not ($Plugin -eq $true)){
        Show_Error $Plugin
        $NotificationForm.Hide()
        return
    }
    
    # Login
    $LabelMsg.Text = "Getting Credential"
    $Login=Login_Core $mngrusername $mngrpassword
    if(-Not ($Login -eq $true)){
        Show_Error $Login
        $NotificationForm.Hide()
        return
    }

    # Call Password Changing Method
    $Result=ChangePW_Core $username $password
    if(-Not ($Result -eq $true)){
        Show_Error $Result
        $NotificationForm.Hide()
        return
    }
    $NotificationForm.Hide()
}

# Return a label object
function Generate_Label ([int]$locx,[int]$locy,[string]$text){
    $Label = New-Object System.Windows.Forms.Label
    $Label.AutoSize = $True
    $Label.Location = New-Object System.Drawing.Point($locx,$locy) 
    $Label.Text = $text    
    return $Label
}

# Return a textbox object
function Generate_TextBox ([int]$locx,[int]$locy,[int]$sizew,[int]$sizeh,[string]$text){
    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Location = New-Object System.Drawing.Point($locx,$locy)
    $TextBox.Size = New-Object System.Drawing.Size($sizew,$sizeh)
    $TextBox.Text = $text
    return $TextBox
}

Add-Type -AssemblyName System.Windows.Forms

# Main Window
$PowerShellForms=New-Object System.Windows.Forms.Form
$PowerShellForms.Text = "VIA Microsoft Account Management Tool"
$PowerShellForms.AutoSize = $True
$PowerShellForms.StartPosition = "CenterScreen"
$PowerShellForms.SizeGripStyle = "Hide"

#Tab Controller
$FormTabControl = New-object System.Windows.Forms.TabControl 
$FormTabControl.Size = "755,475" 
$PowerShellForms.Controls.Add($FormTabControl)

#Tab1 - Change Password
$Tab1 = New-object System.Windows.Forms.Tabpage
$Tab1.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab1.Name = "Tab1" 
$Tab1.Text = "Reset Passwords” 
$Tab1.AutoSize = $True
$FormTabControl.Controls.Add($Tab1)

$startx=10
$textboxw=260
$textboxh=20

# Site Admin Username Text
$Tab1.Controls.Add((Generate_Label $startx 20 "MngrUserName:"))

# Site Admin Username TextBox
$TextBoxMngrUserName=Generate_TextBox $startx 50 $textboxw $textboxh ""
$Tab1.Controls.Add($TextBoxMngrUserName)

# Site Admin Password Text
$Tab1.Controls.Add((Generate_Label $startx 90 "MngrPassword:"))

# Site Admin Password TextBox
$TextBoxMngrPassWord=Generate_TextBox $startx 120 $textboxw $textboxh ""
$Tab1.Controls.Add($TextBoxMngrPassWord)

# Username Text
$Tab1.Controls.Add((Generate_Label $startx 160 "UserName:"))

# Username TextBox
$TextBoxUserName=Generate_TextBox $startx 190 $textboxw $textboxh ""
$Tab1.Controls.Add($TextBoxUserName)

# Password Text
$Tab1.Controls.Add((Generate_Label $startx 230 "Password:"))

# Password TextBox
$TextBoxPassWord=Generate_TextBox $startx 260 $textboxw $textboxh ""
$Tab1.Controls.Add($TextBoxPassWord)

# Nofitication Form
$NotificationForm=New-Object System.Windows.Forms.Form
$NotificationForm.Text = "Notification"
$NotificationForm.Size = New-Object System.Drawing.Size(400,100) 
$NotificationForm.StartPosition = "CenterScreen"
$LabelMsg = Generate_Label 10 20 ""
$NotificationForm.Controls.Add($LabelMsg)

# Submit Button
$SubmitButton = New-Object System.Windows.Forms.Button
$SubmitButton.Location = New-Object System.Drawing.Point(75,290)
$SubmitButton.Size = New-Object System.Drawing.Size(75,23)
$SubmitButton.Text="Submit"
$SubmitButton.Add_Click({
    # Check Blank
    if ($TextBoxMngrUserName.Text.Length -lt 1) {
        Show_Notification "Please enter Admin email"
        return
    }

    if ($TextBoxMngrPassWord.Text.Length -lt 1) {
        Show_Notification "Please enter Admin Password"
        return
    }

    if ($TextBoxUserName.Text.Length -lt 1) {
        Show_Notification "Please enter User Email"
        return
    }

    if ($TextBoxPassWord.Text.Length -lt 1) {
        Show_Notification "Please enter User Password"
        return
    }

    #To Secure String
    $AdminWP = ConvertTo-SecureString -String $TextBoxMngrPassWord.Text -AsPlainText -Force
    # $UserWP = ConvertTo-SecureString -String $TextBoxPassWord.Text -AsPlainText -Force

    # Run update
    Update_Password $TextBoxMngrUserName.Text $AdminWP $TextBoxUserName.Text $TextBoxPassWord.Text
})
$Tab1.Controls.Add($SubmitButton)

#Tab2 - Create Accounts
$Tab2 = New-object System.Windows.Forms.Tabpage
$Tab2.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab2.UseVisualStyleBackColor = $True 
$Tab2.Name = "Tab1" 
$Tab2.Text = "Create Accounts” 
$FormTabControl.Controls.Add($Tab2)

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'Documents (*.docx)|*.docx|SpreadSheet (*.xlsx)|*.xlsx'
}

$SubmitButton1 = New-Object System.Windows.Forms.Button
$SubmitButton1.Location = New-Object System.Drawing.Point(75,290)
$SubmitButton1.Size = New-Object System.Drawing.Size(75,23)
$SubmitButton1.Text="Submit"
$SubmitButton1.Add_Click({
    $FileBrowser.ShowDialog()
})
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'SpreadSheet (*.csv)|*.csv'
}
# $null = $FileBrowser.ShowDialog()

$Tab2.Controls.Add($SubmitButton1)









# Show

$PowerShellForms.ShowDialog() | Out-Null

# 需要完善的逻辑
# 后缀选择
# 账号密码非空
# 返回信息展示
# 布局合理化
# Use: "ps2exe .\ResetPassword.ps1 .\target.exe -noConsole -credentialGUI" to compile
