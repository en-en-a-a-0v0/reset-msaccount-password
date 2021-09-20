# Main Function to Update Password
Function Update_Password([string]$mgrusername,[string]$mgrpassword,[string]$username,[string]$password){
    
    
    if (-Not (Get-Module -ListAvailable -Name MSOnline)) {
        
        
        $LabelMsg.Text = "Installing module... If any notification comes out, please click ""Yes to All"""
        $NotificationForm.Hide()
        $NotificationForm.Show()
        install-module MSOnline
    }
    $LabelMsg.Text = "Initiallizing Module"
    $NotificationForm.Hide()
        $NotificationForm.Show()
    import-module MSOnline
    $PWord = ConvertTo-SecureString -String $mgrpassword -AsPlainText -Force
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $mgrusername, $PWord
    $LabelMsg.Text = "Getting Credential"
    $NotificationForm.Hide()
        $NotificationForm.Show()
    Connect-MsolService -Credential $Credential
    $LabelMsg.Text = "Changing Password"
    $NotificationForm.Hide()
        $NotificationForm.Show()
    Set-MsolUserPassword -UserPrincipalName $username -NewPassword $password -ForceChangePassword:$False
    $LabelMsg.Text = "Success"
    $NotificationForm.Hide()
        $NotificationForm.Show()
}

Add-Type -AssemblyName System.Windows.Forms
$PowerShellForms=New-Object System.Windows.Forms.Form
$PowerShellForms.Text = "Reset Password"
$PowerShellForms.Size = New-Object System.Drawing.Size(300,450) 
$PowerShellForms.StartPosition = "CenterScreen"
$PowerShellForms.SizeGripStyle = "Hide"

# Username Text
$LabelMgrUserName = New-Object System.Windows.Forms.Label
$LabelMgrUserName.Location = New-Object System.Drawing.Point(10,20) 
$LabelMgrUserName.Text = "MgrUserName:"
$LabelMgrUserName.AutoSize = $True
$PowerShellForms.Controls.Add($LabelMgrUserName)

# Username TextBox
$TextBoxMgrUserName = New-Object System.Windows.Forms.TextBox
$TextBoxMgrUserName.Text="itdeveloper@online.via.edu.au"
$TextBoxMgrUserName.Location = New-Object System.Drawing.Point(10,50)
$TextBoxMgrUserName.Size = New-Object System.Drawing.Size(260,20)
$PowerShellForms.Controls.Add($TextBoxMgrUserName)

# Password Text
$LabelMgrPassWord = New-Object System.Windows.Forms.Label
$LabelMgrPassWord.Location = New-Object System.Drawing.Point(10,90) 
$LabelMgrPassWord.Text = "MgrPassWord:"
$LabelMgrPassWord.AutoSize = $True
$PowerShellForms.Controls.Add($LabelMgrPassWord)

# Password TextBox
$TextBoxMgrPassWord = New-Object System.Windows.Forms.TextBox
$TextBoxMgrPassWord.Location = New-Object System.Drawing.Point(10,120)
$TextBoxMgrPassWord.Size = New-Object System.Drawing.Size(260,20)
$PowerShellForms.Controls.Add($TextBoxMgrPassWord)

# Username Text
$LabelUserName = New-Object System.Windows.Forms.Label
$LabelUserName.Location = New-Object System.Drawing.Point(10,160) 
$LabelUserName.Text = "UserName:"
$LabelUserName.AutoSize = $True
$PowerShellForms.Controls.Add($LabelUserName)

# Username TextBox
$TextBoxUserName = New-Object System.Windows.Forms.TextBox
$TextBoxUserName.Location = New-Object System.Drawing.Point(10,190)
$TextBoxUserName.Text="test_365login@online.via.edu.au"
$TextBoxUserName.Size = New-Object System.Drawing.Size(260,20)
$PowerShellForms.Controls.Add($TextBoxUserName)

# Password Text
$LabelPassWord = New-Object System.Windows.Forms.Label
$LabelPassWord.Location = New-Object System.Drawing.Point(10,230) 
$LabelPassWord.Text = "PassWord:"
$LabelPassWord.AutoSize = $True
$PowerShellForms.Controls.Add($LabelPassWord)

# Password TextBox
$TextBoxPassWord = New-Object System.Windows.Forms.TextBox
$TextBoxPassWord.Location = New-Object System.Drawing.Point(10,260)
$TextBoxPassWord.Size = New-Object System.Drawing.Size(260,20)
$PowerShellForms.Controls.Add($TextBoxPassWord)

# Nofitication Form
$NotificationForm=New-Object System.Windows.Forms.Form
$NotificationForm.Text = "Notification"
$NotificationForm.Size = New-Object System.Drawing.Size(400,100) 
$NotificationForm.StartPosition = "CenterScreen"
$LabelMsg = New-Object System.Windows.Forms.Label
$LabelMsg.Location = New-Object System.Drawing.Point(10,20)
$LabelMsg.AutoSize = $True
$NotificationForm.Controls.Add($LabelMsg)

# Submit Button
$SubmitButton = New-Object System.Windows.Forms.Button
$SubmitButton.Location = New-Object System.Drawing.Point(75,290)
$SubmitButton.Size = New-Object System.Drawing.Size(75,23)
$SubmitButton.Text="Submit"
$SubmitButton.Add_Click({
    $NotificationForm.Show()
    Update_Password $TextBoxMgrUserName.Text $TextBoxMgrPassWord.Text $TextBoxUserName.Text $TextBoxPassWord.Text
    
})
$PowerShellForms.Controls.Add($SubmitButton)

# Test Button
$TestForms=New-Object System.Windows.Forms.Form
$TestForms.Text = "Notification"
$TestForms.Size = New-Object System.Drawing.Size(400,400) 
$TestForms.StartPosition = "CenterScreen"
$LabelMsg = New-Object System.Windows.Forms.Label
$LabelMsg.Location = New-Object System.Drawing.Point(10,20) 
$LabelMsg.Text = "Installing module... If any notification comes out, please click ""Yes to All"""
$LabelMsg.AutoSize = $True
$TestForms.Controls.Add($LabelMsg)
$Test2 = New-Object System.Windows.Forms.Button
$Test2.Location = New-Object System.Drawing.Point(150,290)
$Test2.Size = New-Object System.Drawing.Size(100,23)
$Test2.Text="Test"
$Test2.Add_Click({

    $TestForms.Hide()
    $LabelMsg.Text = "Changed Text"
})
$TestForms.Controls.Add($Test2)


$Test = New-Object System.Windows.Forms.Button
$Test.Location = New-Object System.Drawing.Point(150,290)
$Test.Size = New-Object System.Drawing.Size(100,23)
$Test.Text="Test"

$Test.Add_Click({

    $TestForms.Show()
    
})
$PowerShellForms.Controls.Add($Test)
# Show
$PowerShellForms.ShowDialog()

# 需要完善的逻辑
# 后缀选择
# 账号密码非空
# 返回信息展示
# 布局合理化