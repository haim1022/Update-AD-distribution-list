#Add necessary libraries
Add-Type -AssemblyName System.Windows.Forms
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
#Hide the console
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

#Set up functions
function update-list {
	$errorF = $false #Set to true if there has been an error trying to change the AD group members
    if($Distribution_ComboBox.SelectedIndex -ne 0) {
        $changedF = $false #True if found changes in the group members
        $global:DistributionListUsers = Get-ADGroupMember -Identity $distributionLists[$Distribution_ComboBox.SelectedIndex] | Get-ADUser -Properties SamAccountName, Name, EmployeeID, Office | Sort-Object -Property Office
        foreach($user in $global:tempDistributionListUsers) { #Look for non existing users in the AD group and add them
            $existingUser = $global:DistributionListUsers | Where-Object {$_.Office -like ($user.Office)}
            if($null -eq $existingUser) { #Found user which is not in the AD group - add it
                Add-ADGroupMember -Identity $distributionLists[$Distribution_ComboBox.SelectedIndex] -Members $user.SamAccountName
				if($Error[0] | Select-String "Insufficient access rights") {
					$errorF = $true
				}
                $changedF = $true
            }
        }
        foreach($user in $global:DistributionListUsers) { #Look for existing users in the AD group and remove them
            $existingUser = $global:tempDistributionListUsers | Where-Object {$_.Office -like ($user.Office)}
            if($null -eq $existingUser) {
                Remove-ADGroupMember -Identity $distributionLists[$Distribution_ComboBox.SelectedIndex] -Members $user.SamAccountName  -Confirm:$false
				if($Error[0] | Select-String "Insufficient access rights") {
					$errorF = $true
				}
                $changedF = $true
            }
        }
		if($errorF -eq $false)
		{
			if($changedF -eq $true) {
				[System.Windows.Forms.MessageBox]::Show("בוצע בהצלחה", $main_Form.Text, 0, [System.Windows.Forms.MessageBoxIcon]::None, [System.Windows.Forms.MessageBoxDefaultButton]::Button1, [System.Windows.Forms.MessageBoxOptions]::RightAlign)
			}
			else {
				[System.Windows.Forms.MessageBox]::Show("לא נמצאו שינויים", $main_Form.Text, 0, [System.Windows.Forms.MessageBoxIcon]::None, [System.Windows.Forms.MessageBoxDefaultButton]::Button1, [System.Windows.Forms.MessageBoxOptions]::RightAlign)
			}
		}
		else {
			[System.Windows.Forms.MessageBox]::Show("אין לך הרשאות לבצע שינויים ברשימות תפוצה", $main_Form.Text, 0, [System.Windows.Forms.MessageBoxIcon]::None, [System.Windows.Forms.MessageBoxDefaultButton]::Button1, [System.Windows.Forms.MessageBoxOptions]::RightAlign)
		}
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("יש לבחור ברשימת תפוצה", $main_Form.Text, 0, [System.Windows.Forms.MessageBoxIcon]::None, [System.Windows.Forms.MessageBoxDefaultButton]::Button1, [System.Windows.Forms.MessageBoxOptions]::RtlReading)
    }
}
function search-users {
    $flush = $Entire_User_List_ListBox.Items.Clear()
    $global:currentUsersList = $global:entireUsersList | Where-Object {$_.Office -like ($search_Textbox.Text + "*")}
    foreach($user in $global:currentUsersList) {
        if(($null -ne $user.EmployeeID) -and ($null -ne $user.Office)) {
            $global:existingUserFlag = $false
            foreach($distributionUser in $global:tempDistributionListUsers) {
                if($user.EmployeeID -eq $distributionUser.EmployeeID) {
                    $global:existingUserFlag = $true
                }
            }
            if($global:existingUserFlag -eq $false) {
                $flush = $Entire_User_List_ListBox.Items.Add($user.Office)
            }
        }
    }
}
function refresh-distributionUsersList {
	$global:currentSelectedIndexComboBox = $Distribution_ComboBox.SelectedIndex
	$flush = $Distribution_Users_List_ListBox.Items.Clear()
	if($distributionLists[$Distribution_ComboBox.SelectedIndex] -ne "") {
		$global:tempDistributionListUsers = Get-ADGroupMember -Identity $distributionLists[$Distribution_ComboBox.SelectedIndex] | Get-ADUser -Properties SamAccountName, Name, EmployeeID, Office | Sort-Object -Property Office
		foreach($user in $global:tempDistributionListUsers) {
			if(($null -ne $user.EmployeeID) -and ($null -ne $user.Office)) {
				$flush = $Distribution_Users_List_ListBox.Items.Add($user.Office)
			}
		}
	}
	else {
		$global:tempDistributionListUsers = ""
	}
	search-users #Update users list
}

#Declare global variables
#Changes required here
$availableDivisions = "בחר...", "שיווק", "השוואת מחירים", "טכנולוגיות", "חווית לקוח", "מכירות", "SEO אתרי הקבוצה", "כספים", "לוגיסטיקה"
$distributionLists = "", "", "", "", "", "", "", "", "" #Distribution lists corresponding to the listed divisions
$global:currentSelectedIndexComboBox = 0
$global:existingUserFlag = $false #Used to determine if a user exists in the both lists

#Get entire users list and save it
$global:entireUsersList = Get-ADUser -Filter * -SearchBase "OU=RegularUsers,OU=Users,OU=GoldenPages,DC=Golden,DC=Pages" -Properties SamAccountName, Name, EmployeeID, Office | Sort-Object -Property Office
$global:currentUsersList = $global:entireUsersList #User list after filtering the search
$global:DistributionListUsers = 0 #Real AD group users list
$global:tempDistributionListUsers = 0 #Change the temporary list to prevent updating the real AD group before saving

#Define the form's spacing
$formStandardSpace = 20
$formStartXPos = 10
$formStartYPos = 10
$comboBoxWidth = 150
$comboBoxHeight = 30
$textboxHeight = 20
$textboxWidth = 150
$buttonHeight = 22
$buttonWidth = 75

#Main form objects
    #Form properties
        #main_Form
$main_Form = New-Object System.Windows.Forms.Form
$main_Form.Height = 525
$main_Form.Width = 500
$main_Form.Text = "עדכון רשימות תפוצה"
$main_Form.FormBorderStyle = "FixedDialog"
$main_Form.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
$main_Form.RightToLeftLayout = [System.Windows.Forms.RightToLeft]::Yes

    #Labels
        #Title_Label
$Title_Label = New-Object System.Windows.Forms.Label
$Title_Label.Location = New-Object System.Drawing.Point($formStartXPos, $formStartYPos)
$Title_Label.AutoSize = $true
$Title_Label.Text = "עדכון רשימות תפוצה"
$Title_Label.Font = [System.Drawing.Font]::new("Calibri", 16, [System.Drawing.FontStyle]::Bold)
$main_Form.Controls.Add($Title_Label)
        #Distribution_Combobox_Label
$Distribution_Combobox_Label = New-Object System.Windows.Forms.Label
$Distribution_Combobox_Label.Location = New-Object System.Drawing.Point($formStartXPos, ($formStartYPos + $textboxHeight*2))
$Distribution_Combobox_Label.AutoSize = $true
$Distribution_Combobox_Label.Text = "בחר בתפוצה הרלוונטית:"
$Distribution_Combobox_Label.Font = [System.Drawing.Font]::new("Calibri", 12)
$main_Form.Controls.Add($Distribution_Combobox_Label)
        #Entire_User_List_Label
$Entire_User_List_Label = New-Object System.Windows.Forms.Label
$Entire_User_List_Label.Location = New-Object System.Drawing.Point($formStartXPos, ($formStartYPos + $textboxHeight*4))
$Entire_User_List_Label.AutoSize = $true
$Entire_User_List_Label.Text = "רשימת המשתמשים הכללית:"
$Entire_User_List_Label.Font = [System.Drawing.Font]::new("Calibri", 12)
$main_Form.Controls.Add($Entire_User_List_Label)
        #Distribution_List_Label 
$Distribution_List_Label = New-Object System.Windows.Forms.Label
$Distribution_List_Label.Location = New-Object System.Drawing.Point(($formStartXPos + $textboxWidth*2), ($formStartYPos + $textboxHeight*4))
$Distribution_List_Label.AutoSize = $true
$Distribution_List_Label.Text = "רשימת התפוצה שנבחרה:"
$Distribution_List_Label.Font = [System.Drawing.Font]::new("Calibri", 12)
$main_Form.Controls.Add($Distribution_List_Label )

    #ComboBoxes
        #Distribution_ComboBox
$Distribution_ComboBox = New-Object System.Windows.Forms.ComboBox
$Distribution_ComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$Distribution_ComboBox.Width = $comboBoxWidth
$Distribution_ComboBox.Height = $comboBoxHeight
$Distribution_ComboBox.Location = New-Object System.Drawing.Point(($main_Form.Width - $comboBoxWidth - $formStandardSpace), ($formStartYPos + $textboxHeight*2))
$Distribution_ComboBox.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
foreach($list in $availableDivisions) {
    $flush = $Distribution_ComboBox.Items.Add($list)
}
$Distribution_ComboBox.SelectedIndex = $global:currentSelectedIndexComboBox
$main_Form.Controls.Add($Distribution_ComboBox)

    #ListBoxes
        #Entire_User_List_ListBox
$Entire_User_List_ListBox = New-Object System.Windows.Forms.ListBox
$Entire_User_List_ListBox.Location = New-Object System.Drawing.Point(($Entire_User_List_Label.Location.X + $formStandardSpace/2), ($Entire_User_List_Label.Location.Y + $formStandardSpace/2 + $textboxHeight))
$Entire_User_List_ListBox.Width = $Entire_User_List_Label.Size.Width - $formStandardSpace
$Entire_User_List_ListBox.Height = 300
$test = search-users #Fill the user list
$main_Form.Controls.Add($Entire_User_List_ListBox)
        #Distribution_Users_List_ListBox
$Distribution_Users_List_ListBox = New-Object System.Windows.Forms.ListBox
$Distribution_Users_List_ListBox.Location = New-Object System.Drawing.Point(($Distribution_List_Label.Location.X + $formStandardSpace/2), ($Distribution_List_Label.Location.Y + $formStandardSpace/2 + $textboxHeight))
$Distribution_Users_List_ListBox.Width = $Entire_User_List_Label.Size.Width - $formStandardSpace
$Distribution_Users_List_ListBox.Height = 300
$main_Form.Controls.Add($Distribution_Users_List_ListBox)

    #Textboxes
        #search_Textbox
$search_Textbox = New-Object System.Windows.Forms.Textbox
$search_Textbox.Width = $Entire_User_List_ListBox.Width
$search_Textbox.Height = $textboxHeight
$search_Textbox.Location = New-Object System.Drawing.Point(($Entire_User_List_ListBox.Location.X), ($Entire_User_List_ListBox.Location.Y + $Entire_User_List_ListBox.Height))
$main_Form.Controls.Add($search_Textbox)

    #Buttons
        #add_To_List_Button
$add_To_List_Button = New-Object System.Windows.Forms.Button
$add_To_List_Button.Width = $buttonWidth
$add_To_List_Button.Height = $buttonHeight
$add_To_List_Button.Location = New-Object System.Drawing.Point(($main_Form.Width/2 - $buttonWidth/2), ($Entire_User_List_ListBox.Location.Y + $Entire_User_List_ListBox.Height/2 - $buttonHeight*2))
$add_To_List_Button.Text = ">`nהוסף לרשימה"
$add_To_List_Button.AutoSize = $true
$main_Form.Controls.Add($add_To_List_Button)
        #remove_From_List_Button
$remove_From_List_Button = New-Object System.Windows.Forms.Button
$remove_From_List_Button.Width = $buttonWidth
$remove_From_List_Button.Height = $buttonHeight
$remove_From_List_Button.Location = New-Object System.Drawing.Point(($main_Form.Width/2 - $buttonWidth/2), ($Entire_User_List_ListBox.Location.Y + $Entire_User_List_ListBox.Height/2))
$remove_From_List_Button.Text = "<`nהסר מהרשימה"
$remove_From_List_Button.AutoSize = $true
$main_Form.Controls.Add($remove_From_List_Button)
        #search_Button
$search_Button = New-Object System.Windows.Forms.Button
$search_Button.Width = $buttonWidth
$search_Button.Height = $buttonHeight
$search_Button.Location = New-Object System.Drawing.Point(($search_Textbox.Location.X + $textboxWidth + $formStandardSpace),$search_Textbox.Location.Y)
$search_Button.Text = "חפש"
$main_Form.Controls.Add($search_Button)
        #run_Button
$run_Button = New-Object System.Windows.Forms.Button
$run_Button.Width = $buttonWidth
$run_Button.Height = $buttonHeight
$run_Button.Location = New-Object System.Drawing.Point(($main_Form.Width - $formStandardSpace - $buttonWidth*2),($main_Form.Height - $formStandardSpace*2 - $buttonHeight))
$run_Button.Text = "עדכן רשימה"
$main_Form.Controls.Add($run_Button)
        #cancel_Button
$cancel_Button = New-Object System.Windows.Forms.Button
$cancel_Button.Width = $buttonWidth
$cancel_Button.Height = $buttonHeight
$cancel_Button.Location = New-Object System.Drawing.Point(($main_Form.Width - $formStandardSpace - $buttonWidth),($main_Form.Height - $formStandardSpace*2 - $buttonHeight))
$cancel_Button.Text = "יציאה"
$main_Form.Controls.Add($cancel_Button)

    #Custom events
        #Distribution_ComboBox
$Distribution_ComboBox.Add_SelectedIndexChanged(
        {
            refresh-distributionUsersList
        }
    )
        #search_Textbox
$search_Textbox.Add_KeyDown(
        {
            if(($null -ne $_.KeyCode) -and ($_.KeyCode -ne "ControlKey")) { #Form GUI updates only after second key press because of how powershell handles GUI
                [System.Windows.Forms.SendKeys]::SendWait("^")
            }
            search-users #Update users list
        }
    )
        #add_To_List_Button
$add_To_List_Button.Add_Click(
        {
            if(($Distribution_ComboBox.SelectedIndex -ne 0) -and ($Entire_User_List_ListBox.SelectedIndex -ne -1)) {
                $global:tempDistributionListUsers += ($global:entireUsersList | Where-Object {$_.Office -like ($Entire_User_List_ListBox.Text + "*")})
                $flush = $Distribution_Users_List_ListBox.Items.Add($global:tempDistributionListUsers[$global:tempDistributionListUsers.Count – 1].Office)
                search-users #Update users list
            }
        }
    )
        #remove_From_List_Button
$remove_From_List_Button.Add_Click(
        {
            if(($Distribution_ComboBox.SelectedIndex -ne 0) -and ($Distribution_Users_List_ListBox.SelectedIndex -ne -1)) {
                $temptempDistributionListUsers = @() #Create a temporary array to remove the user from
                foreach($object in $tempDistributionListUsers) {
                    if($object.Office -like ($Distribution_Users_List_ListBox.Text + "*")) {
                        #Don't add the selected user to the new array
                    }
                    else {
                        $temptempDistributionListUsers += $object
                    }
                }
                $flush = $Distribution_Users_List_ListBox.Items.Remove($Distribution_Users_List_ListBox.Text)
                $global:tempDistributionListUsers = $temptempDistributionListUsers
                search-users #Update users list
            }
        }
    )
        #run_button
$run_button.Add_Click(
        {
		    update-list
			refresh-distributionUsersList
        }
    )
        #cancel_button
$cancel_button.Add_Click(
        {
            $main_Form.close()
        }
    )
        #main_Form
$main_Form.Add_KeyDown(
        {
            if($_.KeyCode -eq "Escape") {
                $main_Form.Close()
            }
        }
    )

$main_Form.ShowDialog()

taskkill /im "Run.exe"