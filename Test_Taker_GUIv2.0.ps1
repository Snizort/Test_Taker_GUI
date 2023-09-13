$ErrorActionPreference = 'SilentlyContinue'

$sDir = "C:\Users\rbsmi\Documents\Practice_Test_Question"

Function Test-File{
    Param($Path)
        if(Test-Path -Path $Path){
            $true
        }
        else{
            New-Item -Path $Path -ItemType Directory
            $false
        }
}

Function TestGui{

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $Form = New-Object System.Windows.Forms.Form

    $Form.ClientSize = '1500, 1000'
    $Form.StartPosition = 'CenterScreen'
    $Form.BackColor = "#3760b5"
    $Form.FormBorderStyle = 'FixedDialog'
    $Form.MaximizeBox = 'False'
    $Form.Text = "Test GUI"

    LaunchMenu

    $Form.Add_FormClosing({
        $script:timer.Dispose()
    })

    $Form.Controls.Add($button1)

    $form.Add_FormClosing({
       $Script:timer.Dispose()
       ClearScreen -Form $Form  
    })

    [void]$Form.ShowDialog()
}

Function LaunchMenu{

    $dName = Get-ChildItem $sDir
    $lBox1Size = ($dname.Count +2)*20
    $Script:selectedItems = @()
    $TestBank = $null

    $Script:lBox1 = New-Object System.Windows.Forms.ListBox
    $Script:lBox1.Location = New-Object System.Drawing.Point(130,12)
    $Script:lBox1.Size = New-Object System.Drawing.Size(200,$lBox1Size)
    $Script:lBox1.Font = New-Object System.Drawing.Font("Lucida Console",18,[System.Drawing.FontStyle]::Regular)

    ForEach($item in $dname.name){$lBox1.Items.Add($Item)}

    $buttonYcord = $lbox1Size + 40 
    $button1 = New-Object System.Windows.Forms.Button
    $button1.Location = New-Object System.Drawing.Point(130,$buttonYcord)
    $button1.Size = New-Object System.Drawing.Size(90,40)
    $button1.Enabled = $true
    $button1.Text = "Select"
    $button1.BackColor = "#6D6D6D"
    $button1.TextAlign = "MiddleCenter"

    $button2 = New-Object System.Windows.Forms.Button
    $button2.Location = New-Object System.Drawing.Point(240 ,$buttonYcord)
    $button2.Size = New-Object System.Drawing.Size(90,40)
    $button2.Enabled = $true
    $button2.Text = "Back"
    $button2.BackColor = "#6D6D6D"
    $button2.TextAlign = "MiddleCenter"

    $Form.Controls.Add($Script:lBox1)
    $Form.Controls.Add($button1)
    $Form.Controls.Add($button2)
 
    $Button1.Add_Click({
        if($Script:lbox1.SelectedItems -like '*.*' -and $Script:lbox1.SelectedItems -notlike '*.csv' -or $Script:lbox1.SelectedIndex -eq -1){
            [System.Windows.MessageBox]::Show("Please select a File","Error")
        }
        elseif($Script:lbox1.SelectedItems -like '*.csv'){
            Write-Host "You found a CSV"
            $destName = $sDir + "\" + ($Script:selectedItems -join "\") + "\" + $Script:lbox1.SelectedItem
            $TestBank = Import-Csv -Path $destName
            ClearScreen -form $Form
            RunTest -Test $TestBank
        }
        else{
            $Script:selectedItems += $Script:lBox1.SelectedItems
            $destName = $sDir + "\" + ($Script:selectedItems -join "\")
            $Script:lBox1.Items.Clear()
            $dname = Get-ChildItem $destName
            for($i=0; $i-lt $dname.count;$i++){
                if($dname[$i].Name -Like '*.csv'){
                    $Script:lBox1.Items.Add($dname[$i].Name)
                }
                elseif($dname[$i].Name -notlike '*.*'){
                    $Script:lBox1.Items.Add($dname[$i].Name)
                }
            }
        }
    })
    $Button2.Add_Click({
        if($Script:selectedItems.Count -gt 1){
            $Script:selectedItems = $Script:selectedItems[0..($Script:selectedItems.Count -2)]
            $destName = $sDir + "\" + ($Script:selectedItems -join "\")
        }
        elseif($Script:selectedItems.Count -lt 1 -or $Script:selectedItems.Count -eq 1){
            $Script:selectedItems.Clear()
            $destName = $sDir
        }
        $Script:lBox1.Items.Clear()
        $dname = Get-ChildItem $destName
        for($i=0; $i-lt $dname.count;$i++){
            if($dname[$i].Name -Like '*.csv'){
                $Script:lBox1.Items.Add($dname[$i].Name)
            }
            elseif($dname[$i].Name -notlike '*.*'){
                $Script:lBox1.Items.Add($dname[$i].Name)
            }
        }
    })
}

Function RunTest{
    param($Test)
    $Script:questionIndex = 0
    $Script:Questions = $Test

    $Script:Button1 = New-Object System.Windows.Forms.Button
    $Button1.Location = New-Object System.Drawing.Point(600,400)
    $Button1.Size = New-Object System.Drawing.Size(300, 200)
    $Button1.Enable = $true
    $Button1.Text = "Start Test"
    $Button1.Font = New-Object System.Drawing.Font("Lucida Console",16,[System.Drawing.FontStyle]::Regular)
    $Button1.BackColor = "#FFFFFF"
    $Button1.TextAlign = "MiddleCenter"

    $Script:labelTimer = New-Object System.Windows.Forms.Label
    $Script:labelTimer.Location = New-Object System.Drawing.Point(600,25)
    $Script:labelTimer.Size = New-Object System.Drawing.Size(300,50)
    $Script:labelTimer.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
    $Script:labelTimer.TextAlign = "MiddleCenter"
    $Script:labelTimer.BackColor = '#FFFFFF'
    $Script:labelTimer.Name = 'Timer'
    $Script:labelTimer.Text = '02:30:00'

    $Script:Timer = New-Object System.Windows.Forms.Timer
    $Script:Timer.Interval = 1000
    $Script:remainingTime = [TimeSpan]::FromMinutes(150)
    $Script:Timer.Add_Tick({
        $Script:remainingTime = $Script:remainingTime.Subtract([TimeSpan]::FromSeconds(1))
        $Script:labelTimer.text = $Script:remainingTime.ToString("hh\:mm\:ss")

        if($Script:remainingTime.TotalSeconds -lt 0){
            $Script:Timer.Stop()

        }
    })

    $Form.Controls.Add($button1)
    $Form.Controls.Add($Script:labelTimer)
    $Form.Controls.Add($Script:Timer)

    $Button1.Add_Click({  
        $Script:questionIndex++
        if($Script:questionIndex -eq 1){
            $button1.Location = New-Object System.Drawing.Point(1300,925)
            $button1.Size = New-Object System.Drawing.Size(150,50)
            $button1.Font = New-Object System.Drawing.Font("Lucida Console",16,[System.Drawing.FontStyle]::Regular)
            $button1.Text = "Select"
            $button1.BackColor = "#FFFFFF"
            $button1.TextAlign = "MiddleCenter"
            $Script:AnswerChoice = @()
            $Script:Timer.Start()
            QuestionLoader -index $Script:questionIndex -questions $Script:Questions
        }
        elseif($Script:questionIndex -lt $Script:Questions.Count+1){
            QuestionCheck -form $Form -data $Script:Questions -index $Script:questionIndex
            QuestionLoader -index $Script:questionIndex -questions $Script:Questions
        }
        elseif($Script:questionIndex -eq $Script:Questions.Count+1){
            QuestionCheck -form $Form -data $Script:Questions -index $Script:questionIndex
            rmQControl -form $Form
            $Script:Timer.Stop()
            $button1.Text = "Main Menu"
            QuestionGrader            
        }
        else{
            Write-Host "To main menu"
            ClearScreen -form $Form
            LaunchMenu
        }
    })
}
Function QuestionLoader {
    param($index, $questions)
    rmQControl -form $Form
    if($index-lt $questions.Count -or $index -eq $questions.Count){
        $columns = @("A", "B", "C", "D", "E", "F", "G", "H")
        $emptyColumns = @()
        foreach($column in $columns){
            $value = $questions[$index-1].$column
            if($value -ne ""){
                $emptyColumns += $column
            }
        }
        controlLoader -data $emptyColumns -questions $questions -index $index
    }
    else{
        Write-Host "Fuck up"
    }
}

Function controlLoader{
    param($data, $questions, $index)
    
    $Script:label1 = New-Object System.Windows.Forms.Label
    $Script:label1.Location = New-Object System.Drawing.Point(250, 100)
    $Script:label1.Size = New-Object System.Drawing.Size(1000, 100)
    $Script:label1.BackColor = "#FFFFFF"
    $Script:label1.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
    $Script:label1.Text = $questions[$index-1].'Qnumb' + '.) ' + $questions[$index-1].'Question'
        
    $form.Controls.Add($script:label1)
    For($i=0; $i-lt $data.Length;$i++){
        $name = $data[$i]
        $distance = 250 + ($i * 75)

        $Script:cBox = New-Object System.Windows.Forms.CheckBox
        $Script:cBox.Location = New-Object System.Drawing.Point(250, $distance)
        $Script:cBox.Size = New-Object System.Drawing.Size(50,50)
        $Script:cBox.BackColor = '#FFFFFF'
        $Script:cBox.Text = $data[$i]

        $Script:label = New-Object System.Windows.Forms.Label
        $Script:label.Location = New-Object System.Drawing.Point(376, $distance)
        $Script:label.Size = New-Object System.Drawing.Size(835, 50)
        $Script:label.BackColor = "#FFFFFF"
        $Script:label.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
        $Script:label.Text = $questions[$index-1].$($data[$i])
        


        Set-Variable -Name "cBox$name" -Value $cBox
        Set-Variable -Name "label$name" -value $label
        $form.Controls.Add($Script:cBox)
        $form.Controls.Add($Script:label)
    }
}

Function QuestionCheck {
    param($form, $data, $index)

    $checkBoxNames = @()

    foreach ($control in $form.Controls) {
        if ($control -is [System.Windows.Forms.CheckBox] -and $control.Checked) {
            $checkBoxNames += $control.Text
            Write-Host "worked"
        }
    }

    $Script:AnswerChoice += 
    [PSCustomObject]@{
        CAnswer = $data[$index-2].'Answer'
        UAnswer = ($checkBoxNames -Join ",")
        Grade = ''
    }
}

Function QuestionGrader {
    $Grade = 0
    for($i=0;$i-ne$Script:AnswerChoice.count;$i++){
        if($Script:AnswerChoice[$i].CAnswer -eq $Script:AnswerChoice[$i].UAnswer){
            $Script:AnswerChoice[$i].Grade = 'C'
            $Grade++
        }
        else{
            $Script:AnswerChoice[$i].Grade = 'X'
        }
    }
    $Grade = ($Grade/$Script:AnswerChoice.Count)*100

    $label = New-Object System.Windows.Forms.label
    $label.location = New-Object System.Drawing.Point(600,400)
    $label.Size = New-Object System.Drawing.Size(300, 200)
    $label.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
    $label.BackColor = "#FFFFFF"
    $label.Text = $Grade

    $Form.Controls.Add($label)
}

#Function Clears Screen, any variable call after will re-initialize the variable
Function ClearScreen{
    param([System.Windows.Forms.Form]$form)
    $controlsToDispose = @()

    foreach($control in $Form.controls){
        if($control -is [System.Windows.Forms.Control]){
            $controlsToDispose += $control
        }
    }
    forEach($control in $controlsToDispose){
        $control.Dispose()
    }
}

Function rmQControl {
    param([System.Windows.Forms.Form]$form)
    $controls = @()
    foreach ($control in $form.Controls){
        if($control -is [System.Windows.Forms.CheckBox] -or $control -is [System.Windows.Forms.Label] -and $control.name -ne 'timer'){
            $controls += $Control
        }
    }
    forEach($control in $controls){
        $form.Controls.Remove($control)
        $control.Dispose()
    }
}