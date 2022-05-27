
Echo "Loading Dictionary please Wait..."


    $fromDir = $PSScriptRoot
    Set-Location $fromDir

# Loading Windows Form
#-----------------------------------------------

#region
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
#endregion            


# Windows FORM configuration
#-----------------------------------------------

# Creating Form
$form = New-Object Windows.Forms.Form

# Windows Settings
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $False
$form.MinimizeBox = $True

# title
$form.Text = "Erwan's SUTOM Finder"

# Size 
$form.Size = New-Object System.Drawing.Size(850,460)

# Loading Jira base
#-----------------------------------------------

$fromDir = Get-Location


$global:BaseMot =  Get-Content $fromDir\listeMots.csv | ConvertFrom-Csv -Delimiter ';'
$global:Selected = $global:BaseMot
$Pattern = ""

# Creating components
#-----------------------------------------------

# Title
$label_Title = New-Object System.Windows.Forms.Label
$label_Title.AutoSize = $true
$label_Title.Location = New-Object System.Drawing.Point(20,10)
$label_Title.Size = New-Object System.Drawing.Size(150,10)
$label_Title.Text = "Erwan's SUTOM Finder"
$label_Title.Font = New-Object System.Drawing.Font("Segoe UI",15,[System.Drawing.FontStyle]::Regular)

# Label_Explain
$label_Explain = New-Object System.Windows.Forms.Label
$label_Explain.AutoSize = $true
$label_Explain.Location = New-Object System.Drawing.Point(260,22)
$label_Explain.Size = New-Object System.Drawing.Size(150,20)
$label_Explain.Text = "Most Frequent letters: e,a,s,i,n,t,r,u,o,d"

# Output Selected Words
$textBox_List = New-Object System.Windows.Forms.TextBox
$textBox_List.Font = New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Regular)
$textBox_List.Location = New-Object System.Drawing.Point(20,50) ### Location of the text box
$textBox_List.Size = New-Object System.Drawing.Size(800,200) ### Size of the text box
$textBox_List.Multiline = $true ### Allows multiple lines of data
$textBox_List.AcceptsReturn = $true ### By hitting enter it creates a new line
$textBox_List.ScrollBars = "Vertical" ### Allows for a vertical scroll bar if the list of text is too big for the window
$textBox_List.text = $Selected | Format-Table -AutoSize | Out-String

# -------

# Label_Begin
$label_Begin = New-Object System.Windows.Forms.Label
$label_Begin.AutoSize = $true
$label_Begin.Location = New-Object System.Drawing.Point(30,272)
$label_Begin.Size = New-Object System.Drawing.Size(150,20)
$label_Begin.Text = "Word begin with (ab,c..)"

# TextBox Begin
$textbox_Begin = New-Object System.Windows.Forms.TextBox
$textbox_Begin.AutoSize = $true
$textbox_Begin.Location = New-Object System.Drawing.Point(200,270)
$textbox_Begin.Name = 'textbox_sw'
$textbox_Begin.Size = New-Object System.Drawing.Size(250,20)
$textbox_Begin.Text = ""

# Bouton Begin
$button_Begin = New-Object System.Windows.Forms.Button
$button_Begin.Text = "Begins"
$button_Begin.Size = New-Object System.Drawing.Size(65,20)
$button_Begin.Location = New-Object System.Drawing.Size(460,270)

# -------

# Label_Select
$label_Select = New-Object System.Windows.Forms.Label
$label_Select.AutoSize = $true
$label_Select.Location = New-Object System.Drawing.Point(30,302)
$label_Select.Size = New-Object System.Drawing.Size(150,20)
$label_Select.Text = "Letters included (t,o,er...)"

# TextBox Select
$textbox_Select = New-Object System.Windows.Forms.TextBox
$textbox_Select.AutoSize = $true
$textbox_Select.Location = New-Object System.Drawing.Point(200,300)
$textbox_Select.Name = 'textbox_sw'
$textbox_Select.Size = New-Object System.Drawing.Size(250,20)
$textbox_Select.Text = ""

# Bouton Select
$button_Select = New-Object System.Windows.Forms.Button
$button_Select.Text = "Include"
$button_Select.Size = New-Object System.Drawing.Size(65,20)
$button_Select.Location = New-Object System.Drawing.Size(460,300)

# -------

# Label_UnSelect
$label_UnSelect = New-Object System.Windows.Forms.Label
$label_UnSelect.AutoSize = $true
$label_UnSelect.Location = New-Object System.Drawing.Point(30,332)
$label_UnSelect.Size = New-Object System.Drawing.Size(150,20)
$label_UnSelect.Text = "Letters excluded (r,s,...)"

# TextBox UnSelect
$textbox_UnSelect = New-Object System.Windows.Forms.TextBox
$textbox_UnSelect.AutoSize = $true
$textbox_UnSelect.Location = New-Object System.Drawing.Point(200,330)
$textbox_UnSelect.Name = 'textbox_sw'
$textbox_UnSelect.Size = New-Object System.Drawing.Size(250,20)
$textbox_UnSelect.Text = ""

# Bouton UnSelect
$button_UnSelect = New-Object System.Windows.Forms.Button
$button_UnSelect.Text = "Exclude"
$button_UnSelect.Size = New-Object System.Drawing.Size(65,20)
$button_UnSelect.Location = New-Object System.Drawing.Size(460,330)

# -------

# Label_Length
$label_Length = New-Object System.Windows.Forms.Label
$label_Length.AutoSize = $true
$label_Length.Location = New-Object System.Drawing.Point(30,362)
$label_Length.Size = New-Object System.Drawing.Size(150,20)
$label_Length.Text = "Select length of word"

# TextBox UnSelect
$textbox_Length = New-Object System.Windows.Forms.TextBox
$textbox_Length.AutoSize = $true
$textbox_Length.Location = New-Object System.Drawing.Point(200,360)
$textbox_Length.Name = 'textbox_sw'
$textbox_Length.Size = New-Object System.Drawing.Size(250,20)
$textbox_Length.Text = ""

# Bouton UnSelect
$button_Length = New-Object System.Windows.Forms.Button
$button_Length.Text = "Length"
$button_Length.Size = New-Object System.Drawing.Size(65,20)
$button_Length.Location = New-Object System.Drawing.Size(460,360)

# -------

# Bouton Cancel
$button_Cancel = New-Object System.Windows.Forms.Button
$button_Cancel.Text = "Restart"
$button_Cancel.Size = New-Object System.Drawing.Size(65,30)
$button_Cancel.Location = New-Object System.Drawing.Size(750,300)
$button_Cancel.Font = New-Object System.Drawing.Font("Lucida Console",9,[System.Drawing.FontStyle]::Regular)

# Bouton Quitter
$button_quit = New-Object System.Windows.Forms.Button
$button_quit.Text = "Quit"
$button_quit.Size = New-Object System.Drawing.Size(65,45)
$button_quit.Location = New-Object System.Drawing.Point(750,340)
$button_quit.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)


# event when clicking Close button
#-----------------------------------------------
$button_quit.Add_Click(
{
    $form.Close();
})

 $button_Begin.Add_Click(
{
    

    [array]$list = @()
    [array]$Selected2  = @()
    $Char = $textbox_Begin.Text   
    $Pattern = '^' + [regex]::escape($textbox_Begin.Text) 
  
     foreach ($line in $Selected)
    {  
    $Check = $line.Mots
  
        if ($Check -match $Pattern) {   

            $Selected2 += [pscustomobject]@{Mots=$Check}   
         }
    }
  
    $global:Selected = $Selected2
    $textBox_List.text = $Selected | Format-Table -AutoSize | Out-String
    $form.refresh

 })



 $button_Select.Add_Click(
{

 

    [array]$Selected2  = @()
    [array]$list = @()

    $InputString = $textbox_Select.Text
    $list =$InputString.Split(",")

    foreach ($line in $list)
    {
    [array]$Selected2  = @()
    $Pattern = [regex]::escape($line)
         
         foreach ($line in $Selected)
        {  
        $Check = $line.Mots
  
             if ($Check -match $Pattern) {   

                 $Selected2 += [pscustomobject]@{Mots=$Check}   
             }
         }

         $Selected = $Selected2
         
    }

    $global:Selected = $Selected2
    $textBox_List.text = $Selected | Format-Table -AutoSize | Out-String
    $form.refresh

 })

  $button_UnSelect.Add_Click(
{

    [array]$Selected2  = @()
    [array]$list = @()

    $InputString = $textbox_UnSelect.Text
    $list =$InputString.Split(",")

    foreach ($line in $list)
    {
    [array]$Selected2  = @()
    $Pattern = [regex]::escape($line)
         
         foreach ($line in $Selected)
        {  
        $Check = $line.Mots
  
             if ($Check -match $Pattern) {   
             #$Selected2 += [pscustomobject]@{Mots=" "}
             }
             else{
                $Selected2 += [pscustomobject]@{Mots=$Check}   
             }

         }

         $Selected = $Selected2
         
    }

    $global:Selected = $Selected2
    $textBox_List.text = $Selected | Format-Table -AutoSize | Out-String
    $form.refresh


 })

   $button_Length.Add_Click(
{


    [array]$Selected2  = @()
  
     foreach ($line in $Selected)
    {  
    $Check = $line.Mots
          
        if ($Check.length -eq [int]$textbox_Length.Text) {   
               
          $Selected2 += [pscustomobject]@{Mots=$Check}
        }  
                      
    }
  
    $global:Selected = $Selected2
    $textBox_List.text = $Selected | Format-Table -AutoSize | Out-String
    $form.refresh

 })


  $button_Cancel.Add_Click(
{

    $global:Selected = $global:BaseMot
    $Pattern = ""

    $textBox_List.text = $Selected | Format-Table -AutoSize | Out-String

    $form.refresh

 })



 
 
# Components
#-----------------------------------------------

# Adding components
$form.Controls.Add($label_Title)
$form.Controls.Add($label_Explain)

$form.Controls.Add($textBox_List)

$form.Controls.Add($label_Begin)
$form.Controls.Add($textbox_Begin)
$form.Controls.Add($button_Begin)

$form.Controls.Add($label_Select)
$form.Controls.Add($textbox_Select)
$form.Controls.Add($button_Select)

$form.Controls.Add($label_UnSelect)
$form.Controls.Add($textbox_UnSelect)
$form.Controls.Add($button_UnSelect)

$form.Controls.Add($label_Length)
$form.Controls.Add($textbox_Length)
$form.Controls.Add($button_Length)


$form.Controls.Add($button_Cancel)
$form.Controls.Add($button_quit)


# Diplaying Windows
$form.ShowDialog()