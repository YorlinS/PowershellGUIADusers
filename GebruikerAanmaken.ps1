#Te importeren modules
Import-Module activedirectory

#Main form code
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#Instellingen voor het hoofdformulier
$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(400,400)
$Form.text                       = "Nieuwe gebruiker aanmaken"
$Form.TopMost                    = $false
$Form.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#663366")
$Form.StartPosition              = [System.Windows.Forms.FormStartPosition]::CenterScreen

#Instellingen voor afbeelding in het formulier
$PictureBox1                     = New-Object system.Windows.Forms.PictureBox
$PictureBox1.width               = 100
$PictureBox1.height              = 100
$PictureBox1.location            = New-Object System.Drawing.Point(26,300)
$PictureBox1.imageLocation       = "C:\fontyslogo.png"
$PictureBox1.SizeMode            = [System.Windows.Forms.PictureBoxSizeMode]::zoom
$Label1                          = New-Object system.Windows.Forms.Label

#Boventitel in het formulier
$Label1.text                     = "Nieuwe gebruiker aanmaken"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(34,27)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',16,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))
$Label1.ForeColor                = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

#Label voor het vakje voornaam
$labelVoornaam                   = New-Object system.Windows.Forms.Label
$labelVoornaam.text              = "Voornaam:"
$labelVoornaam.AutoSize          = $true
$labelVoornaam.width             = 25
$labelVoornaam.height            = 10
$labelVoornaam.location          = New-Object System.Drawing.Point(37,69)
$labelVoornaam.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$labelVoornaam.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

#Invulveld voor het vakje voornaam
$tbVoornaam                      = New-Object system.Windows.Forms.TextBox
$tbVoornaam.multiline            = $false
$tbVoornaam.width                = 100
$tbVoornaam.height               = 20
$tbVoornaam.location             = New-Object System.Drawing.Point(171,65)
$tbVoornaam.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Label voor het vakje achternaam
$labelAchternaam                 = New-Object system.Windows.Forms.Label
$labelAchternaam.text            = "Achternaam:"
$labelAchternaam.AutoSize        = $true
$labelAchternaam.width           = 25
$labelAchternaam.height          = 10
$labelAchternaam.location        = New-Object System.Drawing.Point(36,116)
$labelAchternaam.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$labelAchternaam.ForeColor       = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

#Invulveld voor het vakje achternaam
$tbAchternaam                    = New-Object system.Windows.Forms.TextBox
$tbAchternaam.multiline          = $false
$tbAchternaam.width              = 100
$tbAchternaam.height             = 20
$tbAchternaam.location           = New-Object System.Drawing.Point(171,112)
$tbAchternaam.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Label voor het vakje functie
$labelFunctie                    = New-Object system.Windows.Forms.Label
$labelFunctie.text               = "Functie"
$labelFunctie.AutoSize           = $true
$labelFunctie.width              = 25
$labelFunctie.height             = 10
$labelFunctie.location           = New-Object System.Drawing.Point(37,165)
$labelFunctie.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$labelFunctie.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

#Invulveld voor het vakje functie
$tbFunctie                       = New-Object system.Windows.Forms.TextBox
$tbFunctie.multiline             = $false
$tbFunctie.width                 = 100
$tbFunctie.height                = 20
$tbFunctie.location              = New-Object System.Drawing.Point(171,161)
$tbFunctie.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Label voor het vakje afdeling
$labelAfdeling                   = New-Object system.Windows.Forms.Label
$labelAfdeling.text              = "Afdeling"
$labelAfdeling.AutoSize          = $true
$labelAfdeling.width             = 25
$labelAfdeling.height            = 10
$labelAfdeling.location          = New-Object System.Drawing.Point(36,221)
$labelAfdeling.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$labelAfdeling.ForeColor         = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

#Invulveld voor het vakjke afdeling
$tbAfdeling                      = New-Object system.Windows.Forms.TextBox
$tbAfdeling.multiline            = $false
$tbAfdeling.width                = 100
$tbAfdeling.height               = 20
$tbAfdeling.location             = New-Object System.Drawing.Point(172,217)
$tbAfdeling.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Knop voor het aanmaken van de gebruiker
$createButton                    = New-Object system.Windows.Forms.Button
$createButton.text               = "Maak aan"
$createButton.width              = 100
$createButton.height             = 30
$createButton.location           = New-Object System.Drawing.Point(171,263)
$createButton.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',12,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$createButton.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")
$createButton.BackColor          = [System.Drawing.ColorTranslator]::FromHtml("#e5007d")

#Label waarin de status wordt weergegeven voor het aanmaken van een gebruiker
$labelStatus                     = New-Object system.Windows.Forms.Label
$labelStatus.AutoSize            = $false
$labelStatus.width               = 162
$labelStatus.height              = 50
$labelStatus.location            = New-Object System.Drawing.Point(162,329)
$labelStatus.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',32,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$labelStatus.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#00ff00")

$Form.controls.AddRange(@($PictureBox1,$Label1,$labelVoornaam,$tbVoornaam,$labelAchternaam,$tbAchternaam,$labelFunctie,$tbFunctie,$labelAfdeling,$tbAfdeling,$createButton, $labelStatus))

#Knop om formulier door te sturen
$createButton.Add_Click({ 

#Voor- en achternaam van de nieuwe gebruiker 
$voornaam      = $tbVoornaam.Text
$achternaam    = $tbAchternaam.Text

#Voornaam en Achternaam omzetten naar een gebruikersnaam
$eerstelettervoornaam         = $voornaam.Substring(0, 1)
$gebruikersnaam               = "$eerstelettervoornaam.$achternaam"
$gebruikersnaam               = $gebruikersnaam -replace '\s',''
$gebruikersnaamklein          = $gebruikersnaam.ToLower()

#Variabele aanmaken voor het email adres van de gebruiker
$email = $gebruikersnaamklein + "@yssem3.local"

#Variabelen voor functie en afdeling 
$functie       = $tbFunctie.Text
$afdeling      = $tbAfdeling.Text

#Welke map moet de gebruiker inkomen na het aanmaken
$OU = "OU=$afdeling,DC=YSsem3,DC=local"

#Random gegenereerd wachtwoord voor het account, welke gebruikt dient te worden bij de eerste keer inloggen
$wachtwoord = -join ((33..122) | Get-Random -Count 8 | % {[char]$_})

#Controleren of gebruiker al bestaat
if (Get-ADUser -F {SamAccountName -eq $gebruikersnaamklein})
{
    #Als gebruiker al bestaat geef dan een melding
    $labelStatus.Text                = "Mislukt"
    $labelStatus.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#ff0000")
    [System.Windows.Forms.MessageBox]::Show("Het aanmaken van $gebruikersnaamklein is mislukt.`nGebruiker: $gebruikersnaamklein bestaat al`nProbeer het nogmaals!", "Gebruiker niet aangemaakt!")
}
else
{
#Maak nieuwe gebruiker aan in AD met gespecifeerde variabelen
	New-ADUser `
           -SamAccountName $gebruikersnaamklein `
           -Name "$voornaam $achternaam" `
           -GivenName $voornaam `
           -Surname $achternaam `
           -Enabled $True `
           -DisplayName "$voornaam $achternaam" `
           -Path $OU `
           -EmailAddress $email `
           -Title $functie `
           -Department $afdeling `
           -AccountPassword (convertto-securestring $wachtwoord -AsPlainText -Force) -ChangePasswordAtLogon $True

    #Als gebruiker is aangemaakt geef dan een melding
    $labelStatus.Text                = "Gelukt"
    $labelStatus.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#00ff00")
    [System.Windows.Forms.MessageBox]::Show("Het aanmaken van $gebruikersnaamklein is gelukt.`nEmail: $email`nWachtwoord: $wachtwoord", "Gebruiker $gebruikersnaamklein is aangemaakt!")
}

})


#Functie uitvoeren
[void]$Form.ShowDialog()