
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.speech
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object System.Windows.Forms.Form
$Form.ClientSize = New-Object System.Drawing.Size(400, 400)
$Form.Text = "Script de diagnostique"

# Création d'un panneau déroulant pour contenir les boutons horizontaux
$ScrollPanel = New-Object System.Windows.Forms.Panel
$ScrollPanel.Location = New-Object System.Drawing.Point(0, 0)
$ScrollPanel.Size = New-Object System.Drawing.Size($Form.ClientSize.Width, 200)  # Ajustez la hauteur du panneau
$ScrollPanel.AutoScroll = $true
$Form.Controls.Add($ScrollPanel)


# Calcul de la position horizontale centrée pour les boutons
$centerX = 10
$buttonSpacing = 10

$path = "C:\Diag"
If (!(Test-Path $path)) {
    New-Item -ItemType Directory -Force -Path $path
}

$resultats = "$path\$env:COMPUTERNAME.txt"

# Calcul de la position horizontale centrée
#$centerX = ($Form.ClientSize.Width - 260) / 2

# Bouton Gestionnaire de peripheriques
$Bouton = New-Object System.Windows.Forms.Button
$Bouton.Location = New-Object System.Drawing.Point($centerX, 50)
$Bouton.Width = 260
$Bouton.Height = 40
$Bouton.Text = "Gestionnaire de peripheriques"
$Form.Controls.Add($Bouton)
$Bouton.Add_Click({ Start-Process 'devmgmt.msc' })
$ScrollPanel.Controls.Add($Bouton)



# Bouton Nom du poste
$Bnom = New-Object System.Windows.Forms.Button
$Bnom.Location = New-Object System.Drawing.Point($centerX, 100)
$Bnom.Width = 260
$Bnom.Height = 40
$Bnom.Text = "Nom du poste"
$Form.Controls.Add($Bnom)
$ScrollPanel.Controls.Add($Bnom)


$Bnom.Add_Click({
    $nom_pc = [System.Environment]::MachineName

    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Résultat"
    $ResultForm.Size = New-Object System.Drawing.Size(300, 150)

    $ResultLabel = New-Object System.Windows.Forms.Label
    $ResultLabel.Location = New-Object System.Drawing.Point(20, 20)
    $ResultLabel.Size = New-Object System.Drawing.Size(260, 60)
    $ResultLabel.Text = "Le nom du poste est : $nom_pc"
    $ResultForm.Controls.Add($ResultLabel)
    $ResultForm.ShowDialog()
})


# Bouton Infos complémentaires (BIOS...)
$Binfos = New-Object System.Windows.Forms.Button
$Binfos.Location = New-Object System.Drawing.Point($centerX, 150)
$Binfos.Width = 260
$Binfos.Height = 40
$Binfos.Text = "Infos complementaires"
$Form.Controls.Add($Binfos)
$ScrollPanel.Controls.Add($Binfos)


$Binfos.Add_Click({
    $infos = systeminfo 

    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Résultat infos complémentaires"
    $ResultForm.Size = New-Object System.Drawing.Size(500, 300)

    $ResultTextBox = New-Object System.Windows.Forms.TextBox
    $ResultTextBox.Multiline = $true
    $ResultTextBox.ScrollBars = "Vertical"
    $ResultTextBox.Location = New-Object System.Drawing.Point(20, 20)
    $ResultTextBox.Size = New-Object System.Drawing.Size(460, 220)
    $ResultTextBox.Text = $infos
    $ResultForm.Controls.Add($ResultTextBox)

    $ResultForm.ShowDialog()
})

# Bouton Test du haut-parleur
$Bson = New-Object System.Windows.Forms.Button
$Bson.Location = New-Object System.Drawing.Point($centerX, 200)
$Bson.Width = 260
$Bson.Height = 40
$Bson.Text = "Test du haut-parleur"
$Form.Controls.Add($Bson)
$ScrollPanel.Controls.Add($Bson)


$Bson.Add_Click({
    $speaker = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $speaker.Speak("Ceci est un test, je fonctionne donc correctement !")

    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Test du haut-parleur"
    $ResultForm.Size = New-Object System.Drawing.Size(500, 300)

    $ResultLabel = New-Object System.Windows.Forms.Label
    $ResultLabel.Location = New-Object System.Drawing.Point(20, 20)
    $ResultLabel.Size = New-Object System.Drawing.Size(460, 220)
    $ResultLabel.Text = "Le test du haut-parleur a été effectué."
    $ResultForm.Controls.Add($ResultLabel)

    $ResultForm.ShowDialog()
})


# Bouton Récupérer les adresses MAC
$Bmac = New-Object System.Windows.Forms.Button
$Bmac.Width = 260
$Bmac.Height = 40
$Bmac.Text = "Recuperer les adresses MAC"
$Bmac.Location = New-Object System.Drawing.Point($centerX, 250)
$Form.Controls.Add($Bmac)
$ScrollPanel.Controls.Add($Bmac)


$Bmac.Add_Click({
    $mac_address = Get-NetAdapter | Select ifIndex, Name, MacAddress
    
    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Recuperation des adresses MAC"
    $ResultForm.Size = New-Object System.Drawing.Size(600, 400)

    $ResultTextBox = New-Object System.Windows.Forms.TextBox
    $ResultTextBox.Multiline = $true
    $ResultTextBox.ScrollBars = "Vertical"
    $ResultTextBox.Location = New-Object System.Drawing.Point(20, 20)
    $ResultTextBox.Size = New-Object System.Drawing.Size(540, 300)
    
    foreach ($adapter in $mac_address) {
        $ResultTextBox.Text += "Interface: $($adapter.Name)`r`n"
        $ResultTextBox.Text += "Index: $($adapter.ifIndex)`r`n"
        $ResultTextBox.Text += "Adresse MAC: $($adapter.MacAddress)`r`n`r`n"
    }
    
    $ResultForm.Controls.Add($ResultTextBox)

    $ResultForm.ShowDialog()
})


# Bouton Verification de Citrix sur le poste
$Bcitrix = New-Object System.Windows.Forms.Button
$Bcitrix.Location = New-Object System.Drawing.Point($centerX, 300)
$Bcitrix.Width = 260
$Bcitrix.Height = 40
$Bcitrix.Text = "Verification de Citrix sur le poste"
$Form.Controls.Add($Bcitrix)
$ScrollPanel.Controls.Add($Bcitrix)


$Bcitrix.Add_Click({
    $citrixPath = "C:\Program Files (x86)\Citrix\ICA Client\SelfServicePlugin\SelfService.exe"
    if (Test-Path -Path $citrixPath) {
        $citrixMessage = "Citrix est bien installe sur l'ordinateur."
    } else {
        $citrixMessage = "Citrix n'est pas installe sur l'ordinateur."
    }

    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Verification de Citrix"
    $ResultForm.Size = New-Object System.Drawing.Size(500, 300)

    $ResultLabel = New-Object System.Windows.Forms.Label
    $ResultLabel.Location = New-Object System.Drawing.Point(20, 20)
    $ResultLabel.Size = New-Object System.Drawing.Size(460, 220)
    $ResultLabel.Text = $citrixMessage
    $ResultForm.Controls.Add($ResultLabel)

    $ResultForm.ShowDialog()
})


# Bouton Verification du numero de serie
$Bsn = New-Object System.Windows.Forms.Button
$Bsn.Location = New-Object System.Drawing.Point($centerX, 350)
$Bsn.Width = 260
$Bsn.Height = 40
$Bsn.Text = "Verification du numero de serie"
$Form.Controls.Add($Bsn)
$ScrollPanel.Controls.Add($Bsn)


$Bsn.Add_Click({
    $sn = (Get-WmiObject Win32_BIOS).SerialNumber
    
    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Numero de serie"
    $ResultForm.Size = New-Object System.Drawing.Size(500, 300)

    $ResultLabel = New-Object System.Windows.Forms.Label
    $ResultLabel.Location = New-Object System.Drawing.Point(20, 20)
    $ResultLabel.Size = New-Object System.Drawing.Size(460, 220)
    $ResultLabel.Text = "Le numero de serie est : $sn"
    $ResultForm.Controls.Add($ResultLabel)

    $ResultForm.ShowDialog()
})


# Affiche la version de Windows
$Bwinver = New-Object System.Windows.Forms.Button
$Bwinver.Location = New-Object System.Drawing.Point($centerX, 450)
$Bwinver.Width = 260
$Bwinver.Height = 40
$Bwinver.Text = "Afficher la version de Windows"
$Form.Controls.Add($Bwinver)
$ScrollPanel.Controls.Add($Bwinver)


$Bwinver.Add_Click({
    Write-Output "Affichage de la version de Windows"
    Start-Process "winver.exe"
})


# Ouvre les paramètres du Micro
$Bsound = New-Object System.Windows.Forms.Button
$Bsound.Location = New-Object System.Drawing.Point($centerX, 500)
$Bsound.Width = 260
$Bsound.Height = 40
$Bsound.Text = "Ouvrir les paramètres du Micro"
$Form.Controls.Add($Bsound)
$ScrollPanel.Controls.Add($Bsound)


$Bsound.Add_Click({
    Write-Output "Ouverture des parametres du Micro"
    $wshell = New-Object -ComObject wscript.shell
    $wshell.Run("ms-settings:sound")
})


# Liste les mises à jour de Windows sur le poste
$Bhotfix = New-Object System.Windows.Forms.Button
$Bhotfix.Location = New-Object System.Drawing.Point($centerX, 600)
$Bhotfix.Width = 260
$Bhotfix.Height = 40
$Bhotfix.Text = "Lister les mises à jour de Windows"
$Form.Controls.Add($Bhotfix)
$ScrollPanel.Controls.Add($Bhotfix)


$Bhotfix.Add_Click({
    $hotfixes = Get-HotFix
    $hotfixesText = $hotfixes | Out-String

    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Liste des mises à jour de Windows"
    $ResultForm.Size = New-Object System.Drawing.Size(600, 400)

    $ResultTextBox = New-Object System.Windows.Forms.TextBox
    $ResultTextBox.Multiline = $true
    $ResultTextBox.ScrollBars = "Vertical"
    $ResultTextBox.Location = New-Object System.Drawing.Point(20, 20)
    $ResultTextBox.Size = New-Object System.Drawing.Size(540, 300)
    $ResultTextBox.Text = $hotfixesText
    $ResultForm.Controls.Add($ResultTextBox)

    $ResultForm.ShowDialog()
})


# État de la batterie
$Bbattery = New-Object System.Windows.Forms.Button
$Bbattery.Location = New-Object System.Drawing.Point($centerX, 650)
$Bbattery.Width = 260
$Bbattery.Height = 40
$Bbattery.Text = "Afficher l'etat de la batterie"
$Form.Controls.Add($Bbattery)
$ScrollPanel.Controls.Add($Bbattery)


$Bbattery.Add_Click({
    Write-Output "État de la batterie"
    $batteryReport = "C:\Diag\battery_report.html"
    powercfg /batteryreport /output $batteryReport
    Start-Process $batteryReport
})

$Bcam = New-Object System.Windows.Forms.Button
$Bcam.Location = New-Object System.Drawing.Point(10, 200)
$Bcam.Width = 260
$Bcam.Height = 40
$ScrollPanel.Controls.Add($Bcam)
$Bcam.Text = "Lancer Camera"
$Form.Controls.Add($Bcam)


$Bcam.Add_Click({
    $cam = explorer "shell:appsfolder\microsoft.windowscamera_8wekyb3d8bbwe!app"
    
    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Camera"
    $ResultForm.Size = New-Object System.Drawing.Size(500, 300)

    $ResultLabel = New-Object System.Windows.Forms.Label
    $ResultLabel.Location = New-Object System.Drawing.Point(20, 20)
    $ResultLabel.Size = New-Object System.Drawing.Size(460, 220)
    $ResultLabel.Text = "Le logiciel a bien été ouvert."
    $ResultForm.Controls.Add($ResultLabel)

    $ResultForm.ShowDialog()
})
$Form.ShowDialog()
