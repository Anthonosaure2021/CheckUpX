Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.speech
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object System.Windows.Forms.Form
$Form.ClientSize = New-Object System.Drawing.Size(300, 300)
$Form.Text = "CheckUpX by Nabu"

# Création d'un panneau déroulant pour contenir les boutons horizontaux
$ScrollPanel = New-Object System.Windows.Forms.Panel
$ScrollPanel.Location = New-Object System.Drawing.Point(0, 0)
$ScrollPanel.Size = New-Object System.Drawing.Size($Form.ClientSize.Width, 300)  # Ajustez la hauteur du panneau
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
$Bouton.Location = New-Object System.Drawing.Point($centerX, 0)
$Bouton.Width = 260
$Bouton.Height = 40
$Bouton.Text = "Gestionnaire de peripheriques"
$Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
$Form.Controls.Add($Bouton)
$Bouton.Add_Click({ Start-Process 'devmgmt.msc' })
$ScrollPanel.Controls.Add($Bouton)



# Bouton Nom du poste
$Bnom = New-Object System.Windows.Forms.Button
$Bnom.Location = New-Object System.Drawing.Point($centerX, 50)
$Bnom.Width = 260
$Bnom.Height = 40
$Bnom.Text = "Nom du poste"
$Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
$Form.Controls.Add($Bnom)
$ScrollPanel.Controls.Add($Bnom)


$Bnom.Add_Click({
    $nom_pc = [System.Environment]::MachineName

    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Résultat"
    $Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
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
$Binfos.Location = New-Object System.Drawing.Point($centerX, 100)
$Binfos.Width = 260
$Binfos.Height = 40
$Binfos.Text = "Infos complementaires"
$Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
$Form.Controls.Add($Binfos)
$ScrollPanel.Controls.Add($Binfos)


$Binfos.Add_Click({
    $infos = systeminfo 

    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Résultat infos complémentaires"
    $Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
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
$Bson.Location = New-Object System.Drawing.Point($centerX, 150)
$Bson.Width = 260
$Bson.Height = 40
$Bson.Text = "Test du haut-parleur"
$Bson.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
$Form.Controls.Add($Bson)
$ScrollPanel.Controls.Add($Bson)

$Bson.Add_Click({
    $speaker = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $speakerVolume = $speaker.Volume

    if ($speakerVolume -gt 0) {
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
    } else {
        # Gérer le cas où le haut-parleur n'est pas disponible ou est muet
        [System.Windows.Forms.MessageBox]::Show("Le haut-parleur n'est pas disponible ou est muet.", "Erreur")
    }
})

# Bouton Récupérer les adresses MAC
$Bmac = New-Object System.Windows.Forms.Button
$Bmac.Width = 260
$Bmac.Height = 40
$Bmac.Text = "Recuperer les adresses MAC"
$Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
$Bmac.Location = New-Object System.Drawing.Point($centerX, 200)
$Form.Controls.Add($Bmac)
$ScrollPanel.Controls.Add($Bmac)


$Bmac.Add_Click({
    $mac_address = Get-NetAdapter | Select ifIndex, Name, MacAddress
    
    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Recuperation des adresses MAC"
    $Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
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
$Bcitrix.Location = New-Object System.Drawing.Point($centerX, 250)
$Bcitrix.Width = 260
$Bcitrix.Height = 40
$Bcitrix.Text = "Verification de Citrix sur le poste"
$Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
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
    $Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
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
$Bsn.Location = New-Object System.Drawing.Point($centerX, 300)
$Bsn.Width = 260
$Bsn.Height = 40
$Bsn.Text = "Verification du numero de serie"
$Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
$Form.Controls.Add($Bsn)
$ScrollPanel.Controls.Add($Bsn)


$Bsn.Add_Click({
    $sn = (Get-WmiObject Win32_BIOS).SerialNumber
    
    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Numero de serie"
    $Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
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
$Bwinver.Location = New-Object System.Drawing.Point($centerX, 350)
$Bwinver.Width = 260
$Bwinver.Height = 40
$Bwinver.Text = "Afficher la version de Windows"
$Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
$Form.Controls.Add($Bwinver)
$ScrollPanel.Controls.Add($Bwinver)


$Bwinver.Add_Click({
    Write-Output "Affichage de la version de Windows"
    $Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
    Start-Process "winver.exe"
})


# Ouvre les paramètres du Micro
$Bsound = New-Object System.Windows.Forms.Button
$Bsound.Location = New-Object System.Drawing.Point($centerX, 400)
$Bsound.Width = 260
$Bsound.Height = 40
$Bsound.Text = "Ouvrir les paramètres du Micro"
$Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
$Form.Controls.Add($Bsound)
$ScrollPanel.Controls.Add($Bsound)


$Bsound.Add_Click({
    Write-Output "Ouverture des parametres du Micro"
    $Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
    $wshell = New-Object -ComObject wscript.shell
    $wshell.Run("ms-settings:sound")
})


# Liste les mises à jour de Windows sur le poste
$Bhotfix = New-Object System.Windows.Forms.Button
$Bhotfix.Location = New-Object System.Drawing.Point($centerX, 450)
$Bhotfix.Width = 260
$Bhotfix.Height = 40
$Bhotfix.Text = "Lister les mises à jour de Windows"
$Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
$Form.Controls.Add($Bhotfix)
$ScrollPanel.Controls.Add($Bhotfix)


$Bhotfix.Add_Click({
    $hotfixes = Get-HotFix
    $hotfixesText = $hotfixes | Out-String

    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Liste des mises à jour de Windows"
    $Bouton.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
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


# État de la batterie (intégrité)
$Bbattery = New-Object System.Windows.Forms.Button
$Bbattery.Location = New-Object System.Drawing.Point($centerX, 500)
$Bbattery.Width = 260
$Bbattery.Height = 40
$Bbattery.Text = "Afficher l'etat de la batterie"
$Bbattery.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
$Form.Controls.Add($Bbattery)
$ScrollPanel.Controls.Add($Bbattery)

$Bbattery.Add_Click({
    $batteryInfo = Get-WmiObject -Class Win32_Battery

    if ($batteryInfo) {
        $batteryStatus = "État de la batterie : " + $batteryInfo.Status
        $batteryStatus += "`r`n"
        $batteryStatus += "Niveau de charge : " + $batteryInfo.EstimatedChargeRemaining + "%"
    } else {
        $batteryStatus = "Aucune batterie détectée sur cet ordinateur."
    }

    [System.Windows.Forms.MessageBox]::Show($batteryStatus, "État de la batterie")
})



$Bcam = New-Object System.Windows.Forms.Button
$Bcam.Location = New-Object System.Drawing.Point($centerX, 550)  # Ajustez la position verticale selon vos besoins
$Bcam.Width = 260
$Bcam.Height = 40
$Bcam.Text = "Lancer Camera"
$Bcam.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
$Form.Controls.Add($Bcam)
$ScrollPanel.Controls.Add($Bcam)

$Bcam.Add_Click({
    $cameraDevice = Get-WmiObject -Query "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%webcam%'"

    if ($cameraDevice) {
        # Le pilote de la caméra est présent, nous pouvons lancer l'application de la caméra
        Start-Process -FilePath "microsoft.windows.camera:"

        $ResultForm = New-Object System.Windows.Forms.Form
        $ResultForm.Text = "Camera"
        $ResultForm.Size = New-Object System.Drawing.Size(500, 300)

        $ResultLabel = New-Object System.Windows.Forms.Label
        $ResultLabel.Location = New-Object System.Drawing.Point(20, 20)
        $ResultLabel.Size = New-Object System.Drawing.Size(460, 220)
        $ResultLabel.Text = "Le logiciel a bien été ouvert."
        $ResultForm.Controls.Add($ResultLabel)

        $ResultForm.ShowDialog()
    } else {
        # Gérer le cas où le pilote de la caméra n'est pas présent
        [System.Windows.Forms.MessageBox]::Show("Le pilote de la caméra n'est pas installé sur cet ordinateur.", "Erreur")
    }
})

# Suppression du dossier Diag lors de la fermeture de l'app
$Form.add_FormClosed({
    if (Test-Path $path -PathType Container) {
        Remove-Item -Path $path -Recurse -Force
    }
})



# Bouton Version du BIOS
$Bbios = New-Object System.Windows.Forms.Button
$Bbios.Location = New-Object System.Drawing.Point($centerX, 600)  # Ajustez la position verticale selon vos besoins
$Bbios.Width = 260
$Bbios.Height = 40
$Bbios.Text = "Version du BIOS"
$Bbios.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
$Form.Controls.Add($Bbios)
$ScrollPanel.Controls.Add($Bbios)

$Bbios.Add_Click({
    $biosInfo = Get-WmiObject -Class Win32_BIOS

    if ($biosInfo) {
        $biosVersion = "Version du BIOS : " + $biosInfo.Version
    } else {
        $biosVersion = "Impossible de récupérer la version du BIOS."
    }

    [System.Windows.Forms.MessageBox]::Show($biosVersion, "Version du BIOS")
})

# Bouton Vérifier les pilotes par catégorie
$BverifPilotes = New-Object System.Windows.Forms.Button
$BverifPilotes.Location = New-Object System.Drawing.Point($centerX, 650)  # Ajustez la position verticale selon vos besoins
$BverifPilotes.Width = 260
$BverifPilotes.Height = 40
$BverifPilotes.Text = "Vérifier les pilotes par catégorie"
$BverifPilotes.Padding = New-Object Windows.Forms.Padding(0) # Supprime les marges
$Form.Controls.Add($BverifPilotes)
$ScrollPanel.Controls.Add($BverifPilotes)

$BverifPilotes.Add_Click({
    # Obtient la liste des pilotes installés sur la machine
    $drivers = Get-WmiObject -Class Win32_PnPSignedDriver | Where-Object { $_.DriverDate -ne $null }

    # Crée un tableau associatif pour stocker les pilotes par catégorie
    $driverCategories = @{}

    # Parcourt chaque pilote et les classe par catégorie
    foreach ($driver in $drivers) {
        $category = $driver.DriverVersion
        if ($driverCategories.ContainsKey($category)) {
            $driverCategories[$category] += @($driver)
        } else {
            $driverCategories[$category] = @($driver)
        }
    }

    # Construit une chaîne de texte avec les pilotes classés par catégorie
    $resultText = ""
    foreach ($category in $driverCategories.Keys) {
        $resultText += "Catégorie : $category`r`n"
        foreach ($driver in $driverCategories[$category]) {
            $resultText += "   Nom : $($driver.DeviceName)`r`n"
            $resultText += "   Fabricant : $($driver.Manufacturer)`r`n"
            $resultText += "   Version : $($driver.DriverVersion)`r`n"
            $resultText += "   Date : $($driver.DriverDate)`r`n`r`n"
        }
    }

    # Crée une boîte de dialogue personnalisée avec une barre de défilement
    $ResultForm = New-Object System.Windows.Forms.Form
    $ResultForm.Text = "Pilotes par catégorie"
    $ResultForm.Size = New-Object System.Drawing.Size(600, 400)  # Ajustez la taille ici
    $ResultForm.StartPosition = "CenterScreen"  # Centrez la fenêtre sur l'écran

    $ResultTextBox = New-Object System.Windows.Forms.TextBox
    $ResultTextBox.Multiline = $true
    $ResultTextBox.ScrollBars = "Vertical"  # Ajoute une barre de défilement vertical
    $ResultTextBox.Location = New-Object System.Drawing.Point(20, 20)
    $ResultTextBox.Size = New-Object System.Drawing.Size(540, 300)
    $ResultTextBox.Text = $resultText
    $ResultForm.Controls.Add($ResultTextBox)

    $ResultForm.ShowDialog()
})
$Form.ShowDialog()

