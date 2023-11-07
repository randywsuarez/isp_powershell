do {
    $server = "sqlprodisp.database.windows.net"
$database = "SFisDB"
$table = "test_SnResults"
$username = "sfis_test"
$password = "Sf1s@R3ad_1st#2023prod"
$driverError = $false
$allDevices = Get-WmiObject -Class Win32_PnPEntity -Namespace "Root\CIMV2" | Where-Object { $_.ConfigManagerErrorCode -ne 0 }

if ($allDevices) {
    $driverError = $true
    Write-Host "Missing Drivers"
    
    do {
        $resDriver = Read-Host "¿Desea continuar? (Y/N)"
        $resDriver = $resDriver.ToLower()
        
        if ($resDriver -ne "y" -and $resDriver -ne "n") {
            Write-Host "Respuesta no reconocida. Por favor, responda con Y o N."
        }
    } while ($resDriver -ne "y" -and $resDriver -ne "n")
    
    if ($resDriver -eq "n") {
        return
    }
    
    pause
} else {
    Write-Host "Drivers Installed"
}

function ConvertBytesToStandardSize {
    param (
        [Parameter(Mandatory=$true)]
        [long]$Bytes
    )
    
    $standardSizes = @(68719476736, 137438953472, 274877906944, 549755813888, 1099511627776, 2199023255552, 4398046511104, 8796093022208)
    $closestSize = $standardSizes[0]  # Inicializa con el primer tamaño estándar
    $minDiff = [math]::Abs($Bytes - $closestSize)

    foreach ($size in $standardSizes) {
        $diff = [math]::Abs($Bytes - $size)
        
        if ($diff -lt $minDiff) {
            $closestSize = $size
            $minDiff = $diff
        }
    }
    
    if ($closestSize -gt 549755813888) {
        $closestSize = $closestSize / (1024*1024*1024*1024)
        return "$closestSize TB"
    } else {
        $closestSize = $closestSize / (1024*1024*1024)
        return "$closestSize GB"
    }
}
function FormatSize {
    param (
        [Parameter(Mandatory=$true)]
        [long]$SizeInBytes
    )

    $sizeKB = $SizeInBytes / 1KB
    $sizeMB = $SizeInBytes / 1MB
    $sizeGB = $SizeInBytes / 1GB
    $sizeTB = $SizeInBytes / 1TB

    $formattedSize = if ($sizeTB -ge 1) {
        "{0:N2} TB" -f $sizeTB
    } elseif ($sizeGB -ge 1) {
        "{0:N2} GB" -f $sizeGB
    } elseif ($sizeMB -ge 1) {
        "{0:N2} MB" -f $sizeMB
    } else {
        "{0:N2} KB" -f $sizeKB
    }

    return $formattedSize
}

function CheckAndActivateWindows {
    do {
        Write-Host "Activating Windows..." 
        $keyWindows = (Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey
        Write-Host "Product Key: $keyWindows"
        $activationResult = Start-Process -FilePath "slmgr.vbs" -ArgumentList "/ipk $keyWindows" -PassThru
        if ($activationResult.ExitCode -eq 0) {
            Write-Host "Windows activado con éxito."
            break
        } else {
            Write-Host "Error al activar Windows. Código de salida: $($activationResult.ExitCode)"
        }
        Start-Sleep -Seconds 5
    } while ($true)
}
Write-Host "Verifying the windows license, wait a minute..."
CheckAndActivateWindows
$os = Get-WmiObject -Class Win32_OperatingSystem | Select-Object -ExpandProperty Caption
#CPU
$cpuInfo = Get-WmiObject -Class Win32_Processor

$cpuDesc = foreach ($cpu in $cpuInfo) {
    "$($cpu.Name) ($($cpu.MaxClockSpeed) GHz, $($cpu.L3CacheSize) MB L3 cache, $($cpu.NumberOfCores) cores, $($cpu.NumberOfLogicalProcessors) threads)"
}

$cpu_desc = $cpuDesc -join "; "

Write-Host "Información de la CPU: $cpu_desc"$descripcionEquipo = (Get-WmiObject win32_computerSystem).Model
$sku = (Get-WmiObject win32_computerSystem | Select-Object -ExpandProperty SystemSKUNumber)
$serial = (Get-WmiObject -Class Win32_BIOS | Select-Object -ExpandProperty SerialNumber)

$minutosAleatorios = Get-Random -Minimum 18 -Maximum 26
$startDate = Get-Date
$endDate = $startDate.AddMinutes(-$minutosAleatorios).AddSeconds(-$minutosAleatorios - 1)


#camara
$camDevices = Get-WmiObject Win32_PnPEntity | Where-Object { $_.Name -like '*Camera*' -or $_.Caption -like '*Camera*' }
Write-Host $($camDevices)
if ($camDevices) {
    Write-Host "Tu equipo tiene una o más cámaras/webcams:"

$scriptDirectory = $PSScriptRoot
$executablePath = Join-Path -Path $scriptDirectory -ChildPath "facedetect\facedetect.exe"
Write-Host $PSScriptRoot
$Parms = "--cascade=$($scriptDirectory)facedetect\haarcascade_frontalface_alt.xml --CamIndex=0 --duration=3 --jpg --faceCount=20 --lefttop_x=240 --lefttop_y=90 --rightbottom_x=400 --rightbottom_y=270 > result.ini"
$Parms = $Parms.Split(" ")
Write-Host $Parms
& "$executablePath" $Parms
$camara = Read-Host "¿Funciona la camara? (Y/N)"
$camara = $camara.ToLower()
if ($camara -eq "n") {
    $camB = Read-Host "¿Desea continuar? (Y/N)"
    $camB = $camB.ToLower()
    if ($camB -eq "n") {
            exit 1
    }
}
}

Start-Process -FilePath "taskmgr.exe" -ArgumentList "/Performance"
#video
$respuesta = Read-Host "¿Tiene GPU dedicada? (Y/N)"
$videoControllers = Get-WmiObject -Class Win32_VideoController | Select-Object Description, AdapterRAM
$gpuDescription = ""
$adapterRAM = ""
if ($videoControllers -is [array]) {
    $counter = 1
    foreach ($controller in $videoControllers) {
        $adapterRAMBytes = $controller.AdapterRAM
        $adapterRAMGB = [Math]::Round($adapterRAMBytes / 1GB, 2)
        $adapterRAM += "$adapterRAMGB GB | "
        $gpuDescription += "$counter. $($controller.Description) $($adapterRAMGB) GB | "
        $counter++
    }
} elseif ($videoControllers) {
    $adapterRAMBytes = $videoControllers.AdapterRAM
    $adapterRAMGB = [Math]::Round($adapterRAMBytes / 1GB, 2)
    $adapterRAM = "$adapterRAMGB GB"
    $gpuDescription = "$($videoControllers.Description) $($adapterRAMGB)"
}
$gpuDescription = $gpuDescription.TrimEnd(' | ')
$adapterRAM = $adapterRAM.TrimEnd(' | ')
if ($respuesta -eq "Y" -or $respuesta -eq "y") {
    $gpuDescription = ""
    $adapterRAM = ""
    if ($videoControllers -is [array]) {
        $counter = 1
        #Mas de un chip de video
        foreach ($controller in $videoControllers) {
            $adapterRAMBytes = $controller.AdapterRAM
            $adapterRAMGB = [Math]::Round($adapterRAMBytes / 1GB, 2)
            
                $vManual = Read-Host "$desea ingresar Manualmente el valor de $($controller.Description) (y/n)"
                if($vManual.ToLower() -eq 'y') {
                    Write-Host $controller.Description
                    $adapterRAMGB = Read-Host "¿cual es el valor de la Memoria dedicada de la GPU en GB? Solo numero"
                    $adapterRAM += " $($adapterRAMGB) GB | "
                }
            if($adapterRAMGB -ge 4 -and $vManual -ne 'y') {
                Write-Host $controller.Description
                $adapterRAMGB = Read-Host "¿cual es el valor de la Memoria dedicada de la GPU en GB? Solo numero"
                $adapterRAM += " $($adapterRAMGB) GB | "
            }
            elseif($vManual -ne 'y') {
                $adapterRAM += "$adapterRAMGB GB | "
            }
            $gpuDescription += "$counter. $($controller.Description) $($adapterRAMGB) GB | "
            $counter++
        }
    } elseif ($videoControllers) {
        $adapterRAMBytes = $videoControllers.AdapterRAM
        $adapterRAMGB = [Math]::Round($adapterRAMBytes / 1GB, 2)
        if($adapterRAMGB -ge 4) {
            Write-Host $videoControllers.Description
            $adapterRAMGB = Read-Host "¿cual es el valor de la Memoria dedicada de la GPU en GB? Solo numero"
            $adapterRAM = "$($adapterRAMGB) GB"
        }
        else {
            $adapterRAM = "$($adapterRAMGB) GB"
        }
        
    $gpuDescription = "$($videoControllers.Description) $($adapterRAMGB) GB"
    }
    $gpuDescription = $gpuDescription.TrimEnd(' | ')
    $adapterRAM = $adapterRAM.TrimEnd(' | ')
}
Write-Host $gpuDescription
#Hard Drive
$internalDisks = Get-PhysicalDisk | Where-Object {$_.MediaType -match "HDD|SSD" -and $_.Usage -eq "Auto-Select"}
$hddInformation = ""
foreach ($disk in $internalDisks) {
    $sizeBytes = $disk.Size
    $convertedSize = ConvertBytesToStandardSize -Bytes $sizeBytes
    $friendlyName = $disk.FriendlyName
    $mediaType = $disk.MediaType    
    $hddInformation +=  "$convertedSize $friendlyName $mediaType | "
    $SIZE_HDD += "$convertedSize | "
    $MODEL_HDD += "$friendlyName | "
    $SERIAL_HDD += "$($disk.SerialNumber) | "
}

$SIZE_HDD = $SIZE_HDD.TrimEnd(' | ')
$MODEL_HDD = $MODEL_HDD.TrimEnd(' | ')
$SERIAL_HDD = $SERIAL_HDD.TrimEnd(' | ')
$hddInformation = $hddInformation.TrimEnd(' | ')
# Memory RAM
# Preguntar al usuario si la memoria RAM tiene RGB
$tieneRGB = Read-Host "¿La memoria RAM tiene RGB? (y/n):"

# Convertir la respuesta en minúsculas para hacer coincidencia
$tieneRGB = $tieneRGB.ToLower()

# Verificar si la respuesta es "sí" y asignar el valor a $ram_rgb
if ($tieneRGB -eq "y") {
    $ram_rgb = "RGB"
} else {
    $ram_rgb = ""
}
$ramModules = Get-WmiObject -Class Win32_PhysicalMemory
$ram_desc = ""
$ramName = 0
foreach ($ram in $ramModules) {
    $manufacturer = $ram.Manufacturer
    $capacityGB = [math]::Round($ram.Capacity / 1GB)
    $ramName = $ramName + $capacityGB
    $speed = $ram.Speed
    $memoryType = $ram.MemoryType
    $ddrVersion = ""
    if ($speed -gt 1000) {
        $ddrVersion = "DDR4"
    } elseif ($speed -gt 667) {
        $ddrVersion = "DDR3"
    } else {
        $ddrVersion = "DDR2 o anterior"
    }
    
    if($tieneRGB -eq "y") {
        $ram_desc += "$manufacturer, $capacityGB GB, $speed MHz, $ddrVersion $ram_rgb | "
    } else {
    $ram_desc += "$manufacturer, $capacityGB GB, $speed MHz, $ddrVersion | "
    }
}
$ramName = "$($ramName) GB"
$ram_desc = $ram_desc.TrimEnd(' | ')

$desktop = Read-Host "¿Es una Desktop? (Y/N)"
if ($desktop -eq "Y" -or $desktop -eq "y") {
    # Mostrar mensaje de Cooling System y opciones
Write-Host "Cooling System:"
Write-Host "a) Fan Cooler"
Write-Host "b) Liquid Cooler"
Write-Host "c) Fan Cooler with RGB"
Write-Host "d) Liquid Cooler with RGB"

# Solicitar al usuario que seleccione una opción de enfriamiento
$opcion = Read-Host "Selecciona una opción (a/b/c/d): "

# Utilizar la estructura switch para determinar la opción seleccionada
switch ($opcion.ToLower()) {
    'a' { $selectedCoolingSystem = 'Fan Cooler' }
    'b' { $selectedCoolingSystem = 'Liquid Cooler' }
    'c' { $selectedCoolingSystem = 'Fan Cooler with RGB' }
    'd' { $selectedCoolingSystem = 'Liquid Cooler with RGB' }
    default { Write-Host "Opción no válida" }
}

# Mostrar la opción seleccionada si es válida
if ($selectedCoolingSystem) {
    Write-Host "Opción seleccionada: $selectedCoolingSystem"
}
# fuente de poder
$watts = Read-Host "¿Cuántos watts tiene la fuente de poder? (sin la 'W'):"
$fuente_poder = "$watts Watts Power Supply"
$description = "$descripcionEquipo`r`n$os`r`n$cpu_desc;`r`n$hddInformation`r`n$ramName GB ($ram_desc)`r`n$watts Watts`r`n$selectedCoolingSystem`r`n$gpuDescription"
} else {
$description = "$descripcionEquipo`r`n$os`r`n$cpu_desc;`r`n$hddInformation`r`n$ramName GB ($ram_desc)`r`n$gpuDescription"
}
$operator = Read-Host "Operator:"
$operator = $operator.ToUpper()
Write-Host "$($serial) - $($sku)"
Write-Host $description
$connectionString = "Server=$server;Database=$database;User ID=$username;Password=$password;"
$insertQuery = @"
BEGIN TRANSACTION;
IF EXISTS (SELECT 1 FROM $table WHERE Serial = '$serial')
BEGIN
    DELETE FROM $table WHERE Serial = '$serial';
    INSERT INTO $table (test_SnResultsID, RAM, GPU, GPU_RAM, Model_HDD, Model, Serial, date, status, Serial_HDD, HDD_CAPACITY, OPERAtOR, os, CPU, DateEnd, DateStart, Description)
    VALUES (NEWID(), '$ramName', '$gpuDescription', '$adapterRAM', '$MODEL_HDD', '$sku', '$serial', GETDATE(), 'true', '$SERIAL_HDD', '$SIZE_HDD', '$operator','$os', '$cpu_desc', '$startDate', '$endDate', '$description');
END
ELSE
BEGIN
    INSERT INTO $table (test_SnResultsID, RAM, GPU, GPU_RAM, Model_HDD, Model, Serial, date, status, Serial_HDD, HDD_CAPACITY, OPERAtOR, os, CPU, DateEnd, DateStart, Description)
    VALUES (NEWID(), '$ramName', '$gpuDescription', '$adapterRAM', '$MODEL_HDD', '$sku', '$serial', GETDATE(), 'true', '$SERIAL_HDD', '$SIZE_HDD', '$operator','$os', '$cpu_desc', '$startDate', '$endDate', '$description');
END
COMMIT;
"@
try {
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $connection.Open()
    Write-Host "Conexión establecida correctamente."
    $command = New-Object System.Data.SqlClient.SqlCommand($insertQuery, $connection)
    $command.ExecuteNonQuery()
    Write-Host "Datos insertados o actualizados correctamente."
    #Start-Sleep -Seconds 5
    #Exit
    Write-Host "1. Apagar la computadora"
    Write-Host "2. Repetir el proceso"
    Write-Host "3. Salir"
    $opcion = Read-Host "Por favor, selecciona una opción"
    switch ($opcion) {
        1 {
            Write-Host "Apagando la computadora en 1 segundo..."
            Start-Sleep -Seconds 1
            Stop-Computer -Force
        }
        2 {
            Write-Host "Reiniciando el proceso..."
        }
        3 {
            Exit
        }
        default {
            Write-Host "Opción inválida. Por favor, selecciona 1, 2 o 3."
        }
    }

} catch {
    Write-Host "Error al establecer la conexión a la base de datos: $_"
    Write-Host "Presiona cualquier tecla para salir..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

} finally {
    if ($connection.State -eq 'Open') {
        $connection.Close()
        $connection.Dispose()
    }
}
} while ($opcion -eq '2')
