# Определить права на выполнение скриптов (включать если ограничены права учетной записи и ограничен запуск скриптов PS)
#$ExecutionPolicyStatus = get-ExecutionPolicy
#Write-Host "Предустановленное разрешение на выполнение сценариев:" $ExecutionPolicyStatus
#if ($ExecutionPolicyStatus -ne "Unrestricted")
#{
#Set-ExecutionPolicy -Scope CurrentUser Unrestricted -Force
#Write-Host "Разрешение на выполнение сценариев изменено на:" $ExecutionPolicyStatus 
#}

# Ярлык для запуска этого скриптa: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -executionpolicy bypass -file "\\SERVER\_Шаблоны\nCAD Templates\Configs\nc_pRpR_Setup.ps1"

$PSScriptRoot # Путь до каталога откуда запущен данный скрипт

$wshell = New-Object -ComObject Wscript.Shell # Для использования в скрипте всплывающих окон

$ncDefaultVer = "24.0" # Версия нанокада по умолчанию

# Заголовок-шапка окна установщика
Write-Host "`n ---- Автоматическая настройка nanoCAD для пользователя" $env:UserName " ---- `n"
Write-Host "_______________________________________________________________________________________"
Write-Host " Для начала закройте все запущенные экземпляры nanoCAD."
Write-Host ""

$ncVer = Read-Host -Prompt " Введите обозначение номера версии nanoCAD которую необходимо настроить (например 24.1). `n [Если ваша версия 24.0 то просто нажмите на клавиатуре клавишу <Enter>]"
if ($ncVer -eq ""){$ncVer = $ncDefaultVer}

# Автозамена запятой на точку в обозначении введеной версии
if ($ncVer.Contains(",")){$ncVer = $ncVer.Replace(",", ".")}
$ncVer = $ncVer -as [double] -as [string]
if (($ncVer.Contains(".") -ne $true)) {$ncVer = $ncVer + ".0"} # Если версия введена без уточнения после знака разделителя то к версии добавит .0 по умолчанию
if($ncVer -eq ""){$wshell.Popup("Не верно указана версия! Укажите версию в формате хх.хх", 0,"Ошибка")}

else{
Write-Host "Выбрана версия" $ncVer
}

$ncExePath = "C:\Program Files\Nanosoft\nanoCAD x64 $ncVer\ncad.exe" # Путь к exe - файлу nanocad
Start-Process -FilePath $ncExePath -A "/register" # Регистрация нанокад выбранной версии. Это нужно для нормальонй работы нанокада с ActiveX COM

$ncConfigFolderName = "\Config\" # Имя подкаталога в который необходимо копировать файлы настроек
$ncConfigFileName = "cfg.ini" # Имая файла-настройки для копирования и замены
# $ncIconFileName = "myIcons.dll" # Имая дополнительного файла с иконками для панели инструментов для копирования
$ncConfigFolderPath = $env:USERPROFILE + "\AppData\Roaming\Nanosoft\nanoCAD x64 " + $ncVer + $ncConfigFolderName # Полный путь до каталога пользователя с его настройками nanoCAD (с учетом выбранной версии)
$fullConfigFilePath = $ncConfigFolderPath + $ncConfigFileName # Полный путь до файла (заменяемого/копируемого) 
# $fullIconFilePath = $ncConfigFolderPath + $ncIconFileName # Полный путь до файла с иконками

# Полный путь к настроенному конфиг-файлу (для локального установщика путь можно изменить на относительный - в каталог с установщиком, например - $PSScriptRoot+"\ConfigFiles" + $ncConfigFileName)
$sourceFullFilePath = "\\SERVER\_Шаблоны\nCAD Templates\Configs\Roaming_Nanosoft_nanoCAD_Config\" + $ncConfigFileName
# $sourceFullFilePath = $PSScriptRoot+"\ConfigFiles" + $ncConfigFileName # Варант для локального установщика

if (Test-Path -Path $fullConfigFilePath) # Проверка на наличие заданного каталога и файла
{
Write-Host "Каталог версии" $ncVer "обнаружен"

$Output = $wshell.Popup("Каталог версии " + $ncVer + " обнаружен. " + "`nВыполнить копирование файла нстроек?", 0,"Подтверждение", 4+32)
switch ($Output) { 
7 { $wshell.Popup('Действие отменено', 0,"Отмена") } 
6 { 

Copy-Item -Path $sourceFullFilePath -Destination $fullConfigFilePath -Force # Копирование файла в папку назначения с заменой файла если он там уже есть
# Copy-Item -Path $sourceFullFilePath -Destination $fullIconFilePath -Force # Копирование файла в папку назначения (с заменой файла если он там уже есть)
# Так же можно копирнуть например файл базы пользовательских элементов/объектов "MyDB.MCDI", для последубщего его импорта в БД

if (Test-Path -Path $fullConfigFilePath){ $wshell.Popup("Копирование настроек $ncConfigFileName для версии $ncVer успешно выполнено!", 0,"Выполнено") } # Проверка на фактическое наличие файла
else {$wshell.Popup("Копирование настроек не удалось!", 0,"Ошибка")}

} 
default {$wshell.Popup('Неверный ввод')} 
} 

Write-Host "Успешно скопировано из каталога: " $sourceFullFilePath
Write-Host "в каталог: " $ncConfigFolderPath

}

else {
Write-Host "Каталог версии"  $ncVer "не обнаружен. Укажите версию nanoCAD которая установлена на данном ПК"
$wshell.Popup("Не верно указана версия! Укажите версию в формате хх.хх", 0,"Ошибка")
}


# -------------  Импорт reg-файла в реестр с настройками для нанокада (так же возможно выполнять импорт непосредственно в теле кода этого скрипта, но из reg-файла удобнее. Если права пользователя ограничены то винда запросит ввести пароль пользователя)
$regFileName = "nCAD_TemplateDocs+ToolPalette.reg" # Имя reg-файла с настройками nanoCAD
$regFilePath = "\\SERVER\_Шаблоны\nCAD Templates\Configs\Настройка палитр и шаблонов\" + $regFileName # Путь к каталогу с reg-файлом (так же можно заменить на относительный путь например - $PSScriptRoot+"\RegFiles")
$regTreePath = “HKCU:\SOFTWARE\Nanosoft\nanoCAD x64\"+$ncVer+"\Profiles\SPDS\TemplateDoc” 
$regTree = Get-ItemProperty –Path $regTreePath

$nowDate = Get-Date -Format d
$updateTreePath = "HKCU:\SOFTWARE\Nanosoft\nanoCAD x64\"+$ncVer+"\Profiles\SPDS\" # Это путь для фиксации даты обновления с помощью этой утилиты
$updateTree = Get-ItemProperty –Path $updateTreePath
$updateDay = "updateDay"

$Output = $wshell.Popup("Выполнить нстройку шаблонов *.dwt и панели инструментов? `n `nВ следующем окне возможно потребуется указать пароль от вашей стандартной учетной записи", 0,"Подтверждение", 4+32)
switch ($Output) { 
7 { $wshell.Popup('Действие отменено', 0,"Отмена") } 
6 { 

Start-Process -filepath "$env:windir\regedit.exe" -Argumentlist @("/s", "`"$regFilePath`"") # выполнение импорта из reg-файла
New-ItemProperty -Path $updateTreePath -Name "updateDay" -Value $nowDate -PropertyType STRING -Force # Запись даты обновления

# ------- Тут можно выполнить проверку на факт записи ключа в реестр ------
#$updateTree = Get-ItemProperty –Path $updateTreePath # Обязательно обновить содержимое ветки

# Проверка на факт успешного выполнения импорта reg-файла
#if(($regTree.updateCheck -eq "ok") -and ($updateTree.updateDay -eq $nowDate)){ 
#$wshell.Popup("Настройка nanoCAD выполнена успешно!", 0,"Выполнено")
#}
#else{
#Write-Host "Импорт настроек не был выполнен" #$wshell.Popup('Импорт настроек не был выполнен', 0,"Отказ")
#}
# ------- Тут можно выполнить проверку на факт записи ключа в реестр ------

} 
default {$wshell.Popup('Неверный ввод', 0,"Ошибка")} 
} 


Write-Host "---------------------------------------"

$wshell.Popup("Настройка nanoCAD выполнена успешно!", 0,"Выполнено")
