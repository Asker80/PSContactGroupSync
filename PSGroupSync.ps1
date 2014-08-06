param(
    [switch] $CreateLocalGroups,
    [switch] $Test,
    [switch] $Sync
)

function Get-ScriptPath
{
    try
    {
        Split-Path $myInvocation.ScriptName
    }
    catch
    {
        "."
    }
}

$settings = [xml](gc (Join-Path (Get-ScriptPath) "PSGroupSync.xml"))

# No Switches
if ($Test -eq $false -and $Sync -eq $false -and $CreateLocalGroups -eq $false)
{
    Write-Warning 'Скрипт был запущен без ключей запуска!'
@'


Запустите скрипт с ключом -Test для проверки корректности конфигурации скрипта. Повторите запуск с ключом -Test до тех пор, пока не получите сообщение об отсутствии ошибок.

При необходимости запустите скрипт с ключом -CreateLocalGroups для создания локальных групп.

Запускайте скрипт с ключом -Sync для синхронизации членов удаленных групп с локальными.
'@
    exit
}
# -Test
if ($Test -eq $true -and $Sync -eq $false -and $CreateLocalGroups -eq $false)
{
    if ((Test-Path (Join-Path (Get-ScriptPath) "PSGroupSync.xml")) -eq $false)
    {Write-Warning 'Не найден файл конфигурации PSGroupSync.xml. Проверьте наличие файла в каталоге скрипта, а также текущий каталог должен быть каталогом скрипта.';exit}
    if ($settings.Settings.GroupSyncEntry -eq $null)
    {Write-Warning 'В конфигурации скрипта не обнаружено ни одного блока <GroupSyncEntry>. Проверьте корректность конфигурации скрипта.';exit}
    $i=0
    foreach ($settGSE in $settings.Settings.GroupSyncEntry)
    {
        if ((Test-Connection $settGSE.LocalDomainController -Count 2 -Quiet) -eq $false)
        {Write-Warning 'Локальный контролер домена недоступен. Проверьте конфигурацию скрипта и работоспособность сети.';exit}
        if ((Test-Connection $settGSE.RemoteDomainController -Count 2 -Quiet) -eq $false)
        {Write-Warning 'Удаленный контролер домена недоступен. Проверьте конфигурацию скрипта и работоспособность сети.';exit}
        $session = $null
        $session = New-PSSession -Configurationname Microsoft.Exchange –ConnectionUri $settGSE.LocalMSExchangeConnURI -Authentication NegotiateWithImplicitCredential #-Credential $user
        if ($session -eq $null)
        {Write-Warning 'Некорректная строка подключения к Exchange <LocalMSExchangeConnURI>. Проверьте корректность конфигурации скрипта.';exit}
        if ($settGSE.RemoteGroupToSync -eq $null)
        {Write-Warning "В блоке <GroupSyncEntry> №$i не обнаружено ни одного блока <RemoteGroupToSync>. Проверьте корректность конфигурации скрипта.";exit}
        Import-PSSession $Session
        if ((Get-OrganizationalUnit $settGSE.LocalSyncedGroupsOU) -eq $null)
        {Write-Warning 'Указанный в конфигурации скрипта OU для локальных групп не существует. Скорректируйте конфигурацию скрипта или создайте OU.';Remove-PSSession $Session;exit}
        foreach ($rgts in $settGSE.RemoteGroupToSync)
        {
            $arRemGroup = $null
            $arLocalGroup = $null
            $arRemGroup = Get-DistributionGroup $rgts -DomainController $settGSE.RemoteDomainController
            if ($arRemGroup -eq $null)
            {Write-Warning "Группа '$rgts' не найдена на удаленном контролере домена. Проверьте корректность конфигурации скрипта.";Remove-PSSession $Session;exit}
            $arLocalGroup = Get-DistributionGroup $rgts -DomainController $settGSE.LocalDomainController
            if ($arLocalGroup -eq $null)
            {Write-Warning "Группа '$rgts' не найдена на локальном контролере домена. Выполните PSGroupSync.ps1 -CreateLocalGroups.";Remove-PSSession $Session;exit}
        }
        
        $i++
    }
    '
    Ошибок конфигурации скрипта не обнаружено.'
    Remove-PSSession $Session;exit
}
# -CreateLocalGroups
if ($CreateLocalGroups -eq $true -and $Sync -eq $false -and $Test -eq $false)
{
    if ($settings.Settings.GroupSyncEntry -ne $null)
    {
        foreach ($settGSE in $settings.Settings.GroupSyncEntry)
        {
            if ($settGSE.RemoteGroupToSync -ne $null)
            {
                $session = $null
                $session = New-PSSession -Configurationname Microsoft.Exchange –ConnectionUri $settGSE.LocalMSExchangeConnURI -Authentication NegotiateWithImplicitCredential #-Credential $user
                Import-PSSession $Session
                foreach ($rgts in $settGSE.RemoteGroupToSync)
                {
                    $arRemGroup = $null
                    $arLocalGroup = $null
                    $arRemGroup = Get-DistributionGroup $rgts -DomainController $settGSE.RemoteDomainController
                    $arLocalGroup = Get-DistributionGroup $rgts -DomainController $settGSE.LocalDomainController
                    if ($arRemGroup -ne $null)
                    {
                        if ($arLocalGroup -eq $null)
                        {
                            New-DistributionGroup -Name $rgts -OrganizationalUnit $settGSE.LocalSyncedGroupsOU `
                            -SAMAccountName $arRemGroup.Guid.ToString() -Type "Distribution" -Alias $arRemGroup.Guid.ToString() `
                            -MemberJoinRestriction Closed -MemberDepartRestriction Closed -DomainController $settGSE.LocalDomainController
                        }
                    }
                }
                Remove-PSSession $Session
            }
        }
    }
}
# -Sync
if ($Sync -eq $true -and $CreateLocalGroups -eq $false -and $Test -eq $false)
{
    if ($settings.Settings.GroupSyncEntry -ne $null)
    {
        foreach ($settGSE in $settings.Settings.GroupSyncEntry)
        {
            if ($settGSE.RemoteGroupToSync -ne $null)
            {
                $session = $null
                $session = New-PSSession -Configurationname Microsoft.Exchange –ConnectionUri $settGSE.LocalMSExchangeConnURI -Authentication NegotiateWithImplicitCredential #-Credential $user
                Import-PSSession $Session
                foreach ($rgts in $settGSE.RemoteGroupToSync)
                {
                    $arRemGroup = $null
                    $arLocalGroup = $null
                    $arRemGroup = Get-DistributionGroup $rgts -DomainController $settGSE.RemoteDomainController
                    $arLocalGroup = Get-DistributionGroup $rgts -DomainController $settGSE.LocalDomainController
                    if ($arRemGroup -ne $null -and $arLocalGroup -ne $null)
                    {
                        $arRemGroupMembers = $null
                        $arRemGroupMembersAddrs = $null
                        $arLocalGroupMembers = $null
                        $arLocalGroupMembersAddrs = $null
                        $arRemGroupMembers = Get-DistributionGroupMember $rgts -DomainController $settGSE.RemoteDomainController
                        $arRemGroupMembersAddrs = $arRemGroupMembers | %{$_.PrimarySmtpAddress}
                        $arLocalGroupMembers = Get-DistributionGroupMember $rgts -DomainController $settGSE.LocalDomainController
                        $arLocalGroupMembersAddrs = $arLocalGroupMembers | %{$_.PrimarySmtpAddress}

                        if ($arLocalGroupMembersAddrs -ne $null)
                        {
                            $arLocalGroupMembersAddrs | %{
                                if ($arRemGroupMembersAddrs -notcontains $_)
                                {
                                    Remove-DistributionGroupMember -Member $_ -Identity $arLocalGroup.DistinguishedName -DomainController $settGSE.LocalDomainController -Confirm:$false
                                }
                            }
                        }
                        if ($arRemGroupMembersAddrs -ne $null)
                        {
                            $arRemGroupMembersAddrs | %{
                                if ($arLocalGroupMembersAddrs -notcontains $_)
                                {
                                    Add-DistributionGroupMember -Member $_ -Identity $arLocalGroup.DistinguishedName -DomainController $settGSE.LocalDomainController
                                }
                            }
                        }
                    }
                }
                Remove-PSSession $Session
            }
        }
    }
}
