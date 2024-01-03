<#
Desenvolvido por: Paulo Ricardo Ferreira
Data: 10/11/2023
Objetivo: Criar todas as conexões DSN que podem ser utilizatas pelo PowerBI ou Excel
#>

<# Variaveis #>
$NomeDriver = 'Oracle em OraClient12Home1'
$PlataformaDriver = '64-bit'
$PastaTNS = "C:\tns"
$URLOdbc = "https://github.com/pauloricardoferreira/oracle_odbc_driver/archive/refs/heads/master.zip"
$FILEOdbc = "odbc_oracle_client-master"

<# Configurando TNS #>

<# Captura o Nome do Driver ODBC da Oracle#>
$DriverOracle = Get-OdbcDriver | Where-Object { ($_.Name -like "$NomeDriver") -and ($_.Platform -eq "$PlataformaDriver") }

function ConfigurarDriver {
    Write-Host "Criando Pasta em: $PastaTNS"
    New-Item -Path $PastaTNS -ItemType Directory -Force > $null

    Write-Host "Definindo Variavel de Ambiente TNS_ADMIN"
    [Environment]::SetEnvironmentVariable("TNS_ADMIN",$PastaTNS, "User")

    Write-Host "Copiando Arquivo TNS para a pasta $PastaTNS"
    Copy-Item .\tnsnames.ora -Destination $PastaTNS -Force

    <# Lista de Conexões para Configurar#>
    $NomesDSN = @(
        New-Object PSObject -Property @{NomeDSN = "BASEEBS"; DriverDSN = "EBSPRD" }
    )

    <# Configura as conexões #>

    foreach ($i in $NomesDSN) {   
        $Desc = $i.NomeDSN
        $TNS = $i.DriverDSN
        $Nome = $i.NomeDSN
        $DriverOracleNome = $DriverOracle.Name

        # Remove-OdbcDsn -Name $i.NomeDSN -DsnType "User" -ErrorAction SilentlyContinue

        # Write-Host "Criando Conexão:" $Nome
        # Add-OdbcDsn -Name $i.NomeDSN -DriverName $DriverOracleNome -DsnType "User" -SetPropertyValue @("Description=$Desc", "ServerName=$TNS") -PassThru > $null

        # Write-Host $Teste
        if ($null -ne (Get-OdbcDsn -Name $Nome -ErrorAction Ignore)){
            Write-Host "Atualizando Conexão:" $Nome
            Set-OdbcDsn -Name $i.NomeDSN -DsnType "User" -SetPropertyValue @("Description=$Desc", "ServerName=$TNS") -PassThru > $null
        }
        else {
            Write-Host "Criando Conexão:" $Nome
            Add-OdbcDsn -Name $i.NomeDSN -DriverName $DriverOracleNome -DsnType "User" -SetPropertyValue @("Description=$Desc", "ServerName=$TNS") -PassThru > $null
        }
    }
}


function ExtrairArquivo {

    Expand-Archive $HOME\Downloads\$FILEOdbc".zip" -DestinationPath $HOME\Downloads\. -Force -ErrorAction  SilentlyContinue

    Start-Process -FilePath $HOME\Downloads\$FILEOdbc\setup.exe -PassThru
    
}


function DownloadArquivo{
    Write-Host "Driver Não encontrado"
    Write-Warning "Iniciando Download do Arquivo"

    $CHROME = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe' -ErrorAction SilentlyContinue

    if ($null -ne $CHROME){
        [system.Diagnostics.Process]::Start("chrome",$URLOdbc)
    }
    else{
        [system.Diagnostics.Process]::Start("msedge",$URLOdbc)
    }

}

function RemoverArquivo {
    Write-Host "Removendo arquivos temporários"
    Remove-Item $HOME\Downloads\$FILEOdbc".zip" -ErrorAction SilentlyContinue
    Remove-Item $HOME\Downloads\$FILEOdbc -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
}


function Principal {
    <# Verificar se o Driver não está instaldo e se o arquivo ainda não baixado#>
    if ($null -eq $DriverOracle){
        if (!(Test-Path $HOME\Downloads\$FILEOdbc".zip")){
            RemoverArquivo
            DownloadArquivo

            while (!(Test-Path $HOME\Downloads\$FILEOdbc".zip")) {
                Write-Host "Aguardando o Download Finalizar"
                Start-Sleep 5
            }

        }
        
        ExtrairArquivo
    }

    Write-Host "Driver já está instalado"

    ConfigurarDriver
    
}

# Chama a função principal
Principal

Write-Warning "Preciione Enter para Continuar" -WarningAction Inquire