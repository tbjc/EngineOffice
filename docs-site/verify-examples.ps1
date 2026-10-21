# Verifica que los ejemplos de la documentación compilan contra la biblioteca real.
#
# Extrae los bloques ```csharp de las páginas del sitio, los envuelve en un método y los compila
# en un proyecto temporal que referencia los proyectos de la biblioteca. Un ejemplo que no compila
# rompe la verificación.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File docs-site/verify-examples.ps1

param(
    [string]$SitePath = $PSScriptRoot,
    [string]$RepositoryPath = (Split-Path $PSScriptRoot -Parent),
    [switch]$Keep
)

$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)

$docsPath = Join-Path $SitePath 'docs'
if (-not (Test-Path $docsPath)) {
    throw "No se encuentra la carpeta de documentación en $docsPath."
}

$blocks = @()
foreach ($file in Get-ChildItem $docsPath -Recurse -Filter '*.md') {
    $text = [System.IO.File]::ReadAllText($file.FullName, $utf8)
    $matches = [regex]::Matches($text, '(?ms)^```csharp\s*\r?\n(.*?)^```')
    $index = 0
    foreach ($match in $matches) {
        $index++
        $blocks += [pscustomobject]@{
            File = $file.FullName.Substring($docsPath.Length + 1).Replace('\', '/')
            Index = $index
            Code = $match.Groups[1].Value
        }
    }
}

if ($blocks.Count -eq 0) {
    throw 'No se ha encontrado ningún ejemplo en la documentación.'
}

$work = Join-Path ([System.IO.Path]::GetTempPath()) ("officeengine-ejemplos-" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $work | Out-Null
try {
    $project = @"
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>
  <ItemGroup>
    <ProjectReference Include="$($RepositoryPath.Replace('\','/'))/src/OfficeEngine.Word/OfficeEngine.Word.csproj" />
    <ProjectReference Include="$($RepositoryPath.Replace('\','/'))/src/OfficeEngine.Excel/OfficeEngine.Excel.csproj" />
    <ProjectReference Include="$($RepositoryPath.Replace('\','/'))/src/OfficeEngine.Pdf/OfficeEngine.Pdf.csproj" />
    <ProjectReference Include="$($RepositoryPath.Replace('\','/'))/src/OfficeEngine.Pdf.LibreOffice/OfficeEngine.Pdf.LibreOffice.csproj" />
    <ProjectReference Include="$($RepositoryPath.Replace('\','/'))/src/OfficeEngine.Legacy/OfficeEngine.Legacy.csproj" />
    <ProjectReference Include="$($RepositoryPath.Replace('\','/'))/src/OfficeEngine.Formulas/OfficeEngine.Formulas.csproj" />
  </ItemGroup>
</Project>
"@
    [System.IO.File]::WriteAllText((Join-Path $work 'Ejemplos.csproj'), $project, $utf8)

    $usings = New-Object System.Collections.Generic.HashSet[string]
    $bodies = New-Object System.Text.StringBuilder
    foreach ($block in $blocks) {
        [void]$bodies.AppendLine("        // --- $($block.File) bloque $($block.Index) ---")
        [void]$bodies.AppendLine('        {')
        $pending = New-Object System.Collections.Generic.List[string]
        foreach ($line in ($block.Code -split "`r?`n")) {
            if ($line -match '^\s*using\s+[^=;]+;\s*$') {
                [void]$usings.Add($line.Trim())
                continue
            }

            $pending.Add($line)
        }

        foreach ($line in $pending) {
            [void]$bodies.AppendLine("            $line")
        }

        [void]$bodies.AppendLine('        }')
    }

    $source = New-Object System.Text.StringBuilder
    foreach ($using in ($usings | Sort-Object)) {
        [void]$source.AppendLine($using)
    }

    [void]$source.AppendLine()
    [void]$source.AppendLine('internal static class Ejemplos')
    [void]$source.AppendLine('{')
    [void]$source.AppendLine('    internal static void Main()')
    [void]$source.AppendLine('    {')
    [void]$source.Append($bodies.ToString())
    [void]$source.AppendLine('    }')
    [void]$source.AppendLine('}')
    $sourcePath = Join-Path $work 'Ejemplos.cs'
    [System.IO.File]::WriteAllText($sourcePath, $source.ToString(), $utf8)

    Write-Host "Ejemplos encontrados: $($blocks.Count)"
    $output = dotnet build (Join-Path $work 'Ejemplos.csproj') -c Release --nologo 2>&1
    if ($LASTEXITCODE -ne 0) {
        $output | ForEach-Object { Write-Host $_ }
        Write-Host "El archivo generado está en $sourcePath"
        throw 'Algún ejemplo de la documentación no compila.'
    }

    Write-Host 'Todos los ejemplos compilan.'
}
finally {
    if ($Keep) {
        Write-Host "Se conserva el trabajo en $work"
    }
    else {
        Remove-Item -Recurse -Force $work -ErrorAction SilentlyContinue
    }
}
