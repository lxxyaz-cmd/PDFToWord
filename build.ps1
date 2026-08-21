param(
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Toolchain = 'C:\Qt\Tools\mingw810_64\bin'
$Objdump = Join-Path $Toolchain 'objdump.exe'
$VcVars = 'C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat'
$Python = 'C:\Users\ZZ\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$PopplerBin = 'C:\Users\ZZ\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin'
$BuildDir = Join-Path $ProjectRoot 'build'
$DistDir = Join-Path $ProjectRoot 'dist'
$EmbeddedToolsDir = Join-Path $BuildDir 'embedded-tools'
$env:Path = "$Toolchain;$env:Path"

foreach ($required in @($VcVars, $Objdump, $Python, (Join-Path $PopplerBin 'pdftoppm.exe'), (Join-Path $ProjectRoot 'editable_engine\editable_converter.py'))) {
    if (-not (Test-Path -LiteralPath $required)) {
        throw "Missing build dependency: $required"
    }
}

# PyInstaller and the conversion libraries live in the runtime above, which is
# refreshed from time to time. Say plainly what is missing instead of failing
# later inside PyInstaller with an unrelated message.
$MissingModules = & $Python -c "import importlib.util, sys; print(';'.join(m for m in sys.argv[1:] if importlib.util.find_spec(m) is None))" PyInstaller pymupdf pdf2docx docx
if ($LASTEXITCODE -ne 0) { throw "Cannot query the build Python: $Python" }
if ($MissingModules) {
    throw "Missing Python packages: $MissingModules. Install them first: & '$Python' -m pip install pyinstaller pymupdf pdf2docx python-docx"
}

if ($Clean) {
    if ((Resolve-Path -LiteralPath $BuildDir -ErrorAction SilentlyContinue).Path -like "$ProjectRoot*") {
        Remove-Item -LiteralPath $BuildDir -Recurse -Force
    }
    if ((Resolve-Path -LiteralPath $DistDir -ErrorAction SilentlyContinue).Path -like "$ProjectRoot*") {
        Remove-Item -LiteralPath $DistDir -Recurse -Force
    }
}

New-Item -ItemType Directory -Force -Path $BuildDir, $DistDir, $EmbeddedToolsDir | Out-Null

$EditableDist = Join-Path $BuildDir 'editable-dist'
$EditableWork = Join-Path $BuildDir 'editable-work'
& $Python -m PyInstaller --noconfirm --clean --onefile --name editable_engine `
    --exclude-module pandas --exclude-module openpyxl `
    --exclude-module matplotlib --exclude-module scipy --exclude-module PIL `
    --distpath $EditableDist --workpath $EditableWork --specpath $BuildDir `
    (Join-Path $ProjectRoot 'editable_engine\editable_converter.py')
if ($LASTEXITCODE -ne 0) { throw 'Editable conversion engine build failed' }
Copy-Item -LiteralPath (Join-Path $EditableDist 'editable_engine.exe') -Destination (Join-Path $EmbeddedToolsDir 'editable_engine.exe') -Force

# Copy only the recursive non-system dependency closure needed by pdftoppm.
$seen = @{}
$queue = [System.Collections.Generic.Queue[string]]::new()
$queue.Enqueue('pdftoppm.exe')
while ($queue.Count -gt 0) {
    $name = $queue.Dequeue()
    if ($seen.ContainsKey($name)) { continue }
    $sourcePath = Join-Path $PopplerBin $name
    if (-not (Test-Path -LiteralPath $sourcePath)) { continue }
    $seen[$name] = $true
    Copy-Item -LiteralPath $sourcePath -Destination (Join-Path $EmbeddedToolsDir $name) -Force
    & $Objdump -p $sourcePath 2>$null | Select-String 'DLL Name:' | ForEach-Object {
        $dependency = ($_.Line -split 'DLL Name:')[1].Trim()
        if (Test-Path -LiteralPath (Join-Path $PopplerBin $dependency)) {
            $queue.Enqueue($dependency)
        }
    }
}

# Generate a resource table so every runtime dependency is carried inside the
# native executable. At runtime the C++ shell extracts these files to a unique
# process-local temporary directory and removes them when it exits.
$ResourceLines = [System.Collections.Generic.List[string]]::new()
$HeaderLines = [System.Collections.Generic.List[string]]::new()
$HeaderLines.Add('#pragma once')
$HeaderLines.Add('struct EmbeddedToolResource { int id; const wchar_t* name; };')
$HeaderLines.Add('inline constexpr EmbeddedToolResource kEmbeddedToolResources[] = {')
$ResourceId = 2000
Get-ChildItem -LiteralPath $EmbeddedToolsDir -File | Sort-Object Name | ForEach-Object {
    $RelativePath = 'build\\embedded-tools\\' + $_.Name
    $ResourceLines.Add(('{0} RCDATA "{1}"' -f $ResourceId, $RelativePath))
    $EscapedName = $_.Name.Replace('\\', '\\\\').Replace('"', '\\"')
    $HeaderLines.Add(('    {{ {0}, L"{1}" }},' -f $ResourceId, $EscapedName))
    $ResourceId++
}
$HeaderLines.Add('};')
$ResourceLines | Set-Content -LiteralPath (Join-Path $BuildDir 'embedded_tools.rcinc') -Encoding ASCII
$HeaderLines | Set-Content -LiteralPath (Join-Path $BuildDir 'embedded_tools.h') -Encoding UTF8

$Source = Join-Path $ProjectRoot 'src\main.cpp'
$Resource = Join-Path $BuildDir 'app.res'
$Output = Join-Path $DistDir 'PDFToWord-Single.exe'
$CompileCommand = 'call "' + $VcVars + '" >nul && cd /d "' + $ProjectRoot +
    '" && rc.exe /nologo /fo "' + $Resource + '" app.rc && cl.exe /nologo /std:c++17 /O2 /MT /EHsc /utf-8 /W4 /I"' + $BuildDir + '" /Fo:"' + (Join-Path $BuildDir 'main.obj') + '" "' +
    $Source + '" "' + $Resource + '" /Fe:"' + $Output + '" /link /SUBSYSTEM:WINDOWS ole32.lib oleaut32.lib uuid.lib comctl32.lib shell32.lib user32.lib gdi32.lib comdlg32.lib'
& $env:ComSpec /d /s /c $CompileCommand
if ($LASTEXITCODE -ne 0) { throw 'C++ compilation failed' }

Copy-Item -LiteralPath (Join-Path $ProjectRoot 'README.md') -Destination (Join-Path $DistDir 'README.md') -Force

Write-Host "Built: $Output"
Write-Host "Embedded engine files: $($seen.Count + 1)"
Write-Host "Single-file size: $((Get-Item -LiteralPath $Output).Length) bytes"
