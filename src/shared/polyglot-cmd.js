(function () {
    'use strict';

    /** PowerShell -EncodedCommand erwartet UTF-16LE (ohne BOM) als Base64 */
    function toEncodedCommandBase64(psText) {
        const out = [];
        for (let i = 0; i < psText.length; i++) {
            const c = psText.charCodeAt(i);
            out.push(c & 0xff, (c >> 8) & 0xff);
        }
        let binary = '';
        for (let j = 0; j < out.length; j++) {
            binary += String.fromCharCode(out[j]);
        }
        return btoa(binary);
    }

    /**
     * Eine .cmd-Datei: Batch startet PowerShell mit kleinem Loader (-EncodedCommand),
     * der das eigentliche PS1 aus derselben Datei (nach Marker) extrahiert, nach %TEMP% schreibt und ausführt.
     * So liegt keine heruntergeladene .ps1 im Benutzerordner (weniger „Mark of the Web“-Probleme).
     *
     * @param {{ title: string, echoLine: string, psBody: string }} opts
     * @returns {string}
     */
    window.ms365BuildPolyglotCmd = function (opts) {
        const title = opts.title || 'MS365';
        const echoLine = opts.echoLine || '';
        const psBody = String(opts.psBody || '').replace(/\r\n|\n/g, '\r\n');
        const marker = '::MS365_PS_BEGIN::';

        const loaderPs = [
            '$Path = $env:MS365_SELF',
            'if ([string]::IsNullOrWhiteSpace($Path)) { Write-Error \'MS365: Umgebungsvariable MS365_SELF fehlt.\'; exit 1 }',
            '$text = [System.IO.File]::ReadAllText($Path)',
            '$marker = \'' + marker + '\'',
            '$idx = $text.IndexOf($marker)',
            'if ($idx -lt 0) { Write-Error \'MS365: Marker nicht gefunden.\'; exit 1 }',
            '$body = $text.Substring($idx + $marker.Length).TrimStart()',
            '$tmp = Join-Path $env:TEMP (\'ms365-\' + [guid]::NewGuid().ToString() + \'.ps1\')',
            '$utf8 = New-Object System.Text.UTF8Encoding $true',
            '[System.IO.File]::WriteAllText($tmp, $body, $utf8)',
            'try {',
            '  . $tmp',
            '  exit $LASTEXITCODE',
            '} finally {',
            '  Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue',
            '}'
        ].join('\r\n');

        const b64 = toEncodedCommandBase64(loaderPs);

        const batchLines = [
            '@echo off',
            'setlocal EnableExtensions',
            'if not defined SystemRoot set "SystemRoot=C:\\Windows"',
            'set "PS_EXE=%SystemRoot%\\System32\\WindowsPowerShell\\v1.0\\powershell.exe"',
            'if exist "%SystemRoot%\\Sysnative\\WindowsPowerShell\\v1.0\\powershell.exe" set "PS_EXE=%SystemRoot%\\Sysnative\\WindowsPowerShell\\v1.0\\powershell.exe"',
            'if exist "%SystemRoot%\\System32\\chcp.com" ("%SystemRoot%\\System32\\chcp.com" 65001 >nul 2>&1)',
            'title ' + title,
            'cd /d "%~dp0"',
            'set "MS365_SELF=%~f0"',
            'echo.',
            'echo ' + echoLine,
            'echo.',
            '"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -EncodedCommand ' + b64,
            'set ERR=%ERRORLEVEL%',
            'if not "%ERR%"=="0" (',
            '  echo.',
            '  echo Fehlercode: %ERR%',
            ')',
            'echo.',
            'pause',
            'exit /b',
            marker,
            psBody
        ];
        return batchLines.join('\r\n');
    };
})();

