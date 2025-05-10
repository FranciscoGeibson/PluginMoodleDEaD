<?php
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['xlsx_file'])) {
    // Configurações - usando /tmp
    $tmpDir = '/tmp/xlsx_to_xls_' . uniqid() . '/';
    $maxExecutionTime = 300; // 5 minutos

    // Criar diretório temporário
    if (!mkdir($tmpDir, 0700, true)) {
        http_response_code(500);
        echo "Erro ao criar diretório temporário.";
        exit;
    }

    // Aumentar tempo de execução
    set_time_limit($maxExecutionTime);

    // Validar arquivo
    $filename = basename($_FILES['xlsx_file']['name']);
    $fileExt = strtolower(pathinfo($filename, PATHINFO_EXTENSION));

    if ($fileExt !== 'xlsx') {
        http_response_code(400);
        echo "Por favor, envie apenas arquivos XLSX.";
        rmdir($tmpDir);
        exit;
    }

    // Mover arquivo para /tmp
    $xlsxPath = $tmpDir . $filename;
    if (!move_uploaded_file($_FILES['xlsx_file']['tmp_name'], $xlsxPath)) {
        http_response_code(400);
        echo "Erro ao salvar o arquivo enviado.";
        rmdir($tmpDir);
        exit;
    }

    // Preparar conversão
    $xlsFilename = pathinfo($filename, PATHINFO_FILENAME) . '.xls';
    $xlsPath = $tmpDir . $xlsFilename;

    // Definir HOME temporário para o LibreOffice
    putenv('HOME=/tmp');

    // Comando de conversão
    $cmd = sprintf(
        'libreoffice --headless --convert-to xls:"MS Excel 97" --outdir %s %s 2>&1',
        escapeshellarg($tmpDir),
        escapeshellarg($xlsxPath)
    );

    exec($cmd, $output, $status);

    // Verificar conversão com timeout
    $timeout = time() + $maxExecutionTime;
    while (!file_exists($xlsPath) && time() < $timeout) {
        usleep(500000); // Espera 0.5 segundo
    }

    if (file_exists($xlsPath)) {
        // Enviar arquivo convertido
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename="' . $xlsFilename . '"');
        header('Content-Length: ' . filesize($xlsPath));
        readfile($xlsPath);

        // Limpeza
        unlink($xlsxPath);
        unlink($xlsPath);
        rmdir($tmpDir);
        exit;
    }

    // Se chegou aqui, houve erro na conversão
    http_response_code(500);
    echo "Falha na conversão. ";
    echo "Status: $status, Saída: " . implode("\n", $output);

    // Limpeza em caso de erro
    if (file_exists($xlsxPath)) unlink($xlsxPath);
    if (file_exists($xlsPath)) unlink($xlsPath);
    rmdir($tmpDir);
    exit;
}
?>

<!DOCTYPE html>
<html lang='pt-BR'>

<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Conversão de XLSX para XLS</title>
</head>

<body>
    <h1>Converter XLSX para XLS</h1>
    <form method="post" enctype="multipart/form-data">
        <input type="file" name="xlsx_file" accept=".xlsx" required>
        <button type="submit">Converter</button>
    </form>
</body>

</html>