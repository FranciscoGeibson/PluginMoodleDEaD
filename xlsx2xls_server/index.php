<?php

$allowed_ips = ['127.0.0.1', '::1', '172.17.0.1'];
$allowed_token = 'segredo123';

$client_ip = $_SERVER['REMOTE_ADDR'];
$headers = getallheaders();

if (!in_array($client_ip, $allowed_ips)) {
    http_response_code(403);
    echo "Acesso negado: IP não autorizado.";
    exit;
} else if (!isset($headers['X-Auth-Token']) || $headers['X-Auth-Token'] !== $allowed_token) {
    http_response_code(403);
    echo "Acesso negado: token inválido.";
    exit;
} else if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['xlsx_file'])) {
    $maxExecutionTime = 300; // 5 minutos

    // Aumentar tempo de execução
    set_time_limit($maxExecutionTime);

    // Limitar tamanho do arquivo
    $maxSize = 200 * 1024; // 200 KB
    if ($_FILES['xlsx_file']['size'] > $maxSize) {
        http_response_code(413); // Payload Too Large
        echo "Arquivo muito grande. Tamanho máximo permitido: 200 KB.";
        exit;
    }

    // Verifica MIME
    $finfo = finfo_open(FILEINFO_MIME_TYPE);
    $mime = finfo_file($finfo, $_FILES['xlsx_file']['tmp_name']);
    finfo_close($finfo);

    if ($mime !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
        http_response_code(400);
        echo "Por favor, envie apenas arquivos XLSX válidos.";
        exit;
    }

    // Sanitizar nome do arquivo
    $originalName = basename($_FILES['xlsx_file']['name']);
    $sanitizedFilename = preg_replace('/[^a-zA-Z0-9_\-\.]/', '', $originalName);
    $fileExt = strtolower(pathinfo($sanitizedFilename, PATHINFO_EXTENSION));

    if ($fileExt !== 'xlsx') {
        http_response_code(400);
        echo "Extensão de arquivo inválida.";
        exit;
    }

    // Criar diretório temporário exclusivo
    $tmpDir = '/tmp/xlsx_to_xls_' . bin2hex(random_bytes(8)) . '/';
    if (!mkdir($tmpDir, 0700, true)) {
        http_response_code(500);
        echo "Erro ao criar diretório temporário.";
        exit;
    }

    // Mover o arquivo para o diretório temporário
    $xlsxPath = $tmpDir . $sanitizedFilename;
    if (!move_uploaded_file($_FILES['xlsx_file']['tmp_name'], $xlsxPath)) {
        http_response_code(400);
        echo "Erro ao salvar o arquivo enviado.";
        rmdir($tmpDir);
        exit;
    }

    // Nome do arquivo de saída
    $xlsFilename = pathinfo($sanitizedFilename, PATHINFO_FILENAME) . '.xls';
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

    // Espera com timeout
    $timeout = time() + $maxExecutionTime;
    while (!file_exists($xlsPath) && time() < $timeout) {
        usleep(500000); // Espera 0.5 segundo
    }

    if (file_exists($xlsPath)) {
        // Envia o arquivo convertido
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename="' . basename($xlsFilename) . '"');
        header('Content-Length: ' . filesize($xlsPath));
        readfile($xlsPath);

        // Limpeza
        unlink($xlsxPath);
        unlink($xlsPath);
        rmdir($tmpDir);
        exit;
    }

    // Falha na conversão
    http_response_code(500);
    echo "Falha na conversão.\n";
    echo "Status: $status\nSaída:\n" . implode("\n", $output);

    // Limpeza
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