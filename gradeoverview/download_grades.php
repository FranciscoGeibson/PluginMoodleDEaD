<?php
// Inclui a configuração do Moodle
require_once(__DIR__ . '/../../config.php'); 
require_login(); // Certifica-se de que o usuário está logado

$courseid = required_param('courseid', PARAM_INT); // Obtém o ID do curso

// Certifica-se de que o usuário tem permissão para acessar o curso
$context = context_course::instance($courseid);
require_capability('moodle/grade:viewall', $context);

global $DB;

try {
    // Configura o cabeçalho do arquivo para download
    header('Content-Type: text/csv');
    header('Content-Disposition: attachment;filename="notas_alunos.csv"');

    // Cria o arquivo CSV
    $output = fopen('php://output', 'w');

    // Escreve o cabeçalho da planilha
    fputcsv($output, ['Nome', 'Sobrenome', 'Nota Final']);

    // Consulta as notas dos alunos
    $sql = "SELECT u.firstname, u.lastname, g.finalgrade
            FROM {user} u
            JOIN {grade_grades} g ON u.id = g.userid
            JOIN {grade_items} gi ON g.itemid = gi.id
            WHERE gi.courseid = :courseid";

    $params = ['courseid' => $courseid];
    $grades = $DB->get_records_sql($sql, $params);

    if ($grades) {
        // Escreve as notas no arquivo CSV
        foreach ($grades as $grade) {
            fputcsv($output, [$grade->firstname, $grade->lastname, $grade->finalgrade]);
        }
    } else {
        // Caso não existam notas, escreve uma linha informativa
        fputcsv($output, ['Nenhuma nota disponível']);
    }

    fclose($output);
    exit;

} catch (Exception $e) {
    // Se ocorrer um erro, exibe uma mensagem de erro amigável
    echo "Ocorreu um erro ao gerar o CSV: " . $e->getMessage();
    exit;
}
