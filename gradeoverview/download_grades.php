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
    fputcsv($output, ['Nome', 'Sobrenome', 'Média Final']);

    // Chama o método para carregar as notas e calcular a média
    $student_grades = load_student_grades($courseid);

    // Se houver notas, escreve no arquivo CSV
    if ($student_grades) {
        foreach ($student_grades as $grade) {
            fputcsv($output, [$grade->firstname, $grade->lastname, number_format($grade->average_grade, 2)]);
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

// Função para carregar as notas dos alunos e calcular a média
function load_student_grades($courseid) {
    global $DB;
    
    // Consulta SQL para obter todas as notas das unidades do curso
    $sql = "SELECT u.id AS userid, u.firstname, u.lastname, g.finalgrade
            FROM {user} u
            JOIN {grade_grades} g ON u.id = g.userid
            JOIN {grade_items} gi ON g.itemid = gi.id
            WHERE gi.courseid = :courseid
              AND gi.itemtype = 'mod'";  // Considera apenas atividades (itens de tipo 'mod')

    $records = $DB->get_records_sql($sql, ['courseid' => $courseid]);

    // Organiza notas por aluno
    $students = [];
    foreach ($records as $record) {
        $userid = $record->userid;
        if (!isset($students[$userid])) {
            $students[$userid] = (object)[
                'firstname' => $record->firstname,
                'lastname' => $record->lastname,
                'grades' => [],
            ];
        }
        if (!is_null($record->finalgrade)) {
            $students[$userid]->grades[] = $record->finalgrade;
        }
    }

    // Calcula a média das notas para cada aluno
    foreach ($students as &$student) {
        $total_grades = count($student->grades);
        $student->average_grade = $total_grades > 0 ? array_sum($student->grades) / $total_grades : 0;
    }

    return $students;
}
