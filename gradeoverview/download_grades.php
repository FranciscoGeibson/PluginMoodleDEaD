<?php
// Inclui a configuração do Moodle
require_once(__DIR__ . '/../../config.php'); 
require_login(); // Certifica-se de que o usuário está logado

$courseid = required_param('courseid', PARAM_INT); // Obtém o ID do curso

// Obtém o curso com base no ID
$course = $DB->get_record("course", ['id' => $courseid]);

// Certifica-se de que o usuário tem permissão para acessar o curso
$context = context_course::instance($courseid);
require_capability('moodle/grade:viewall', $context);

global $DB;

// Função para obter o ID da categoria com base no nome da categoria
function get_category_id_by_name($category_name, $courseid) {
    global $DB;
    $category = $DB->get_record('grade_categories', ['fullname' => $category_name, 'courseid' => $courseid]);
    return $category ? $category->id : null;
}

// Define os IDs das categorias de forma dinâmica
$unit1_category_id = get_category_id_by_name('Unidade 1', $courseid);
$unit2_category_id = get_category_id_by_name('Unidade 2', $courseid);
$unit3_category_id = get_category_id_by_name('Unidade 3', $courseid);

function get_student_grade($studentId, $itemId) {
    global $DB;
    $grade = $DB->get_record('grade_grades', ['itemid' => $itemId, 'userid' => $studentId]);
    return $grade ? number_format($grade->finalgrade, 2) : null;
}

try {
    // Configura o cabeçalho do arquivo para download
    header('Content-Type: text/csv');
    header('Content-Disposition: attachment;filename="notas_alunos.csv"');

    // Cria o arquivo CSV
    $output = fopen('php://output', 'w');

    // Adiciona o título "Notas"
    fputcsv($output, ['Notas']);
    // Cabeçalho com ID do usuário, Nome, e Notas
    fputcsv($output, ['ID do Usuário', 'Nome Completo', 'Unidade 1', 'Unidade 2', 'Unidade 3', 'Média Final']);

    // Obtém o papel "student" para identificar os alunos do curso
    $role = $DB->get_record("role", ['shortname' => 'student']);
    $students = get_role_users($role->id, $context);

    // Carrega os itens de nota para o curso
    $grade_items = $DB->get_records('grade_items', ['courseid' => $courseid]);

    foreach ($students as $student) {
        $fullname = $student->firstname . ' ' . $student->lastname;

        // Inicializa notas
        $grades = [
            $unit1_category_id => null,
            $unit2_category_id => null,
            $unit3_category_id => null,
        ];
        $total_grades = 0;
        $grade_count = 0;

        // Itera sobre cada item de nota do curso
        foreach ($grade_items as $item) {
            $grade = get_student_grade($student->id, $item->id);
            if ($grade !== null) {
                // Atribui a nota conforme a categoria do item
                if ($item->categoryid == $unit1_category_id) {
                    $grades[$unit1_category_id] = $grade;
                } elseif ($item->categoryid == $unit2_category_id) {
                    $grades[$unit2_category_id] = $grade;
                } elseif ($item->categoryid == $unit3_category_id) {
                    $grades[$unit3_category_id] = $grade;
                }
                // Atualiza o total e o contador
                $total_grades += floatval($grade); // Converte a nota para float
                $grade_count++; // Incrementa o contador de notas
            }
        }

        // Calcula a média final
        $average = $grade_count > 0 ? number_format($total_grades / $grade_count, 2) : '';

        // Adiciona as informações do aluno ao CSV
        fputcsv($output, [
            $student->id,
            $fullname,
            $grades[$unit1_category_id],
            $grades[$unit2_category_id],
            $grades[$unit3_category_id],
            $average
        ]);
    }

    fclose($output);
    exit;

} catch (Exception $e) {
    // Se ocorrer um erro, exibe uma mensagem de erro amigável
    error_log("Erro ao gerar o CSV: " . $e->getMessage());
    echo "Ocorreu um erro ao gerar o CSV. Por favor, tente novamente mais tarde.";
    exit;
}
