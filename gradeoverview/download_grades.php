<?php
require_once(__DIR__ . '/../../config.php'); 
require_login();

$courseid = required_param('courseid', PARAM_INT);
$course = $DB->get_record("course", ['id' => $courseid]);

$context = context_course::instance($courseid);
require_capability('moodle/grade:viewall', $context);

global $DB;

// Função para obter o ID da categoria de notas
function get_category_id_by_name($category_name, $courseid) {
    global $DB;
    $category = $DB->get_record('grade_categories', ['fullname' => $category_name, 'courseid' => $courseid]);
    return $category ? $category->id : null;
}

$unit1_category_id = get_category_id_by_name('Unidade 1', $courseid);
$unit2_category_id = get_category_id_by_name('Unidade 2', $courseid);
$unit3_category_id = get_category_id_by_name('Unidade 3', $courseid);

// Função para obter a nota de um aluno em uma categoria
function get_student_grade($studentId, $itemId) {
    global $DB;
    $grade = $DB->get_record('grade_grades', ['itemid' => $itemId, 'userid' => $studentId]);
    return $grade ? number_format($grade->finalgrade, 2) : null;
}

// Função para obter as faltas de um aluno
function get_student_absences($studentId) {
    global $DB;
    // Aqui é assumido que as faltas são armazenadas em um campo personalizado 'faltas'
    $profile = $DB->get_record('user_info_data', ['userid' => $studentId, 'fieldid' => 1]); // Ajuste o campo ID conforme necessário
    return $profile ? $profile->data : 0; // Retorna o número de faltas ou 0 se não encontrado
}

// Lógica de situação do aluno
function get_student_status($average, $absences, $allowed_absences) {
    if ($absences > $allowed_absences) {
        return "REP";
    }
    if ($average < 7) {
        return "REP";
    }
    return "APR";
}

try {
    // Configura os cabeçalhos para o navegador interpretar como um arquivo Excel
    header("Content-Type: application/vnd.ms-excel");
    header("Content-Disposition: attachment; filename=notas_alunos.xls");
    header("Cache-Control: max-age=0");

    // Defina o limite de faltas permitido (exemplo: 25%)
    $course_hours = 150; // Substitua com a duração real do curso
    $allowed_absences = $course_hours * 0.25;

    // Cabeçalho de informações
    echo '<table border="1">';
    echo '<tr><td colspan="8"><strong>PLANILHA DE NOTAS</strong></td></tr>';
    echo '<tr><td colspan="8">' . htmlspecialchars($course->fullname) . ' </td></tr>';  // Nome do curso
    echo '<tr><td colspan="8">Digite as notas das unidades utilizando vírgula para separar a casa decimal.</td></tr>';
    echo '<tr><td colspan="8">O campo faltas deve ser preenchido com o número de faltas do aluno durante o período letivo.</td></tr>';
    echo '<tr><td colspan="8">A situação do aluno em relação a assiduidade é calculada apenas levando em consideração a carga horária da disciplina.</td></tr>';
    echo '<tr><td colspan="8">Devido a isso a situação pode mudar durante a importação da planilha.</td></tr>';
    echo '<tr><td colspan="8">As notas das unidades não vão para o histórico do aluno, no entanto, aparecem em seu portal.</td></tr>';
    echo '<tr><td colspan="8">Altere somente as células em amarelo.</td></tr>';
    echo '<tr><td colspan="8"></td></tr>'; // Linha em branco entre cabeçalho e tabela de notas

    // Tabela de notas
    echo '<tr><th>Matrícula</th><th>Nome</th><th>Unid. 1</th><th>Unid. 2</th><th>Unid. 3</th><th>Resultado</th><th>Faltas</th><th>Situação</th></tr>';

    // Obtém os usuários com o papel de aluno
    $role = $DB->get_record("role", ['shortname' => 'student']);
    $students = get_role_users($role->id, $context);

    foreach ($students as $student) {
        $fullname = $student->firstname . ' ' . $student->lastname;
        $username = $student->username; // Pega o username do estudante (matrícula)

        $grade1 = get_student_grade($student->id, $unit1_category_id);
        $grade2 = get_student_grade($student->id, $unit2_category_id);
        $grade3 = get_student_grade($student->id, $unit3_category_id);

        $grades = [$grade1, $grade2, $grade3];
        $total_grades = 0;
        $grade_count = 0;

        foreach ($grades as $grade) {
            if ($grade !== null) {
                $total_grades += floatval($grade);
                $grade_count++;
            }
        }

        $average = $grade_count > 0 ? number_format($total_grades / $grade_count, 2) : '';

        // Obtém o número de faltas
        $absences = get_student_absences($student->id);

        // Define a situação do aluno
        $status = get_student_status(floatval($average), intval($absences), $allowed_absences);

        // Substitui ponto por vírgula nos valores numéricos
        $grade1 = $grade1 !== null ? str_replace('.', ',', $grade1) : '';
        $grade2 = $grade2 !== null ? str_replace('.', ',', $grade2) : '';
        $grade3 = $grade3 !== null ? str_replace('.', ',', $grade3) : '';
        $average = $average !== '' ? str_replace('.', ',', $average) : '';

        echo '<tr>';
        echo '<td>' . htmlspecialchars($username) . '</td>'; // Exibe o username no lugar do ID
        echo '<td>' . htmlspecialchars($fullname) . '</td>';
        echo '<td>' . $grade1 . '</td>';
        echo '<td>' . $grade2 . '</td>';
        echo '<td>' . $grade3 . '</td>';
        echo '<td>' . $average . '</td>';
        echo '<td>' . $absences . '</td>'; // Coluna de faltas
        echo '<td>' . htmlspecialchars($status) . '</td>'; // Coluna de situação
        echo '</tr>';
    }

    echo '</table>';
    exit;

} catch (Exception $e) {
    error_log("Erro ao gerar o arquivo XLS: " . $e->getMessage());
    echo "Ocorreu um erro ao gerar o arquivo XLS. Por favor, tente novamente mais tarde.";
    exit;
}
