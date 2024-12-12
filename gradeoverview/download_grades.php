<?php
require_once(__DIR__ . '/../../config.php');

require_login();

$courseid = required_param('courseid', PARAM_INT);
$course = $DB->get_record("course", ['id' => $courseid]);

$context = context_course::instance($courseid);
require_capability('block/gradeoverview:view', $context);

global $DB;

// Função para obter o ID da categoria de notas
function get_category_id_by_name($category_name, $courseid) {
    global $DB;
    $category = $DB->get_record('grade_categories', ['fullname' => $category_name, 'courseid' => $courseid]);
    return $category ? $category->id : null;
}

// Função para obter a nota de um aluno em uma categoria
function get_student_grade($studentId, $itemId) {
    global $DB;
    $grade = $DB->get_record('grade_grades', ['itemid' => $itemId, 'userid' => $studentId]);
    return $grade ? number_format($grade->finalgrade, 2, ',', '') : null;
}

// Função para obter a carga horária do curso usando $DB->get_record
function get_course_duration($courseid) {
    global $DB;
    $field_name = 'edwcoursedurationinhours'; // Nome breve do campo personalizado

    // Obter o campo personalizado pelo shortname
    $field = $DB->get_record('customfield_field', ['shortname' => $field_name]);

    if ($field) {
        // Obter o valor do campo personalizado para o curso
        $data = $DB->get_record('customfield_data', ['fieldid' => $field->id, 'instanceid' => $courseid]);
        if ($data) {
            return "{$data->value}h";
        }
    }
    return "(Duração não especificada)";
}

// Obter a carga horária do curso
$course_duration = get_course_duration($courseid);

// if de verificação da turma
$course_shortname = htmlspecialchars($course->shortname);
if (strpos($course_shortname, '(T2)') !== false) {
    $course_turma = '02';
} else {
    $course_turma = '01';
}

// Cabeçalhos do arquivo CSV
header("Content-Type: text/csv; charset=UTF-8");
header("Content-Disposition: attachment; filename=\"notas_alunos.csv\"");
header("Cache-Control: no-cache, no-store, must-revalidate");
header("Pragma: no-cache");
header("Expires: 0");

// Imprimir o BOM para garantir a codificação UTF-8
echo "\xEF\xBB\xBF";

// Abrir saída para escrever o CSV
$output = fopen('php://output', 'w');

// Preparar dados do arquivo
$course_shortname = htmlspecialchars($course->shortname);
$course_fullname = htmlspecialchars($course->fullname);
$course_polo = "Polo: MULTIPOLO"; // fixo

// formatação sem usar sprintf
$shortname_parts = explode('-', $course_shortname, 2);
$year = trim($shortname_parts[0]);
$code = isset($shortname_parts[1]) ? trim($shortname_parts[1]) : '';

$formatted_course_info = $code . ' - ' . $course_fullname . ' ' . $course_duration . ' - Turma: ' . $course_turma . ' (' . $year . ') - ' . $course_polo;

// Escrever no CSV
$intro_text = [
    [''],
    ['', 'PLANILHA DE NOTAS'],
    ['', $formatted_course_info],
    [''],
    ['', 'Digite as notas das unidades utilizando vírgula para separar a casa decimal.'],
    ['', 'O campo faltas deve ser preenchido com o número de faltas do aluno durante o período letivo.'],
    ['', 'A situação do aluno em relação à assiduidade é calculada apenas levando em consideração a carga horária da disciplina.'],
    ['', 'Devido a isso a situação pode mudar durante a importação da planilha.'],
    ['', 'As notas das unidades não vão para o histórico do aluno, no entanto, aparecem em seu portal.'],
    ['', 'Altere somente as células em amarelo.'],
];

foreach ($intro_text as $line) {
    fputcsv($output, $line, ';');
}

// Adicionar uma linha em branco para separação
fputcsv($output, [''], ';');

// Escrever os cabeçalhos da tabela
$headers = ['', 'Matrícula', 'Nome', 'Unid. 1', 'Unid. 2', 'Unid. 3', 'Rec.', 'Resultado', 'Faltas', 'Situação'];
fputcsv($output, $headers, ';');

// Adicionar os dados dos alunos
$students = get_role_users($DB->get_record("role", ['shortname' => 'student'])->id, $context);
$unit1_category_id = get_category_id_by_name('Unidade 1', $courseid);
$unit2_category_id = get_category_id_by_name('Unidade 2', $courseid);
$unit3_category_id = get_category_id_by_name('Unidade 3', $courseid);

$absences - 0;
$row = 13;
foreach ($students as $student) {
    $fullname = $student->firstname . ' ' . $student->lastname;
    $username = $student->username;
    $grade1 = get_student_grade($student->id, $unit1_category_id);
    $grade2 = get_student_grade($student->id, $unit2_category_id);
    $grade3 = get_student_grade($student->id, $unit3_category_id);

    // Formatando notas
    $grade1 = $grade1 !== null ? str_replace('.', ',', $grade1) : '';
    $grade2 = $grade2 !== null ? str_replace('.', ',', $grade2) : '';
    $grade3 = $grade3 !== null ? str_replace('.', ',', $grade3) : '';

    // Fórmulas
    $result_formula = <<<EOD
    =SE(OU(D$row="-"; D$row=""; E$row="-"; E$row=""; F$row="-"; F$row=""); "-"; SE(OU(G$row=""; G$row<0; G$row="-"); (ARRED((((D$row*4*10)+(E$row*5*10)+(F$row*6*10))/150)*10; 0)/10); (ARRED(((SE(MÍNIMO(D$row; E$row; F$row)=D$row; (D$row*4*10)+(E$row*5*10)+(F$row*6*10)-(D$row*6*10)+(G$row*6*10); SE(MÍNIMO(D$row; E$row; F$row)=E$row; (D$row*4*10)+(E$row*5*10)+(F$row*6*10)-(E$row*6*10)+(G$row*6*10); SE(MÍNIMO(D$row; E$row; F$row)=F$row; (D$row*4*10)+(E$row*5*10)+(F$row*6*10)-(F$row*6*10)+(G$row*6*10); ))))/150)*10; 0)/10)))
    EOD;

    $status_formula = <<<EOD
    =SE(OU(H$row="-";H$row="");"-";SE(I$row>15;SE(H$row>=6;SE((OU(D$row<7;E$row<7;F$row<7));"RENF";"REPF");"REMF");SE(H$row>=7;"APR";SE(H$row>=6;SE(E((OU(D$row<7;E$row<7;F$row<7));(OU(G$row="-";G$row="")));"REC";SE(OU(G$row="-";G$row="");SE(H$row>=7;"APR";"APRN");SE(G$row>=7;"APR";"REPN")));SE(H$row>=7;SE(OU(G$row="-";G$row="");"REC";"REP");"REP")))))
    EOD;

    $data = [
        '',
        $username,
        $fullname,
        $grade1,
        $grade2,
        $grade3,
        '',
        $result_formula,
        $absences,
        $status_formula,
    ];

    fputcsv($output, $data, ';');
    $row++;
}

// Fechar o CSV
fclose($output);
exit;
