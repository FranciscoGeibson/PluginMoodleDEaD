<?php
require_once __DIR__ . '/../../config.php';
//require_once __DIR__ . '/../../vendor/autoload.php';
require_once 'C:\Users\Windows\Documents\MoodleWindowsInstaller-latest-405\server\vendor\autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

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

// Função para obter a matrícula do aluno a partir de um campo personalizado
function get_student_enrollment($studentId) {
    global $DB;

    // Obter o campo de perfil pelo nome breve 'matricula'
    $field = $DB->get_record('user_info_field', ['shortname' => 'matricula']);

    if ($field) {
        // Procurar o valor da matrícula do aluno
        $data = $DB->get_record('user_info_data', ['fieldid' => $field->id, 'userid' => $studentId]);
        if ($data) {
            // Retornar a matrícula sem os 3 primeiros caracteres
            return substr($data->data, 3); // Remove os 3 primeiros caracteres (DLL)
        }
    }

    return null; // Retorna null se não encontrar o campo ou o dado
}


// Função para obter a carga horária do curso
function get_course_duration($courseid) {
    global $DB;
    $field_name = 'edwcoursedurationinhours';
    $field = $DB->get_record('customfield_field', ['shortname' => $field_name]);

    if ($field) {
        $data = $DB->get_record('customfield_data', ['fieldid' => $field->id, 'instanceid' => $courseid]);
        if ($data) {
            return "{$data->value}h";
        }
    }
    return "(Duração não especificada)";
}

$course_duration = get_course_duration($courseid);

// Determinar a turma
$course_shortname = htmlspecialchars($course->shortname);
$course_shortname_cleaned = preg_replace('/\s*\(T[12]\)/', '', $course_shortname); // Remove "(T1)" ou "(T2)"

$course_turma = strpos($course_shortname, '(T2)') !== false ? '02' : '01';

$course_fullname = htmlspecialchars($course->fullname);
$course_polo = "Polo: MULTIPOLO";

// Extrair o ano do course_shortname_cleaned
// Assumimos que o formato é algo como "2024.2 - DSI0007 - Nome do Curso"
preg_match('/^(\d{4}\.\d)\s*-\s*(.+)/', $course_shortname_cleaned, $matches);
$course_year = isset($matches[1]) ? $matches[1] : 'AnoIndefinido';
$course_shortname_cleaned_year = isset($matches[2]) ? str_replace(' - ', '', $matches[2]) : $course_shortname_cleaned;



// Dados para o cabeçalho
$header_info = "$course_shortname_cleaned_year - $course_fullname ($course_duration) - Turma: $course_turma ($course_year) - $course_polo";

// Criar a planilha
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Ajustar a largura da coluna A
//$sheet->getColumnDimension('A')->setWidth(2);

// Ajustar a largura da coluna K
//$sheet->getColumnDimension('K')->setWidth(2);

// Ajustar a largura da coluna B
$sheet->getColumnDimension('B')->setWidth(14);

// Ajustar a largura da coluna C
$sheet->getColumnDimension('C')->setWidth(54);

// Ajustar a largura da coluna D
$sheet->getColumnDimension('D')->setWidth(5.22);

// Ajustar a largura da coluna E
$sheet->getColumnDimension('E')->setWidth(5.22);

// Ajustar a largura da coluna F
$sheet->getColumnDimension('F')->setWidth(5.22);

// Ajustar a largura da coluna G
$sheet->getColumnDimension('G')->setWidth(5.22);

// Ajustar a largura da coluna H
$sheet->getColumnDimension('H')->setWidth(9);

// Ajustar a largura da coluna I
$sheet->getColumnDimension('I')->setWidth(5.22);

// Ajustar a largura da coluna J
$sheet->getColumnDimension('J')->setWidth(5.22);


// Dados introdutórios
$intro_text = [
    [''],
    ['PLANILHA DE NOTAS'],
    [$header_info],
    [''],
    ['Digite as notas das unidades utilizando vírgula para separar a casa decimal.'],
    ['O campo faltas deve ser preenchido com o número de faltas do aluno durante o período letivo.'],
    ['A situação do aluno em relação à assiduidade é calculada apenas levando em consideração a carga horária da disciplina.'],
    ['Devido a isso a situação pode mudar durante a importação da planilha.'],
    ['As notas das unidades não vão para o histórico do aluno, no entanto, aparecem em seu portal.'],
    ['Altere somente as células em amarelo.'],
    [''],
];

// Escrever os dados introdutórios na planilha
$row = 1;
$start_column = 'B';
$end_column = 'J'; // Define o alcance horizontal da tabela
foreach ($intro_text as $line) {
    $sheet->fromArray($line, null, "{$start_column}{$row}");


    // Mesclar as células da linha inteira
    $sheet->mergeCells("{$start_column}{$row}:{$end_column}{$row}");
    

    // Aplicar fonte Arial, tamanho 10, a toda a planilha
    $spreadsheet->getDefaultStyle()->getFont()->setName('Arial')->setSize(10);

    // Formatar como negrito
    $sheet->getStyle("{$start_column}{$row}:{$end_column}{$row}")->getFont()->setBold(true);
    $row++;
}


// Mesclar as células da linha 4, de B a J
//$sheet->mergeCells('B4:J4');

// Centralizar o texto na célula mesclada
//$sheet->getStyle('B4')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

// Aplicar formatação em negrito à célula mesclada
//$sheet->getStyle('B4')->getFont()->setBold(true);

// Preencher o conteúdo da linha 4
//$sheet->setCellValue('B4', $header_info);




// Cabeçalhos da tabela
$headers = ['Matrícula', 'Nome', 'Unid. 1', 'Unid. 2', 'Unid. 3', 'Rec.', 'Resultado', 'Faltas', 'Sit.'];
$sheet->fromArray($headers, null, "B{$row}");
$sheet->getStyle("B{$row}:J{$row}")->getFont()->setBold(true); // Matrícula
$sheet->getStyle("B{$row}:J{$row}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER); // Centralizar


// Aplicar fonte Arial, tamanho 10, a toda a planilha
$spreadsheet->getDefaultStyle()->getFont()->setName('Arial')->setSize(10);



$row++;

// Dados dos alunos
$students = get_role_users($DB->get_record("role", ['shortname' => 'student'])->id, $context);
$unit1_category_id = get_category_id_by_name('Unidade 1', $courseid);
$unit2_category_id = get_category_id_by_name('Unidade 2', $courseid);
$unit3_category_id = get_category_id_by_name('Unidade 3', $courseid);

foreach ($students as $student) {
    $fullname = strtoupper($student->firstname . ' ' . $student->lastname);
    //$username = $student->username;


    // Obter a matrícula do aluno a partir do campo personalizado
    $matricula = get_student_enrollment($student->id);

    // Caso a matrícula seja nula, exibir uma mensagem de erro ou usar um valor padrão
    if (is_null($matricula)) {
        $matricula = 'Matrícula não encontrada';
    }

    $grade1 = get_student_grade($student->id, $unit1_category_id);
    $grade2 = get_student_grade($student->id, $unit2_category_id);
    $grade3 = get_student_grade($student->id, $unit3_category_id);

    $grade1 = $grade1 !== null ? str_replace('.', ',', $grade1) : '';
    $grade2 = $grade2 !== null ? str_replace('.', ',', $grade2) : '';
    $grade3 = $grade3 !== null ? str_replace('.', ',', $grade3) : '';

    $result_formula = <<<EOD
    =SE(OU(D$row="-"; D$row=""; E$row="-"; E$row=""; F$row="-"; F$row=""); "-"; SE(OU(G$row=""; G$row<0; G$row="-"); (ARRED((((D$row*4*10)+(E$row*5*10)+(F$row*6*10))/150)*10; 0)/10); (ARRED(((SE(MÍNIMO(D$row; E$row; F$row)=D$row; (D$row*4*10)+(E$row*5*10)+(F$row*6*10)-(D$row*6*10)+(G$row*6*10); SE(MÍNIMO(D$row; E$row; F$row)=E$row; (D$row*4*10)+(E$row*5*10)+(F$row*6*10)-(E$row*6*10)+(G$row*6*10); SE(MÍNIMO(D$row; E$row; F$row)=F$row; (D$row*4*10)+(E$row*5*10)+(F$row*6*10)-(F$row*6*10)+(G$row*6*10); ))))/150)*10; 0)/10)))
    EOD;


    //$status_formula = "IF(OR(H{$row}="-", H{$row}=""), "-", IF(I{$row}>15, IF(H{$row}>=6, IF(OR(D{$row}<7, E{$row}<7, F{$row}<7), "RENF", "REPF"), "REMF"), IF(H{$row}>=7, "APR", IF(H{$row}>=6, IF(AND(OR(D{$row}<7, E{$row}<7, F{$row}<7), OR(G{$row}="-", G{$row}="")), "REC", IF(OR(G{$row}="-", G{$row}=""), IF(H{$row}>=7, "APR", "APRN"), IF(G{$row}>=7, "APR", "REPN"))), IF(H{$row}>=7, IF(OR(G{$row}="-", G{$row}=""), "REC", "REP"), "REP")))))";

    //$sheet->setCellValueExplicit("H{$row}", $result_formula, DataType::TYPE_FORMULA);
    //$sheet->setCellValueExplicit("J{$row}", $status_formula, DataType::TYPE_FORMULA);
    //$result_formula = "=SE(OU(D{$row}=\"-\"; D{$row}=\"\"; E{$row}=\"-\"; E{$row}=\"\"; F{$row}=\"-\"; F{$row}=\"\"); \"-\"; SE(OU(G{$row}=\"\"; G{$row}<0; G{$row}=\"-\"; G{$row}=\"\"), ARRED((((D{$row}*4*10)+(E{$row}*5*10)+(F{$row}*6*10))/150)*10; 0)/10; ARRED((((D{$row}*4*10)+(E{$row}*5*10)+(F{$row}*6*10))/150)*10; 0)/10))";

    // Insira diretamente na célula sem tratamento especial
    //$sheet->setCellValue('H13', '=D13*E13');
    //$sheet->setCellValue("H{$row}", "=SE(OU(D{$row}=\"-\"; D{$row}=\"\"; E{$row}=\"-\"; E{$row}=\"\"; F{$row}=\"-\"; F{$row}=\"\"); \"-\"; SE(OU(G{$row}=\"\"; G{$row}<0; G{$row}=\"-\"; G{$row}=\"\"), ARRED((((D{$row}*4*10)+(E{$row}*5*10)+(F{$row}*6*10))/150)*10; 0)/10; ARRED((((D{$row}*4*10)+(E{$row}*5*10)+(F{$row}*6*10))/150)*10; 0)/10))");

    // Corrigir a fórmula para o formato em inglês
    //$result_formula = "=IF(OR(D$row="-", D$row="", E$row="-", E$row="", F$row="-", F$row=""), "-", IF(OR(G$row="", G$row<0, G$row="-"), ROUND((((D$row*4*10)+(E$row*5*10)+(F$row*6*10))/150)*10, 0)/10, ROUND(((IF(MIN(D$row, E$row, F$row)=D$row, (D$row*4*10)+(E$row*5*10)+(F$row*6*10)-(D$row*6*10)+(G$row*6*10), IF(MIN(D$row, E$row, F$row)=E$row, (D$row*4*10)+(E$row*5*10)+(F$row*6*10)-(E$row*6*10)+(G$row*6*10), IF(MIN(D$row, E$row, F$row)=F$row, (D$row*4*10)+(E$row*5*10)+(F$row*6*10)-(F$row*6*10)+(G$row*6*10)))))/150)*10, 0)/10)))"

    $result_formula = "=IF(OR(D{$row}=\"-\", D{$row}=\"\", E{$row}=\"-\", E{$row}=\"\", F{$row}=\"-\", F{$row}=\"\"), \"-\", IF(OR(G{$row}=\"\", G{$row}<0, G{$row}=\"-\", G{$row}=\"\"), ROUND((((D{$row}*4*10)+(E{$row}*5*10)+(F{$row}*6*10))/150)*10, 0)/10, ROUND((((D{$row}*4*10)+(E{$row}*5*10)+(F{$row}*6*10))/150)*10, 0)/10))";
    $status_formula = "=IF(OR(H{$row}=\"-\", H{$row}=\"\"), \"-\", IF(I{$row}>15, IF(H{$row}>=6, IF(OR(D{$row}<7, E{$row}<7, F{$row}<7), \"RENF\", \"REPF\"), \"REMF\"), IF(H{$row}>=7, \"APR\", IF(H{$row}>=6, IF(AND(OR(D{$row}<7, E{$row}<7, F{$row}<7), OR(G{$row}=\"-\", G{$row}=\"\")), \"REC\", IF(OR(G{$row}=\"-\", G{$row}=\"\"), IF(H{$row}>=7, \"APR\", \"APRN\"), IF(G{$row}>=7, \"APR\", \"REPN\"))), IF(H{$row}>=7, IF(OR(G{$row}=\"-\", G{$row}=\"\"), \"REC\", \"REP\"), \"REP\")))))";



    // Inserindo a fórmula na célula, por exemplo, em "H{$row}"
    $sheet->setCellValue("H{$row}", $result_formula);
    $sheet->setCellValue("J{$row}", $status_formula);


    // Remove a aspas simples (caso tenha sido adicionada)
    //$sheet->getCell("H{$row}")->setValueExplicit($formula, DataType::TYPE_FORMULA);


    // Adicione os outros dados normalmente
    $data = [
        $matricula,
        $fullname,
        $grade1,
        $grade2,
        $grade3,
        '',
        '',//$result_formula,
        '',
        //'', // Deixe a fórmula de Situação em J{$row}, já configurada acima
    ];

    // Aplicar formatação em negrito às colunas de Matrícula e Nome
    $sheet->getStyle("B{$row}")->getFont()->setBold(true); // Matrícula
    $sheet->getStyle("C{$row}")->getFont()->setBold(true); // Nome
    $sheet->getStyle("D{$row}")->getFont()->setBold(true); // Unid. 1
    $sheet->getStyle("E{$row}")->getFont()->setBold(true); // Unid. 2
    $sheet->getStyle("F{$row}")->getFont()->setBold(true); // Unid. 3
    $sheet->getStyle("G{$row}")->getFont()->setBold(true); // Resultado
    $sheet->getStyle("H{$row}")->getFont()->setBold(true); // Faltas
    $sheet->getStyle("I{$row}")->getFont()->setBold(true); // Situação

    // Alinhar a coluna B à esquerda
    $sheet->getStyle("B{$row}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

    $sheet->fromArray($data, null, "B{$row}");
    $row++;
}


// Criar o nome do arquivo
$filename = "notas_{$course_shortname_cleaned_year}_T{$course_turma}_{$course_year}.xls";

// Baixar o arquivo
$writer = new Xls($spreadsheet); // Alterado para Xls
header('Content-Type: application/vnd.ms-excel'); // MIME type para arquivos XLS
header("Content-Disposition: attachment; filename=\"{$filename}\"");// Alterado para .xls notas_DSI0007_T01_20242.xls
$writer->save('php://output');
exit;
