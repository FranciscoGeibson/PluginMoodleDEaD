<?php
require_once __DIR__ . '/../../config.php';
require_once("$CFG->libdir/phpspreadsheet/vendor/autoload.php");


use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require_login();

$courseid = required_param('courseid', PARAM_INT);
$course = $DB->get_record("course", ['id' => $courseid]);

$context = context_course::instance($courseid);
require_capability('block/exportgradesigaa:view', $context);

// Variaveis globais
global $DB, $SESSION;


function check_conversion_server($server_url, $secret)
{
    // Verifica se o servidor de conversão está operacional
    $headers = [
        'X-Auth-Token: ' . $secret
    ];
    $ch = curl_init();
    curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
    curl_setopt($ch, CURLOPT_URL, $server_url . '?healthcheck=1');
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
    curl_setopt($ch, CURLOPT_TIMEOUT, 5);
    
    $response = curl_exec($ch);
    $http_code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    return $http_code === 200;
}


// Função para obter a carga horária do curso
function get_course_duration($courseid)
{
    global $DB;
    $field_name = 'edwcourseduration';
    $field = $DB->get_record('customfield_field', ['shortname' => $field_name]);

    if ($field) {
        $data = $DB->get_record('customfield_data', ['fieldid' => $field->id, 'instanceid' => $courseid]);
        if ($data) {
            return "{$data->value}h";
        }
    }
    return "(Duração não especificada)";
}

// Função para obter o ID da categoria de notas com base no nome da unidade
function get_category_id_by_name($category_prefix, $courseid)
{
    global $DB;

    // Remove espaços em branco no início e no final do prefixo
    $category_prefix = trim($category_prefix);

    // Busca todas as categorias do curso
    $categories = $DB->get_records('grade_categories', ['courseid' => $courseid]);

    foreach ($categories as $category) {
        // Remove espaços em branco no início e no final do nome da categoria
        $category_name = trim($category->fullname);

        // Verifica se o nome da categoria começa com o prefixo desejado (case-insensitive)
        if (stripos($category_name, $category_prefix) === 0) {
            // Verifica se o prefixo corresponde exatamente ao início do nome da categoria
            $prefix_length = strlen($category_prefix);
            $next_char = substr($category_name, $prefix_length, 1); // Pega o próximo caractere após o prefixo

            // Se o próximo caractere for um hífen, espaço ou o final da string, considera como correspondência válida
            if ($next_char === '-' || $next_char === ' ' || $next_char === '') {
                return $category->id; // Retorna o ID da categoria encontrada
            }
        }
    }
    return null; // Retorna null se não encontrar a categoria
}

// Função para obter o ID do item de nota com base no nome da atividade da Quarta Prova
function get_item_id_by_activity_name_Quarta_Prova($activity_prefix, $courseid)
{
    global $DB;

    // Remove espaços em branco no início e no final do prefixo
    $activity_prefix = trim($activity_prefix);

    // Busca todos os itens de nota do curso
    $items = $DB->get_records('grade_items', ['courseid' => $courseid, 'itemtype' => 'mod']);

    foreach ($items as $item) {
        // Remove espaços em branco no início e no final do nome da atividade
        $item_name = trim($item->itemname);

        // Verifica se o nome da atividade começa com o prefixo desejado (case-insensitive)
        if (stripos($item_name, $activity_prefix) === 0) {
            // Verifica se o prefixo corresponde exatamente ao início do nome da atividade
            $prefix_length = strlen($activity_prefix);
            $next_char = substr($item_name, $prefix_length, 1); // Pega o próximo caractere após o prefixo

            // Se o próximo caractere for um hífen, espaço ou o final da string, considera como correspondência válida
            if ($next_char === '-' || $next_char === ' ' || $next_char === '') {
                return $item->id; // Retorna o ID do item de nota encontrado
            }
        }
    }
    return null; // Retorna null se não encontrar o item
}

// Função para obter o ID do item de nota com base no ID da categoria
function get_item_id_by_category($category_id, $courseid)
{
    global $DB;

    // Busca o item de nota na tabela grade_items
    $item = $DB->get_record('grade_items', [
        'iteminstance' => $category_id, // ID da categoria
        'courseid' => $courseid, // ID do curso
        'itemtype' => 'category' // Tipo de item (categoria)
    ]);
    return $item ? $item->id : null; // Retorna o ID do item ou null se não encontrar
}

// Função para obter a nota de um aluno em uma categoria
function get_student_grade($studentId, $itemId)
{
    global $DB;

    // Busca a nota do aluno na tabela grade_grades
    $grade = $DB->get_record('grade_grades', [
        'itemid' => $itemId, // ID do item de nota
        'userid' => $studentId // ID do aluno
    ]);
    return $grade ? ($grade->finalgrade / 10) : 0; // Retorna a nota dividida por 10 ou 0 se não encontrar
}

// Função para obter a matrícula do aluno a partir de um campo personalizado
function get_student_enrollment($studentId)
{
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

function is_student_active_in_course($studentId, $courseId)
{
    global $DB;

    // Verifica se o usuário está ativo no sistema
    $user = $DB->get_record('user', [
        'id' => $studentId,
        'deleted' => 0,
        'suspended' => 0
    ]);

    if (!$user) {
        return false;
    }

    // Pega todos os métodos de inscrição ativos no curso
    $enrol_methxlsx = $DB->get_records('enrol', ['courseid' => $courseId, 'status' => 0]);

    foreach ($enrol_methxlsx as $enrol) {
        $enrolment = $DB->get_record('user_enrolments', [
            'enrolid' => $enrol->id,
            'userid' => $studentId,
            'status' => 0 // status 0 = inscrição ativa
        ]);

        if ($enrolment) {
            return true; // Está ativo
        }
    }
    return false; // Não está ativo
}


$course_duration = get_course_duration($courseid);

$spreadsheet = new Spreadsheet();
$activeWorksheet = $spreadsheet->getActiveSheet();

if (is_null($course_duration)) {
    echo $SESSION->gradeoverview_error = get_string('ch_nao_definida', 'block_exportgradesigaa'); // AJUSTAR

} else {

    //Envia para o servidor de conversão
    $conversion_server = 'http://localhost:8080/index.php';
    $secret = 'segredo123'; // mesma chave do servidor

    if (check_conversion_server($conversion_server, $secret)) {
        //Colocar a pesquisa da turma recursiva, T2...T3...T4...T5
        // Determinar a turma
        $course_shortname = htmlspecialchars($course->shortname);
        $course_shortname_cleaned = preg_replace('/\s*\(T[12]\)/', '', $course_shortname); // Remove "(T1)" ou "(T2)"

        $course_turma = strpos($course_shortname, '(T2)') !== false ? '02' : '01';

        $course_fullname = mb_strtoupper(htmlspecialchars($course->fullname));
        $course_polo = "Polo: MULTIPOLO";

        // Extrair o ano do course_shortname_cleaned
        // Assumimos que o formato é algo como "2024.2 - DSI0007 - Nome do Curso"
        preg_match('/^(\d{4}\.\d)\s*-\s*(.+)/', $course_shortname_cleaned, $matches);
        $course_year = isset($matches[1]) ? $matches[1] : 'AnoIndefinido';
        $course_shortname_cleaned_year = isset($matches[2]) ? str_replace(' - ', '', $matches[2]) : $course_shortname_cleaned;

        // Dados para o cabeçalho
        //$header_info = "$course_shortname_cleaned_year - $course_fullname ($course_duration) - Turma: $course_turma ($course_year) - $course_polo";

        $header_info = "$course_shortname_cleaned_year - $course_fullname ($course_duration) - Turma: $course_turma ($course_year) - $course_polo";

        // Dados introdutórios
        $intro_text = [
            ['', ''],
            ['', get_string('planilha_notas', 'block_exportgradesigaa')],
            ['', $header_info],
            ['', ''],
            ['', get_string('instrucao1', 'block_exportgradesigaa')],
            ['', get_string('instrucao2', 'block_exportgradesigaa')],
            ['', get_string('instrucao3', 'block_exportgradesigaa')],
            ['', get_string('instrucao4', 'block_exportgradesigaa')],
            ['', get_string('instrucao5', 'block_exportgradesigaa')],
            ['', get_string('instrucao6', 'block_exportgradesigaa')],
            ['', '']
        ];

        // Adiciona as informações introdutórias à planilha
        $row = 1;
        foreach ($intro_text as $line) {
            $col = 2;
            foreach ($line as $cell) {
                //$worksheet->write_string($row, $col++, $cell);
                $activeWorksheet->setCellValue([$col, $row], $cell);
            }
            $row++;
        }

        // Cabeçalhos da tabela
        $headers = ['', 'Matrícula', 'Nome', 'Unid. 1', 'Unid. 2', 'Unid. 3', 'Rec.', 'Resultado', 'Faltas', 'Sit.'];
        $col = 1;
        foreach ($headers as $header) {
            //$worksheet->write_string($row, $col++, $header);
            $activeWorksheet->setCellValue([$col++, $row], $header);
        }
        $row++;

        // Obter o ID do item de nota da "Quarta Prova"
        $quarta_prova_item_id = get_item_id_by_activity_name_Quarta_Prova('Quarta Prova', $courseid);

        // Dados dos alunos
        $students = get_role_users($DB->get_record("role", ['shortname' => 'student'])->id, $context);

        // Mapeamento da carga horária para o número de unidades
        $unidades_por_ch = [
            45 => 2, // Até 45 horas: 2 unidades
            PHP_INT_MAX => 3, // Acima de 45 horas: 3 unidades
        ];

        // Determinar o número de unidades com base na carga horária
        $num_unidades = 3; // Valor padrão (mais de 45 horas)
        foreach ($unidades_por_ch as $ch_maxima => $unidades) {
            if (intval($course_duration) <= $ch_maxima) {
                $num_unidades = $unidades;
                break;
            }
        }

        // Obter os IDs das categorias de notas com base no número de unidades
        $unit1_category_id = get_category_id_by_name('Unidade 1', $courseid);
        $unit2_category_id = get_category_id_by_name('Unidade 2', $courseid);
        $unit3_category_id = ($num_unidades == 3) ? get_category_id_by_name('Unidade 3', $courseid) : null;
        //$rec_category_id = get_category_id_by_name('Recuperação', $courseid); // Busca a categoria da recuperação

        // Passo 2: Obter o ID do item de nota
        $unit1_item_id = get_item_id_by_category($unit1_category_id, $courseid);
        $unit2_item_id = get_item_id_by_category($unit2_category_id, $courseid);
        $unit3_item_id = get_item_id_by_category($unit3_category_id, $courseid);
        //$rec_item_id = get_item_id_by_category($rec_category_id, $courseid); // Busca o item da recuperação

        //Pegar a REC, se existir
        //Sobre as Unidades, a carga horária <= 45 é de 2 Unidades. Maior que > 45, é de 3 Unidades.

        foreach ($students as $student) {

            if (!is_student_active_in_course($student->id, $courseid)) {
                continue; // Pula o aluno se estiver suspenso ou inativo
            }

            $fullname = mb_strtoupper($student->firstname . ' ' . $student->lastname);
            // Obter a matrícula do aluno a partir do campo personalizado
            $matricula = get_student_enrollment($student->id);

            $grade1 = get_student_grade($student->id, $unit1_item_id);
            $grade2 = get_student_grade($student->id, $unit2_item_id);
            $grade3 = get_student_grade($student->id, $unit3_item_id);
            //$rec_grade = get_student_grade($student->id, $rec_item_id); // Nota da recuperação
            $quarta_prova_grade = get_student_grade($student->id, $quarta_prova_item_id); // Nota da Quarta Prova

            // Adiciona os dados do aluno à planilha
            $data = [
                '',
                (string) $matricula, // Matrícula sempre como string
                $fullname,
                $grade1, //grade1
                $grade2, //grade2
                $grade3, //grade3
                $quarta_prova_grade, //Recuperação
                '', //$result_formula,
                0,  // Número de faltas padrão
                '', //$status_formula,
            ];

            // Escreve os dados do aluno na planilha
            foreach ($data as $col => $cell) {
                if ($col == 1) { // Matrícula sempre como string
                    //$activeWorksheet->setCellValue([$col + 1, $row], (string) $cell);
                    $activeWorksheet->getCell([$col + 1, $row])
                        ->setValueExplicit(
                            $cell,
                            \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING
                        );
                } elseif (is_numeric($cell)) {
                    $activeWorksheet->setCellValue([$col + 1, $row], $cell);
                } else {
                    $activeWorksheet->setCellValue([$col + 1, $row], $cell);
                }
            }

            $row++; // Avança para a próxima linha
        }

        // Após a última matrícula, adicionar uma única linha em branco
        $activeWorksheet->setCellValue([2, $row], ' ');

        // salva o xlsx temporariamente
        $xlsx_filename = clean_filename(format_string("notas_{$course->shortname}.xlsx"));

        $writer = new Xlsx($spreadsheet);
        $writer->save($xlsx_filename);

        // Primeiro verifica se o servidor está operacional
        $cfile = new CURLFile($xlsx_filename, 'application/vnd.oasis.opendocument.spreadsheet', basename($xlsx_filename));
        $dados = ['xlsx_file' => $cfile];

        $headers = [
            'X-Auth-Token: ' . $secret
        ];
        $ch = curl_init();
        curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);
        curl_setopt($ch, CURLOPT_URL, $conversion_server);
        curl_setopt($ch, CURLOPT_POST, 1);
        curl_setopt($ch, CURLOPT_POSTFIELDS, $dados);
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);

        $response = curl_exec($ch);
        $http_code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
        curl_close($ch);

        if ($http_code === 200) {
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment; filename="' . str_replace('.xlsx', '.xls', $xlsx_filename) . '"');

            // Define o cookie para indicar que o download foi iniciado
            setcookie('fileDownload', 'true', time() + 60, '/');

            echo $response;
            unlink($xlsx_filename);
            exit;
        } else {
            $OUTPUT->notification('message', get_string('servidor_nao_encontrado', 'block_exportgradesigaa'), 'error');

            // Fallback: disponibiliza o xlsx original se a conversão falhar
            readfile($xlsx_filename);
        }
    } else {
        // Log do erro para diagnóstico
        //error_log("Servidor de conversão indisponível");
        $SESSION->gradeoverview_error = get_string('servidor_nao_encontrado', 'block_exportgradesigaa');
        redirect(new moodle_url('/course/view.php', ['id' => $courseid]));

        readfile($xlsx_filename);
    }

    // Limpeza
    unlink($xlsx_filename);

    exit;
}
