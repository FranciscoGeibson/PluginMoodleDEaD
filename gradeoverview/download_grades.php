<?php
require_once __DIR__ . '/../../config.php';
require_once($CFG->libdir . '/odslib.class.php');

require_login();


$courseid = required_param('courseid', PARAM_INT);
$course = $DB->get_record("course", ['id' => $courseid]);


$context = context_course::instance($courseid);
require_capability('block/gradeoverview:view', $context);

global $DB;
// Função para obter a carga horária do curso
function get_course_duration($courseid) {
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
function get_category_id_by_name($category_prefix, $courseid) {
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
function get_item_id_by_activity_name_Quarta_Prova($activity_prefix, $courseid) {
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
function get_item_id_by_category($category_id, $courseid) {
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
function get_student_grade($studentId, $itemId) {
    global $DB;

    // Busca a nota do aluno na tabela grade_grades
    $grade = $DB->get_record('grade_grades', [
        'itemid' => $itemId, // ID do item de nota
        'userid' => $studentId // ID do aluno
    ]);

    return $grade ? ($grade->finalgrade / 10) : null; // Retorna a nota dividida por 10 ou null se não encontrar
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

function is_student_active_in_course($studentId, $courseId) {
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
    $enrol_methods = $DB->get_records('enrol', ['courseid' => $courseId, 'status' => 0]);

    foreach ($enrol_methods as $enrol) {
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


// Criar objeto ODS
$filename = "notas_{$course->shortname}.ods";
$workbook = new MoodleODSWorkbook(format_string($course->fullname, true));
$workbook->send($filename);
$worksheet = $workbook->add_worksheet("Notas");

$course_duration = get_course_duration($courseid);


if (is_null($course_duration)) {
    echo "Erro: A carga horária (CH) da turma não foi definida.\n";
} else {
        //Colocar a pesquisa da turma recursiva, T2...T3...T4...T5
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
        //$header_info = "$course_shortname_cleaned_year - $course_fullname ($course_duration) - Turma: $course_turma ($course_year) - $course_polo";

        $header_info = "$course_shortname_cleaned_year - $course_fullname ($course_duration) - Turma: $course_turma ($course_year)";

        // Dados introdutórios
        $intro_text = [
            ['', ''],  // Linha em branco
            ['', 'PLANILHA DE NOTAS'],  // Adicionando um espaço antes do título
            ['', $header_info],  // Mantendo o cabeçalho com deslocamento
            ['', ''],  // Linha em branco
            ['', 'Digite as notas das unidades utilizando vírgula para separar a casa decimal.'],
            ['', 'O campo faltas deve ser preenchido com o número de faltas do aluno durante o período letivo.'],
            ['', 'A situação do aluno em relação a assiduidade é calculada apenas levando em consideração a carga horária da disciplina.'],
            ['', 'Devido a isso a situação pode mudar durante a importação da planilha.'],
            ['', 'As notas das unidades não vão para o histórico do aluno, mas aparecem no portal do aluno.'],
            ['', 'Altere somente as células em amarelo.'],
            ['', ''],  // Linha em branco
        ];


        // Adiciona as informações introdutórias à planilha
        $row = 0;
        foreach ($intro_text as $line) {
            $col = 0;
            foreach ($line as $cell) {
                $worksheet->write_string($row, $col++, $cell);
            }
            $row++;
        }

        // Cabeçalhos da tabela
        $headers = ['', 'Matrícula', 'Nome', 'Unid. 1', 'Unid. 2', 'Unid. 3', 'Rec.', 'Resultado', 'Faltas', 'Sit.'];
        $col = 0;
        foreach ($headers as $header) {
            $worksheet->write_string($row, $col++, $header);
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

        
        // Obtém a lista de alunos matriculados no curso
        //$students = get_role_users($DB->get_record("role", ['shortname' => 'student'])->id, $context);
        // Contar quantos alunos existem antes de escrever na planilha
        $totalStudents = count($students);
        $studentCount = 0; // Contador para acompanhar a matrícula atual

        foreach ($students as $student) {

            if (!is_student_active_in_course($student->id, $courseid)) {
                continue; // Pula o aluno se estiver suspenso ou inativo
            }

            $fullname = strtoupper($student->firstname . ' ' . $student->lastname);
            // Obter a matrícula do aluno a partir do campo personalizado
            $matricula = get_student_enrollment($student->id);

            /*
            // Caso a matrícula seja nula, exibir uma mensagem de erro ou usar um valor padrão
            if (is_null($matricula)) {
                $matricula = 'Matrícula não encontrada'; // Ou qualquer outro valor padrão
            } else {
                $matricula = (string) $matricula; // Converte a matrícula para string
            }
            */

            $grade1 = get_student_grade($student->id, $unit1_item_id);
            $grade2 = get_student_grade($student->id, $unit2_item_id);
            $grade3 = get_student_grade($student->id, $unit3_item_id);
            //$rec_grade = get_student_grade($student->id, $rec_item_id); // Nota da recuperação
            $quarta_prova_grade = get_student_grade($student->id, $quarta_prova_item_id); // Nota da Quarta Prova

            /*
            if (is_null($grade1)) {
                echo "Erro: Nota da Unidade 1 inválida ou não encontrada para o aluno {$student->firstname} {$student->lastname}.\n";
            } else {
                //echo "Nota da Unidade 1 para o aluno {$student->firstname} {$student->lastname}: {$grade1}\n";
            }

            if (is_null($grade2)) {
                echo "Erro: Nota da Unidade 2 inválida ou não encontrada para o aluno {$student->firstname} {$student->lastname}.\n";
            } else {
                //echo "Nota da Unidade 2 para o aluno {$student->firstname} {$student->lastname}: {$grade2}\n";
            }

            if (is_null($grade3)) {
                echo "Erro: Nota da Unidade 3 inválida ou não encontrada para o aluno {$student->firstname} {$student->lastname}.\n";
            } else {
                //echo "Nota da Unidade 3 para o aluno {$student->firstname} {$student->lastname}: {$grade3}\n";
            }
            */
            
            // Adiciona os dados do aluno à planilha
            $data = [
                '', (string) $matricula, // Matrícula sempre como string
                $fullname,
                $grade1, //grade1
                $grade2, //grade2
                $grade3, //grade3
                $quarta_prova_grade, //Recuperação
                '',//$result_formula,
                0,  // Número de faltas padrão
                '', //$status_formula,
            ];

            // Escreve os dados do aluno na planilha
            foreach ($data as $col => $cell) {
                if ($col === 1) { // Matrícula sempre como string
                    $worksheet->write_string($row, $col, (string) $cell);
                } elseif (is_numeric($cell)) {
                    $worksheet->write_number($row, $col, $cell);
                } else {
                    $worksheet->write_string($row, $col, $cell);
                }
            }
            
            $row++; // Avança para a próxima linha
            $studentCount++; // Incrementa o contador de matrículas
        }
        
        // Após a última matrícula, adicionar uma única linha em branco
        if ($totalStudents > 0) {
            $num_colunas = count($data); // Número total de colunas
            for ($col = 0; $col < $num_colunas; $col++) {
                $worksheet->write_string($row, $col, ''); // Escreve uma célula vazia
            }
            $row++; // Avança para a próxima linha após o espaço em branco
        }
        // Finaliza e exporta o arquivo
        $workbook->close();
        exit;

   
    }


