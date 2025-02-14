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
    //return $grade ? number_format($grade->finalgrade, 1, ',', '') : null;
    return $grade ? round($grade->finalgrade, 1) : null; // Retorna um valor numérico arredondado

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

// Criar objeto ODS
$filename = "notas_{$course->shortname}.ods";
$workbook = new MoodleODSWorkbook(format_string($course->fullname, true));
$workbook->send($filename);
$worksheet = $workbook->add_worksheet("Notas");

$course_duration = get_course_duration($courseid);

if(intval($course_duration) > 45){
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

    // Dados dos alunos
    $students = get_role_users($DB->get_record("role", ['shortname' => 'student'])->id, $context);


    $unit1_category_id = get_category_id_by_name('Unidade 1', $courseid);
    $unit2_category_id = get_category_id_by_name('Unidade 2', $courseid);
    $unit3_category_id = get_category_id_by_name('Unidade 3', $courseid);
    //Pegar a REC




    //Usar hifen como delimitador, para o for correr e verificar até ele.
    //Verificar e não permitir quando houver erros de escrita.
    //Verificar casas decimais da turma


    //Sobre as Unidades, a carga horária <= 45 é de 2 Unidades. Maior que > 45, é de 3 Unidades.


    // Obtém a lista de alunos matriculados no curso
    //$students = get_role_users($DB->get_record("role", ['shortname' => 'student'])->id, $context);
    // Contar quantos alunos existem antes de escrever na planilha
    $totalStudents = count($students);
    $studentCount = 0; // Contador para acompanhar a matrícula atual

    foreach ($students as $student) {
        $fullname = strtoupper($student->firstname . ' ' . $student->lastname);

        // Obter a matrícula do aluno a partir do campo personalizado
        $matricula = get_student_enrollment($student->id);

        // Caso a matrícula seja nula, exibir uma mensagem de erro ou usar um valor padrão
        if (is_null($matricula)) {
            $matricula = 'Matrícula não encontrada';
        } else {
            $matricula = (string) $matricula; // Converte a matrícula para string
        }
        

        $grade1 = get_student_grade($student->id, $unit1_category_id);
        $grade2 = get_student_grade($student->id, $unit2_category_id);
        $grade3 = get_student_grade($student->id, $unit3_category_id);

        /*$result_formula = <<<EOD
        =SE(OU(D13="-"; D13=""; E13="-"; E13=""; F13="-"; F13=""); "-"; SE(OU(G13=""; G13<0; G13="-"); (ARRED((((D13*4*10)+(E13*5*10)+(F13*6*10))/150)*10; 0)/10); (ARRED(((SE(MÍNIMO(D13; E13; F13)=D13; (D13*4*10)+(E13*5*10)+(F13*6*10)-(D13*6*10)+(G13*6*10); SE(MÍNIMO(D13; E13; F13)=E13; (D13*4*10)+(E13*5*10)+(F13*6*10)-(E13*6*10)+(G13*6*10); SE(MÍNIMO(D13; E13; F13)=F13; (D13*4*10)+(E13*5*10)+(F13*6*10)-(F13*6*10)+(G13*6*10); ))))/150)*10; 0)/10)))
        EOD;
        //=SE(OU(H13="-";H13="");"-";SE(I13>18;SE(H13>=6;SE((OU(D13<7;E13<7;F13<7));"RENF";"REPF");"REMF");SE(H13>=7;"APR";SE(H13>=6;SE(E((OU(D13<7;E13<7;F13<7));(OU(G13="-";G13="")));"REC";SE(OU(G13="-";G13="");SE(H13>=7;"APR";"APRN");SE(G13>=7;"APR";"REPN")));SE(H13>=7;SE(OU(G13="-";G13="");"REC";"REP");"REP")))))

        $status_formula = <<<EOD
        =SE(OU(H13="-";H13="");"-";SE(I13>18;SE(H13>=6;SE((OU(D13<7;E13<7;F13<7));"RENF";"REPF");"REMF");SE(H13>=7;"APR";SE(H13>=6;SE(E((OU(D13<7;E13<7;F13<7));(OU(G13="-";G13="")));"REC";SE(OU(G13="-";G13="");SE(H13>=7;"APR";"APRN");SE(G13>=7;"APR";"REPN")));SE(H13>=7;SE(OU(G13="-";G13="");"REC";"REP");"REP")))))
        EOD;
        */

        // Definir um row específico para fórmulas (linha 12 em vez de 11)
        $formula_row = $row + 1;

        /*$result_formula = <<<EOD
        =SE(OU(D{$formula_row}="-"; D{$formula_row}=""; E{$formula_row}="-"; E{$formula_row}=""; F{$formula_row}="-"; F{$formula_row}=""); "-"; SE(OU(G{$formula_row}=""; G{$formula_row}<0; G{$formula_row}="-"); (ARRED((((D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10))/150)*10; 0)/10); (ARRED(((SE(MÍNIMO(D{$formula_row}; E{$formula_row}; F{$formula_row})=D{$formula_row}; (D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10)-(D{$formula_row}*6*10)+(G{$formula_row}*6*10); SE(MÍNIMO(D{$formula_row}; E{$formula_row}; F{$formula_row})=E{$formula_row}; (D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10)-(E{$formula_row}*6*10)+(G{$formula_row}*6*10); SE(MÍNIMO(D{$formula_row}; E{$formula_row}; F{$formula_row})=F{$formula_row}; (D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10)-(F{$formula_row}*6*10)+(G{$formula_row}*6*10); ))))/150)*10; 0)/10)))
        EOD;

        $status_formula = <<<EOD
        =SE(OU(H{$formula_row}="-"; H{$formula_row}=""); "-"; SE(I{$formula_row}>18; SE(H{$formula_row}>=6; SE((OU(D{$formula_row}<7; E{$formula_row}<7; F{$formula_row}<7)); "RENF"; "REPF"); "REMF"); SE(H{$formula_row}>=7; "APR"; SE(H{$formula_row}>=6; SE(E((OU(D{$formula_row}<7; E{$formula_row}<7; F{$formula_row}<7));(OU(G{$formula_row}="-"; G{$formula_row}="")));"REC"; SE(OU(G{$formula_row}="-"; G{$formula_row}=""); SE(H{$formula_row}>=7; "APR"; "APRN"); SE(G{$formula_row}>=7; "APR"; "REPN"))); SE(H{$formula_row}>=7; SE(OU(G{$formula_row}="-"; G{$formula_row}=""); "REC"; "REP"); "REP"))))))
        EOD;
        */

        $result_formula = "=SE(OU(D{$formula_row}=\"-\"; D{$formula_row}=\"\"; E{$formula_row}=\"-\"; E{$formula_row}=\"\"; F{$formula_row}=\"-\"; F{$formula_row}=\"\"); \"-\"; SE(OU(G{$formula_row}=\"\"; G{$formula_row}<0; G{$formula_row}=\"-\"); (ARRED((((D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10))/150)*10; 0)/10); (ARRED(((SE(MÍNIMO(D{$formula_row}; E{$formula_row}; F{$formula_row})=D{$formula_row}; (D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10)-(D{$formula_row}*6*10)+(G{$formula_row}*6*10); SE(MÍNIMO(D{$formula_row}; E{$formula_row}; F{$formula_row})=E{$formula_row}; (D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10)-(E{$formula_row}*6*10)+(G{$formula_row}*6*10); SE(MÍNIMO(D{$formula_row}; E{$formula_row}; F{$formula_row})=F{$formula_row}; (D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10)-(F{$formula_row}*6*10)+(G{$formula_row}*6*10); ))))/150)*10; 0)/10))))";

        $status_formula = "=SE(OU(H{$formula_row}=\"-\"; H{$formula_row}=\"\"); \"-\"; SE(I{$formula_row}>18; SE(H{$formula_row}>=6; SE((OU(D{$formula_row}<7; E{$formula_row}<7; F{$formula_row}<7)); \"RENF\"; \"REPF\"); \"REMF\"); SE(H{$formula_row}>=7; \"APR\"; SE(H{$formula_row}>=6; SE(E((OU(D{$formula_row}<7; E{$formula_row}<7; F{$formula_row}<7));(OU(G{$formula_row}=\"-\"; G{$formula_row}=\"\")));\"REC\"; SE(OU(G{$formula_row}=\"-\"; G{$formula_row}=\"\"); SE(H{$formula_row}>=7; \"APR\"; \"APRN\"); SE(G{$formula_row}>=7; \"APR\"; \"REPN\"))); SE(H{$formula_row}>=7; SE(OU(G{$formula_row}=\"-\"; G{$formula_row}=\"\"); \"REC\"; \"REP\"); \"REP\")))))))";

        //$result_formula = "=IF(OR(D{$row}=\"-\", D{$row}=\"\", E{$row}=\"-\", E{$row}=\"\", F{$row}=\"-\", F{$row}=\"\"), \"-\", IF(OR(G{$row}=\"\", G{$row}<0, G{$row}=\"-\", G{$row}=\"\"), ROUND((((D{$row}*4*10)+(E{$row}*5*10)+(F{$row}*6*10))/150)*10, 0)/10, ROUND((((D{$row}*4*10)+(E{$row}*5*10)+(F{$row}*6*10))/150)*10, 0)/10))";
        
        // Adiciona os dados do aluno à planilha
        $data = [
            '', (string) $matricula, // Matrícula sempre como string
            $fullname,
            $grade1, //grade1
            $grade2, //grade2
            $grade3, //grade3
            '', //Rec
            '',//$result_formula,
            0,  // Número de faltas padrão
            '', //$status_formula,
        ];


        /*
        $data2 = [
            '', (string) $matricula, // Matrícula sempre como string
            $fullname,
            $grade1, //grade1
            $grade2, //grade2
            '', //Rec
            '',//$result_formula,
            0,  // Número de faltas padrão
            '', //$status_formula,
        ];
        */
        //$col = 0;
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

} else if(intval($course_duration) <= 45){
    //Colocar a pesquisa da turma recursiva, T2...T3...T4...T5
    // Determinar a turma
    $course_shortname = htmlspecialchars($course->shortname);
    // Remove "(T1)", "(T2)", "(T100)", etc., do course_shortname
    $course_shortname_cleaned = preg_replace('/\s*\(T\d+\)/', '', $course_shortname);

    // Extrai o número após o "(T" para definir a turma
    if (preg_match('/\(T(\d+)\)/', $course_shortname, $matches)) {
        $course_turma = str_pad($matches[1], 2, '0', STR_PAD_LEFT); // Garante que o número tenha 2 dígitos
    } else {
        $course_turma = '01'; // Valor padrão caso não encontre "(TXX)"
    }

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
    $headers = ['', 'Matrícula', 'Nome', 'Unid. 1', 'Unid. 2', 'Rec.', 'Resultado', 'Faltas', 'Sit.'];
    $col = 0;
    foreach ($headers as $header) {
        $worksheet->write_string($row, $col++, $header);
    }
    $row++;

    // Dados dos alunos
    $students = get_role_users($DB->get_record("role", ['shortname' => 'student'])->id, $context);


    $unit1_category_id = get_category_id_by_name('Unidade 1', $courseid);
    $unit2_category_id = get_category_id_by_name('Unidade 2', $courseid);
    //Pegar a REC




    //Usar hifen como delimitador, para o for correr e verificar até ele.
    //Verificar e não permitir quando houver erros de escrita.
    //Verificar casas decimais da turma


    //Sobre as Unidades, a carga horária <= 45 é de 2 Unidades. Maior que > 45, é de 3 Unidades.


    // Obtém a lista de alunos matriculados no curso
    //$students = get_role_users($DB->get_record("role", ['shortname' => 'student'])->id, $context);
    // Contar quantos alunos existem antes de escrever na planilha
    $totalStudents = count($students);
    $studentCount = 0; // Contador para acompanhar a matrícula atual

    foreach ($students as $student) {
        $fullname = strtoupper($student->firstname . ' ' . $student->lastname);

        // Obter a matrícula do aluno a partir do campo personalizado
        $matricula = get_student_enrollment($student->id);

        // Caso a matrícula seja nula, exibir uma mensagem de erro ou usar um valor padrão
        if (is_null($matricula)) {
            $matricula = 'Matrícula não encontrada';
        } else {
            $matricula = (string) $matricula; // Converte a matrícula para string
        }
        

        $grade1 = get_student_grade($student->id, $unit1_category_id);
        $grade2 = get_student_grade($student->id, $unit2_category_id);

        /*$result_formula = <<<EOD
        =SE(OU(D13="-"; D13=""; E13="-"; E13=""; F13="-"; F13=""); "-"; SE(OU(G13=""; G13<0; G13="-"); (ARRED((((D13*4*10)+(E13*5*10)+(F13*6*10))/150)*10; 0)/10); (ARRED(((SE(MÍNIMO(D13; E13; F13)=D13; (D13*4*10)+(E13*5*10)+(F13*6*10)-(D13*6*10)+(G13*6*10); SE(MÍNIMO(D13; E13; F13)=E13; (D13*4*10)+(E13*5*10)+(F13*6*10)-(E13*6*10)+(G13*6*10); SE(MÍNIMO(D13; E13; F13)=F13; (D13*4*10)+(E13*5*10)+(F13*6*10)-(F13*6*10)+(G13*6*10); ))))/150)*10; 0)/10)))
        EOD;
        //=SE(OU(H13="-";H13="");"-";SE(I13>18;SE(H13>=6;SE((OU(D13<7;E13<7;F13<7));"RENF";"REPF");"REMF");SE(H13>=7;"APR";SE(H13>=6;SE(E((OU(D13<7;E13<7;F13<7));(OU(G13="-";G13="")));"REC";SE(OU(G13="-";G13="");SE(H13>=7;"APR";"APRN");SE(G13>=7;"APR";"REPN")));SE(H13>=7;SE(OU(G13="-";G13="");"REC";"REP");"REP")))))

        $status_formula = <<<EOD
        =SE(OU(H13="-";H13="");"-";SE(I13>18;SE(H13>=6;SE((OU(D13<7;E13<7;F13<7));"RENF";"REPF");"REMF");SE(H13>=7;"APR";SE(H13>=6;SE(E((OU(D13<7;E13<7;F13<7));(OU(G13="-";G13="")));"REC";SE(OU(G13="-";G13="");SE(H13>=7;"APR";"APRN");SE(G13>=7;"APR";"REPN")));SE(H13>=7;SE(OU(G13="-";G13="");"REC";"REP");"REP")))))
        EOD;
        */

        // Definir um row específico para fórmulas (linha 12 em vez de 11)
        $formula_row = $row + 1;

        /*$result_formula = <<<EOD
        =SE(OU(D{$formula_row}="-"; D{$formula_row}=""; E{$formula_row}="-"; E{$formula_row}=""; F{$formula_row}="-"; F{$formula_row}=""); "-"; SE(OU(G{$formula_row}=""; G{$formula_row}<0; G{$formula_row}="-"); (ARRED((((D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10))/150)*10; 0)/10); (ARRED(((SE(MÍNIMO(D{$formula_row}; E{$formula_row}; F{$formula_row})=D{$formula_row}; (D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10)-(D{$formula_row}*6*10)+(G{$formula_row}*6*10); SE(MÍNIMO(D{$formula_row}; E{$formula_row}; F{$formula_row})=E{$formula_row}; (D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10)-(E{$formula_row}*6*10)+(G{$formula_row}*6*10); SE(MÍNIMO(D{$formula_row}; E{$formula_row}; F{$formula_row})=F{$formula_row}; (D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10)-(F{$formula_row}*6*10)+(G{$formula_row}*6*10); ))))/150)*10; 0)/10)))
        EOD;

        $status_formula = <<<EOD
        =SE(OU(H{$formula_row}="-"; H{$formula_row}=""); "-"; SE(I{$formula_row}>18; SE(H{$formula_row}>=6; SE((OU(D{$formula_row}<7; E{$formula_row}<7; F{$formula_row}<7)); "RENF"; "REPF"); "REMF"); SE(H{$formula_row}>=7; "APR"; SE(H{$formula_row}>=6; SE(E((OU(D{$formula_row}<7; E{$formula_row}<7; F{$formula_row}<7));(OU(G{$formula_row}="-"; G{$formula_row}="")));"REC"; SE(OU(G{$formula_row}="-"; G{$formula_row}=""); SE(H{$formula_row}>=7; "APR"; "APRN"); SE(G{$formula_row}>=7; "APR"; "REPN"))); SE(H{$formula_row}>=7; SE(OU(G{$formula_row}="-"; G{$formula_row}=""); "REC"; "REP"); "REP"))))))
        EOD;
        */

        $result_formula = "=SE(OU(D{$formula_row}=\"-\"; D{$formula_row}=\"\"; E{$formula_row}=\"-\"; E{$formula_row}=\"\"; F{$formula_row}=\"-\"; F{$formula_row}=\"\"); \"-\"; SE(OU(G{$formula_row}=\"\"; G{$formula_row}<0; G{$formula_row}=\"-\"); (ARRED((((D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10))/150)*10; 0)/10); (ARRED(((SE(MÍNIMO(D{$formula_row}; E{$formula_row}; F{$formula_row})=D{$formula_row}; (D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10)-(D{$formula_row}*6*10)+(G{$formula_row}*6*10); SE(MÍNIMO(D{$formula_row}; E{$formula_row}; F{$formula_row})=E{$formula_row}; (D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10)-(E{$formula_row}*6*10)+(G{$formula_row}*6*10); SE(MÍNIMO(D{$formula_row}; E{$formula_row}; F{$formula_row})=F{$formula_row}; (D{$formula_row}*4*10)+(E{$formula_row}*5*10)+(F{$formula_row}*6*10)-(F{$formula_row}*6*10)+(G{$formula_row}*6*10); ))))/150)*10; 0)/10))))";

        $status_formula = "=SE(OU(H{$formula_row}=\"-\"; H{$formula_row}=\"\"); \"-\"; SE(I{$formula_row}>18; SE(H{$formula_row}>=6; SE((OU(D{$formula_row}<7; E{$formula_row}<7; F{$formula_row}<7)); \"RENF\"; \"REPF\"); \"REMF\"); SE(H{$formula_row}>=7; \"APR\"; SE(H{$formula_row}>=6; SE(E((OU(D{$formula_row}<7; E{$formula_row}<7; F{$formula_row}<7));(OU(G{$formula_row}=\"-\"; G{$formula_row}=\"\")));\"REC\"; SE(OU(G{$formula_row}=\"-\"; G{$formula_row}=\"\"); SE(H{$formula_row}>=7; \"APR\"; \"APRN\"); SE(G{$formula_row}>=7; \"APR\"; \"REPN\"))); SE(H{$formula_row}>=7; SE(OU(G{$formula_row}=\"-\"; G{$formula_row}=\"\"); \"REC\"; \"REP\"); \"REP\")))))))";

        //$result_formula = "=IF(OR(D{$row}=\"-\", D{$row}=\"\", E{$row}=\"-\", E{$row}=\"\", F{$row}=\"-\", F{$row}=\"\"), \"-\", IF(OR(G{$row}=\"\", G{$row}<0, G{$row}=\"-\", G{$row}=\"\"), ROUND((((D{$row}*4*10)+(E{$row}*5*10)+(F{$row}*6*10))/150)*10, 0)/10, ROUND((((D{$row}*4*10)+(E{$row}*5*10)+(F{$row}*6*10))/150)*10, 0)/10))";
        
        // Adiciona os dados do aluno à planilha
        
        
        $data2 = [
            '', (string) $matricula, // Matrícula sempre como string
            $fullname,
            $grade1, //grade1
            $grade2, //grade2
            '', //Rec
            '',//$result_formula,
            0,  // Número de faltas padrão
            '', //$status_formula,
        ];
        
        //$col = 0;
        // Escreve os dados do aluno na planilha
        foreach ($data2 as $col => $cell) {
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
        $num_colunas = count($data2); // Número total de colunas
        for ($col = 0; $col < $num_colunas; $col++) {
            $worksheet->write_string($row, $col, ''); // Escreve uma célula vazia
        }
        $row++; // Avança para a próxima linha após o espaço em branco
    }
    // Finaliza e exporta o arquivo
    $workbook->close();
    exit;
}


