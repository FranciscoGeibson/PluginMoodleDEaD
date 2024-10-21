<?php
class block_gradeoverview extends block_base {

    public function init() {
        $this->title = get_string('pluginname', 'block_gradeoverview');
    }

    public function get_content() {
        if ($this->content !== null) {
            return $this->content;
        }

        global $COURSE, $DB, $OUTPUT;

        $this->content = new stdClass;
        $this->content->text = '';

        // Pega todas as notas dos estudantes da turma
        $courseid = $COURSE->id;  // ID da turma
        $sql = "SELECT u.id, u.firstname, u.lastname, g.finalgrade
                FROM {user} u
                JOIN {grade_grades} g ON u.id = g.userid
                JOIN {grade_items} gi ON g.itemid = gi.id
                WHERE gi.courseid = :courseid";

        $params = ['courseid' => $courseid];
        $grades = $DB->get_records_sql($sql, $params);

        if ($grades) {
            foreach ($grades as $grade) {
                $this->content->text .= $grade->firstname . ' ' . $grade->lastname . ': ' . $grade->finalgrade . '<br>';
            }
        } else {
            $this->content->text .= 'Nenhuma nota disponível.<br>';
        }

        // Gera o link para o download do CSV
        $url = new moodle_url('/blocks/block_gradeoverview/download_grades.php', array('courseid' => $COURSE->id));
        $this->content->text .= html_writer::link($url, 'Baixar Planilha de Notas', array('class' => 'btn btn-primary'));

        return $this->content;
    }
}
