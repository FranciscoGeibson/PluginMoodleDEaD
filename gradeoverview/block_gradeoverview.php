<?php


class block_gradeoverview extends block_base
{

    public function init()
    {
        $this->title = get_string('pluginname', 'block_gradeoverview');
    }

    public function get_content()
    {
        if ($this->content !== null) {
            return $this->content;
        }

        global $COURSE, $DB, $OUTPUT;

        $this->content = new stdClass;
        $this->content->text = '';

        // Gera o link para o download do CSV
        $url = new moodle_url('/blocks/gradeoverview/download_grades.php', array('courseid' => $COURSE->id));
        $this->content->text .= html_writer::link($url, 'Baixar Planilha de Notas', array('class' => 'btn btn-primary'));
        $this->content->text .= html_writer::link('https://dead.uern.br', 'Ver tutorial de utilização', array('class' => 'btn btn-secondary'));

        return $this->content;
    }
}
