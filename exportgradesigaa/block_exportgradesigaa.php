<?php

class block_exportgradesigaa extends block_base
{
    public function init()
    {
        $this->title = get_string('pluginname', 'block_exportgradesigaa');
    }

    public function get_content()
    {
        if ($this->content !== null) {
            return $this->content;
        }

        global $COURSE, $USER, $DB;

        // Verifica se o usuário tem o papel de professor, administrador ou gerente
        $context = context_course::instance($COURSE->id);
        $teacher_role = $DB->get_record('role', ['shortname' => 'editingteacher']);
        $admin_role = $DB->get_record('role', ['shortname' => 'manager']);
        $manager_role = $DB->get_record('role', ['shortname' => 'coursecreator']); // Gerentes geralmente têm o papel 'coursecreator'

        $is_teacher = user_has_role_assignment($USER->id, $teacher_role->id, $context->id);
        $is_admin = user_has_role_assignment($USER->id, $admin_role->id, $context->id);
        $is_manager = user_has_role_assignment($USER->id, $manager_role->id, $context->id);

        // Exibe o bloco apenas para professores, administradores ou gerentes
        if (!$is_teacher && !$is_admin && !$is_manager) {
            return null;
        }

        $this->content = new stdClass;
        $this->content->text = '';

        // Gera o link para o download do CSV
        $url = new moodle_url('/blocks/exportgradesigaa/download_grades.php', array('courseid' => $COURSE->id));
        $this->content->text .= html_writer::link($url, get_string('download_xls', 'block_exportgradesigaa'), array('class' => 'btn btn-primary'));

        // Link para o tutorial
        $this->content->text .= html_writer::link('https://dead.uern.br', get_string('view_tutorial', 'block_exportgradesigaa'), array('class' => 'btn btn-secondary'));

        return $this->content;
    }
}
