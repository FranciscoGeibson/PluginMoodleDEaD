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

        global $COURSE, $USER, $DB, $OUTPUT, $SESSION;

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

        // Exibe notificação de erro, se existir
        if (!empty($SESSION->gradeoverview_error)) {
            $this->content->text .= $OUTPUT->notification($SESSION->gradeoverview_error, get_string('servidor_nao_encontrado', 'block_exportgradesigaa'), 'error');
            
            unset($SESSION->gradeoverview_error);
        }

        // Gera o link para o download do CSV
        $url = new moodle_url('/blocks/exportgradesigaa/download_grades.php', array('courseid' => $COURSE->id));

        $this->content->text .= html_writer::div(
            html_writer::link($url, get_string('download_xls', 'block_exportgradesigaa'), array(
                'class' => 'btn btn-primary',
                'id' => 'download-xls-button'
            )),
            '',
            array('style' => 'margin-bottom: 10px;')
        );

        $this->content->text .= html_writer::div(
            '<div id="loading-spinner" style="display:none; margin-top:10px;">
                <span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span>
                <span style="margin-left: 8px;">' . get_string('aguarde_download', 'block_exportgradesigaa') . '</span>
            </div>'
        );
        
        // Link para o tutorial
        $this->content->text .= html_writer::div(
            html_writer::link('https://dead.uern.br/notasparasigaa', get_string('view_tutorial', 'block_exportgradesigaa'), array('class' => 'btn btn-secondary'))
        );

        // Adiciona o script para o botão de download e animação de loading
        $this->content->text .= html_writer::script("
            document.addEventListener('DOMContentLoaded', function() {
                var btn = document.getElementById('download-xls-button');
                if (btn) {
                    btn.addEventListener('click', function(event) {
                        // Mostra o loading e desativa o botão
                        var spinner = document.getElementById('loading-spinner');
                        if (spinner) spinner.style.display = 'block';

                        btn.classList.add('disabled');
                        btn.setAttribute('aria-disabled', 'true');
                        btn.style.pointerEvents = 'none';

                        // Permite que o link ainda funcione
                        setTimeout(function() {
                            window.location.href = btn.href;
                        }, 100);

                        event.preventDefault();
                    });
                }
            });
        ");

        return $this->content;
    }
}
