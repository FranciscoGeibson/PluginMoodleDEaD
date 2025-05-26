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

        // Verifica a capacidade definida no access.php para o contexto do curso
        $context = context_course::instance($COURSE->id);
        if (!has_capability('block/exportgradesigaa:view', $context)) {
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
                        var spinner = document.getElementById('loading-spinner');
                        if (spinner) spinner.style.display = 'block';

                        btn.classList.add('disabled');
                        btn.setAttribute('aria-disabled', 'true');
                        btn.style.pointerEvents = 'none';

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
