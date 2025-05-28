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

        // Verifica se o usuário tem permissão para visualizar o bloco
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

        // Botões de download e mensagens de sucesso e info
        $this->content->text .= html_writer::div(
            html_writer::div($OUTPUT->pix_icon('i/loading_small', 'loading', 'moodle') . get_string('aguarde_download', 'block_exportgradesigaa'), 'alert alert-info', [
                'id' => 'info-message',
                'style' => 'display:none; margin-top:10px;'
            ]) .
            html_writer::div($OUTPUT->pix_icon('i/grade_correct', 'tick', 'moodle') . get_string('download_sucesso', 'block_exportgradesigaa'), 'alert alert-success', [
                'id' => 'success-message',
                'style' => 'display:none; margin-top:10px;'
            ]),
            '',
            ['style' => 'margin-bottom: 10px;']
        );

        $this->content->text .= html_writer::div($OUTPUT->pix_icon('i/warning', 'warning', 'moodle') . "ATENÇÃO", "gradessigaawarning", ['style' => 'font-weight: bold;']);
        $this->content->text .= html_writer::div(get_string('alerta_beta', 'block_exportgradesigaa'), "gradessigaawarning", ['style' => "margin-bottom: 20px;"]);

        $this->content->text .= html_writer::start_tag("a", ['href' => $url]);
        $this->content->text .= html_writer::start_tag("button", ['class' => 'btn btn-primary', 'id' => 'download-xls-button', 'style' => 'margin-bottom: 10px;', 'data-url' => $url->out(false)]);
        $this->content->text .= get_string('download_xls', 'block_exportgradesigaa');
        $this->content->text .= html_writer::end_tag("button");
        $this->content->text .= html_writer::end_tag("a");

        $this->content->text .= html_writer::start_tag("a", ['href' => 'https://dead.uern.br/notasparasigaa', 'target' => '_blank']);
        $this->content->text .= html_writer::start_tag("button", ['class' => 'btn btn-secondary', 'id' => 'tutorial-xls-button', 'style' => 'margin-bottom: 10px;']);
        $this->content->text .= get_string('view_tutorial', 'block_exportgradesigaa');
        $this->content->text .= html_writer::end_tag("button");
        $this->content->text .= html_writer::end_tag("a");

        // Adiciona o script para o botão de download e animação de loading
        $this->content->text .= html_writer::script("
            document.addEventListener('DOMContentLoaded', function() {
                const btn = document.getElementById('download-xls-button');
                const infoMsg = document.getElementById('info-message');
                const successMsg = document.getElementById('success-message');

                if (btn) {
                    btn.addEventListener('click', function() {
                        const url = btn.getAttribute('data-url');

                        btn.disabled = true;

                        if (infoMsg) infoMsg.style.display = 'block';
                        document.cookie = 'fileDownload=; Max-Age=0; path=/';

                        const iframe = document.createElement('iframe');
                        iframe.style.display = 'none';
                        iframe.src = url;
                        document.body.appendChild(iframe);

                        const interval = setInterval(() => {
                            if (document.cookie.includes('fileDownload=true')) {
                                clearInterval(interval);
                                btn.disabled = false;

                                if (infoMsg) infoMsg.style.display = 'none';
                                if (successMsg) successMsg.style.display = 'block';

                                document.cookie = 'fileDownload=; Max-Age=0; path=/';
                                iframe.remove();
                            }
                        }, 500);
                    });
                }
            });
        ");

        return $this->content;
    }
}
