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

        // Botões de download e mensagens de sucesso e info
        $this->content->text .= html_writer::div(
            html_writer::div(get_string('aguarde_download', 'block_exportgradesigaa'), 'alert alert-info', [
                'id' => 'info-message',
                'style' => 'display:none; margin-top:10px;'
            ]) .
                html_writer::div(get_string('download_sucesso', 'block_exportgradesigaa'), 'alert alert-success', [
                    'id' => 'success-message',
                    'style' => 'display:none; margin-top:10px;'
                ]) .
                html_writer::link($url, get_string('download_xls', 'block_exportgradesigaa'), [
                    'class' => 'btn btn-primary',
                    'id' => 'download-xls-button',
                    'data-url' => $url
                ]),
            '',
            ['style' => 'margin-bottom: 10px;']
        );

        
        
        $this->content->text .= html_writer::div(
            html_writer::link('https://dead.uern.br/notasparasigaa', get_string('view_tutorial', 'block_exportgradesigaa'), array('class' => 'btn btn-secondary'))
        );
        
        $this->content->text .= html_writer::div($OUTPUT->pix_icon('i/warning', 'Edit me', 'moodle') . "ATENÇÃO", "gradessigaawarning", ['style' => 'margin-top: 20px; font-weight: bold;']);
        $this->content->text .= html_writer::div(get_string('alerta_beta', 'block_exportgradesigaa'), "gradessigaawarning", []);




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

                        // Mostra mensagem de aguarde
                        if (infoMsg) {
                            infoMsg.style.display = 'block';
                        }

                        // Remove cookie antigo
                        document.cookie = 'fileDownload=; Max-Age=0; path=/';

                        // Cria iframe invisível
                        const iframe = document.createElement('iframe');
                        iframe.style.display = 'none';
                        iframe.src = url;
                        document.body.appendChild(iframe);

                        // Espera o cookie do download
                        const interval = setInterval(() => {
                            if (document.cookie.includes('fileDownload=true')) {
                                clearInterval(interval);
                                btn.disabled = false;

                                // Oculta a mensagem de info e mostra sucesso
                                if (infoMsg) {
                                    infoMsg.style.display = 'none';
                                }
                                if (successMsg) {
                                    successMsg.style.display = 'block';
                                }

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
