<?php
// Arquivo db/access.php

defined('MOODLE_INTERNAL') || die();

$capabilities = [
    // Define a capacidade do plugin
    'block/exportgradesigaa:view' => [
        'captype' => 'read', // Capacidade de leitura
        'contextlevel' => CONTEXT_COURSE, // Contexto limitado ao curso
        'archetypes' => [
            'teacher' => CAP_ALLOW, // Permite acesso apenas aos professores
            'student' => CAP_PROHIBIT, // Proíbe alunos
            'guest' => CAP_PROHIBIT, // Proíbe visitantes
            'user' => CAP_PROHIBIT, // Proíbe usuários gerais
            'manager' => CAP_ALLOW, // Permite acesso aos gerentes
            'admin' => CAP_ALLOW,  // Permite acesso aos administradores
        ],
        // Clona permissões padrão para consistência (opcional)
        'clonepermissionsfrom' => 'moodle/grade:viewall',
    ],
];
