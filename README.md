# Plugin Moodle - Visão Geral das Notas  

##  Melhorias Planejadas  

- **Reconhecimento Dinâmico**  
  Adicionar o reconhecimento automático do nome breve do curso e da carga horária diretamente no código.  

- **Inclusão de Colunas**  
  Adicionar a coluna "Recuperação" no arquivo XLS gerado.  

- **Automatização de Notas**  
  Integrar o sistema para reconhecer automaticamente as notas dos alunos, utilizando as funcionalidades da API do Moodle, sem necessidade de cálculos manuais no código.  

O **Visão Geral das Notas** é um plugin desenvolvido para a plataforma Moodle que facilita a visualização e exportação das notas dos alunos de um curso específico. Ele foi criado para oferecer uma experiência mais prática e eficiente para educadores e administradores, permitindo a análise rápida do desempenho acadêmico.  

##  Funcionalidades  

- **Visualização Dinâmica**  
  Permite a visualização das notas finais de todos os alunos diretamente na interface do Moodle, com suporte a cursos específicos.  

- **Exportação para XLS**  
  Gera um link que permite o download de uma planilha em formato XLS, contendo os nomes dos alunos e suas respectivas notas finais.  

## 🛠 Estrutura do Código  

- `block_gradeoverview.php`  
  Contém a lógica principal do bloco, incluindo a consulta às notas dos alunos.  

- `download_grades.php`  
  Gera a planilha XLS com as notas dos alunos para download.  
