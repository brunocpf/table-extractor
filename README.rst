table-extractor
=============================
:Authors: 
    Bruno Cesar Pimenta Fernandes <brunocpf@dcc.ufmg.br>
:Version: 1.0
Interface para extrair uma tabela .csv de um documento PDF usando coordenadas absolutas.

Instruções
----------------------------------
#. Execute ``extractor.exe`` na pasta ``dist``. A interface do extrator vai abrir.
#. Selecione o arquivo PDF no campo indicado.
#. Se desejar carregar opções pré-definidas, selecione um arquivo de configuração ``.json`` (dois arquivos pré-configurados estão disponibilizados na pasta ``settings files``) no campo indicado.
#. Modifique os campos, se necessário. Você pode salvar suas modificações em um novo arquivo ``.json`` clicando no botão "Salvar configurações".
#. Clique no botão "< Extrair tabela >" e indique o arquivo ``.csv`` de saída para começar a extração.
#. Quando a barra de progresso chegar ao final, a tabela será criada.


Observações e dicas
----------------------------------
* Use "-1" no campo "Página final" para indicar que todas as páginas a partir da inicial devem ser processadas.
* As coordenadas são absolutas e estão no formato padrão PDF, em pt (1 pt = 1 polegada/72).
* O campo "Alturas" é uma lista de alturas de tabelas separadas por vírgulas, no formato: ``numeroDaPagina1: altura1, numeroDaPagina2: altura2, etc``. Use "Demais" como numero de página para indicar a altura das páginas não especificadas. Use 0 como altura para ignorar a página correspondente.
* Em caso de dúvida ou algum problema encontrado, mande um e-mail.
