# PlanilhaEstoqueControle: Scripts Google Apps Script para Gerenciamento de Estoque no Google Sheets

[![GitHub Repository](https://img.shields.io/badge/GitHub-Repository-blue?logo=github)](https://github.com/FernandaPaulaDeCarvalho/PlanilhaEstoqueControle)

Este repositório contém dois scripts Google Apps Script que automatizam a criação e configuração de uma planilha Google Sheets para gerenciamento de estoque, facilitando o controle de entrada e saída de produtos, além de gerar tabelas de estoque e relatórios. A planilha é criada para um ano específico e projetada para ser simples de usar, mesmo para usuários sem experiência em planilhas ou programação.

**IMPORTANTE:** A configuração da planilha requer a execução de dois scripts distintos, em etapas específicas. Certifique-se de seguir as instruções detalhadas abaixo para configurar e utilizar a planilha corretamente.

## Arquivos Incluídos

*   **ScriptPlanilhaEstoqueControle.gs:**
    *   Responsável pela criação e formatação da planilha, incluindo a criação das abas, a aplicação de formatação condicional, a validação de dados e a proteção das abas.
    *   Deve ser executado em três etapas distintas (`primeiraEtapa`, `segundaEtapa` e `terceiraEtapa`) dentro do Google Apps Script.

*   **ScriptBotaoAmareloGravar.gs:**
    *   Responsável por adicionar a funcionalidade aos botões "GRAVAR" nas abas "ENTRADA" e "SAÍDA", permitindo gravar os dados inseridos.
    *   Deve ser copiado e colado no editor de script da planilha após a execução das três etapas do `ScriptPlanilhaEstoqueControle`.

*   **ImagemBotaoAmareloGravar.png:**
    *   Imagem utilizada para criar os botões "GRAVAR" e deve ser associada ao `ScriptBotaoAmareloGravar`.
    *   O `ScriptPlanilhaEstoqueControle` associa automaticamente o `ScriptBotaoAmareloGravar` à `ImagemBotaoAmareloGravar` durante a criação da planilha.

## Funcionalidades

*   **Criação Automatizada da Planilha:** Cria automaticamente uma planilha Google Sheets com o nome "EstoqueControle[Ano]", onde `[Ano]` é o ano seguinte ao atual (ex: EstoqueControle2025).

*   **Configuração de Abas:** Cria e formata as seguintes abas:
    *   **DETALHADO:** Contém tabelas de estoque, consumo setorial, consumo mensal, e informações detalhadas sobre lotes.
    *   **ENTRADA:** Permite registrar a entrada de produtos no estoque.
    *   **SAÍDA:** Permite registrar a saída de produtos do estoque.
    *   **Leia-me:** Contém instruções detalhadas sobre como usar a planilha.
    *   **invent:** Lista de inventário com código interno, E-FISCO, produtos e quantidade mínima.
    *   **config:** Configurações gerais da planilha, como o ano atual e a lista de setores.

*   **Formatação Condicional:** Aplica formatação condicional para destacar informações importantes, como status de estoque (OK, Crítico, Em Falta), datas de validade de lotes (Vencido, Próximo ao Vencimento), e divergências de dados.

*   **Validação de Dados:** Implementa validação de dados para garantir a consistência das informações inseridas, como a seleção de produtos a partir de uma lista pré-definida.

*   **Proteção de Abas:** Protege as abas "DETALHADO", "config" e "Leia-me" para evitar edições acidentais, permitindo que apenas o proprietário da planilha as modifique. A aba "invent" é protegida com exceção da lista de itens. As abas "ENTRADA" e "SAÍDA" protegem as células com fórmulas.

*   **Botões de Gravação:** Adiciona botões (imagens) nas abas "ENTRADA" e "SAÍDA" para facilitar a gravação dos dados e a execução dos scripts correspondentes.

*   **Fórmulas:** Utiliza diversas fórmulas para calcular automaticamente o estoque atual, gerar relatórios de consumo, e identificar necessidades de reabastecimento.

## Como Usar

### Pré-requisitos:

*   Uma conta Google.
*   Acesso ao Google Sheets.
*   Permissões para executar scripts no Google Apps Script.
*   

### Etapas de Instalação e Configuração:

1.  **Criar a Planilha (Proprietário):**

    *   Ao final de cada ano, acesse o repositório do GitHub: [PlanilhaEstoqueControle](https://github.com/FernandaPaulaDeCarvalho/PlanilhaEstoqueControle).
    *   Procure os arquivos `ScriptPlanilhaEstoqueControle` e `ScriptBotaoAmareloGravar` com a versão mais recente.
    *   Copie o código do arquivo `ScriptPlanilhaEstoqueControle`.
    *   Cole o código copiado no editor do Google App Script da sua conta Google (Novo > Mais > Script do Google Apps).
    *   Execute as funções `primeiraEtapa`, `segundaEtapa` e `terceiraEtapa` **em sequência**.
        *   **ATENÇÃO:** A execução pela primeira vez requer conceder permissões ao script.
        *   **ATENÇÃO:** Aguarde a mensagem de conclusão no painel "Registro de execução" para executar a etapa seguinte.

    *   Ao concluir a terceira etapa, abra o editor de script da planilha recém-criada (Extensões > App Script).
    *   Copie o código do arquivo `ScriptBotaoAmareloGravar` (disponível no repositório) e cole no editor de script da planilha recém-criada.
        *   **ATENÇÃO:** Este script é responsável pela funcionalidade dos botões "GRAVAR" nas abas "ENTRADA" e "SAÍDA". Ele deve ser adicionado separadamente, após a execução das etapas do `ScriptPlanilhaEstoqueControle`.

2.  **Configuração Inicial da Planilha (Proprietário):**

    *   Vá para a aba `INVENT` da planilha do ano anterior e copie a lista de inventário (`invent!A3:D`).
    *   Vá para a aba `INVENT` da nova planilha e cole somente os valores (Ctrl+Shift+V).
    *   Vá para a aba `CONFIG` da planilha do ano anterior e copie a lista de estoque (`config!C3:G`).
    *   Vá para a aba `ENTRADA` da nova planilha, cole somente os valores (Ctrl+Shift+V) e clique no botão GRAVAR.
        *   **ATENÇÃO:** O `ScriptPlanilhaEstoqueControle` clona (cria automaticamente) uma planilha para o ano seguinte. Contudo, é necessário copiar manualmente o inventário e os produtos do estoque do ano anterior e colar na nova planilha.

    *   Vá para a aba `CONFIG` da nova planilha.
    *   Confira o ano na coluna "ANO ATUAL".
    *   Digite os setores na coluna "SETORES".
        *   **ATENÇÃO:** Utilize uma linha para cada setor, até o máximo de 10 setores.

    *   Na aba `INVENT`, digite o código interno, E-FISCO, produto e quantidade mínima para novos itens que entrarão no inventário.
        *   **ATENÇÃO:** Nunca cadastre no inventário produto com "espaço" e/ou "arroba" (@).
        *   **ATENÇÃO:** Verifique periodicamente o banco de dados para incluir novos produtos, fazer correções e eliminar itens redundantes.

    *   Oculte a aba `CONFIG` (opcional).
    *   Compartilhe a planilha com outros usuários (opcional).
        *   **ATENÇÃO:** As fórmulas estão desprotegidas para o proprietário e protegidas para outros usuários.

3.  **Entrada de Produtos:**

    *   Verifique o item desejado na aba `DETALHADO`.
    *   Abra a aba `ENTRADA`.
    *   Na coluna "PRODUTO", digite as primeiras letras do produto e selecione o item desejado no menu suspenso.
    *   Preencha os campos obrigatórios: fabricante, lote, data de validade e quantidade. Se necessário, digite alguma observação.
        *   **ATENÇÃO:** A data de validade deve ser escrita no formato `dd/mm/aaaa`.
        *   **ATENÇÃO:** Os campos que não tiverem dados devem ser preenchidos com ponto (`.`).
        *   **ATENÇÃO:** Nunca cadastre produto usando "arroba" (`@`).

    *   Note que, enquanto estiver preenchendo os campos, o status fica vermelho "PENDENTE".
    *   Quando concluir o preenchimento, o status torna-se amarelo `<<GRAVAR>>`.
    *   Em seguida, clique no botão amarelo `GRAVAR` para executar o script (que você colou no editor de script da planilha).
    *   Confirme a mensagem "Deseja retirar fórmulas e gravar ENTRADA?".
    *   Não mexa, nem clicar em nada enquanto o script estiver em execução.
    *   Se o script for encerrado por algum motivo, ou aparecer o alerta "tempo de execução excedido", clique novamente no botão `GRAVAR`.
    *   Aguarde a mensagem de conclusão "ENTRADA concluída!".
    *   Note que as entradas que estavam amarelo `<<GRAVAR>>` tornam-se verdes "INSERIDO".
        *   **ATENÇÃO:** Após gravar, ainda é permitido fazer alterações. Basta deletar o dado incorreto e digitar novamente.
        *   **ATENÇÃO:** Quando o status torna-se roxo "erro", indica que o produto foi alterado na aba `INVENT`. Corrija o produto na aba `ENTRADA` para que fique igual ao inventário.

4.  **Retirada de Produtos:**

    *   Abra a aba `SAÍDA`.
    *   Na coluna "LOTE @ VALIDADE @ PRODUTO @ FABRICANTE", digite o lote e selecione o item desejado no menu suspenso.
        *   **ATENÇÃO:** Fique atento(a) às especificações e à unidade de medida do produto.
    *   Digite o setor e a quantidade. Se necessário, digite alguma observação.
    *   Note que, enquanto estiver preenchendo os campos, o status fica vermelho "PENDENTE".
    *   Quando preencher corretamente, o status torna-se amarelo `<<GRAVAR>>`.
    *   Ao concluir o preenchimento da retirada, clique no botão amarelo `GRAVAR` para executar o script (que você colou no editor de script da planilha).
    *   Confirme a mensagem "Deseja retirar fórmulas e gravar SAÍDA?".
    *   Não mexa, nem clique em nada enquanto o script estiver em execução.
    *   Se o script for encerrado por algum motivo, ou aparecer o alerta "tempo de execução excedido", clique novamente no botão `GRAVAR`.
    *   Aguarde a mensagem de conclusão "SAÍDA concluída!".
    *   Note que as entradas que estavam amarelo `<<GRAVAR>>` tornam-se verdes "RETIRADO".
        *   **ATENÇÃO:** A retirada de produtos é por lote! Caso deseje retirar dois lotes diferentes, será necessário fazer duas retiradas.
        *   **ATENÇÃO:** Após enviar, ainda é permitido fazer alterações. Basta deletar o dado incorreto e digitar novamente.
        *   **ATENÇÃO:** Quando o status torna-se roxo "erro", indica que o produto foi alterado na aba `ENTRADA`. Corrija o produto na aba `SAÍDA` para que fique igual à entrada.

5.  **Relatórios do Estoque:**

    *   Abra a aba `DETALHADO`.
    *   Confira o status em "ESTOQUE DE PRODUTOS":
        *   `Quant.Atual >= Quant.Mín` o status é verde "OK".
        *   `Quant.Atual < Quant.Mín` o status é amarelo "CRÍTICO".
        *   `Quant.Atual = Zero & Quant.Mín > Zero` o status é vermelho "EM FALTA".
        *   `Quant.Atual = Zero & Quant.Mín = Zero` o status é verde "OK".
        *   `Quant.Atual > Zero & Quant.Mín = Zero` o status é azul "FAZER DOAÇÃO".
        *   `Quant.Atual < Zero` o status é laranja "FAZER REVISÃO", pois não pode haver quantidades negativas no estoque.

    *   Confira periodicamente as tabelas "CONSUMO SETORIAL", "CONSUMO MENSAL", "ABASTECIMENTO MENSAL" e "LOTES".

## Limitações

*   **Dados Manuais:** Requer a cópia manual dos dados do inventário e do estoque inicial da planilha do ano anterior para a nova planilha.
*   **Validade Anual:** A planilha é projetada para ser usada por um ano, sendo necessário criar uma nova planilha ao final de cada ano.
*   **Dependência do Google Apps Script:** Os scripts dependem do Google Apps Script e podem parar de funcionar se houver mudanças na plataforma.
*   **Formulas Complexas:** A modificação das fórmulas requer conhecimento em Google Sheets e Google Apps Script.
*   **Listas Fixas:** A quantidade de setores é limitada a 10.
*   **Disponibilidade do GitHub:** O acesso aos scripts e a este README depende da disponibilidade do GitHub. Interrupções no serviço do GitHub podem impedir temporariamente a instalação ou atualização da planilha.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para:

*   Reportar problemas e sugerir melhorias através das issues do GitHub.
*   Enviar pull requests com suas correções e novas funcionalidades.

## Licença

Este projeto está licenciado sob a [Licença MIT](https://opensource.org/license/mit/). Consulte o arquivo `LICENSE` para obter mais informações sobre os termos da licença.

## Autor

Fernanda Paula de Carvalho [Lattes](https://lattes.cnpq.br/3840867034129036) - [LinkedIn](https://www.linkedin.com/in/fernanda-paula-de-carvalho-phd-37599134) - [GitHub](https://github.com/FernandaPaulaDeCarvalho)
