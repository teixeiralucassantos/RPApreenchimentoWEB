# Preenchimento Automático de Formulários da Web

## Descrição do Projeto

Este projeto tem como objetivo automatizar o preenchimento de formulários online utilizando Python. Ele utiliza a biblioteca Tkinter para criar uma interface gráfica do usuário (GUI) e a biblioteca Selenium para interagir com páginas da web. O usuário pode adicionar dados a um formulário, que são lidos de um arquivo Excel, e preenchidos automaticamente em um formulário da web específico.

## Tecnologias Utilizadas

- **Python:** Linguagem de programação principal utilizada para desenvolver a aplicação.
- **Tkinter:** Biblioteca padrão do Python para criar interfaces gráficas.
- **Selenium:** Ferramenta para automação de navegadores, permitindo interações programáticas com páginas da web.
- **Pandas:** Biblioteca para manipulação e análise de dados, utilizada para ler os dados do Excel.
- **WebDriver Manager:** Facilita o gerenciamento do driver do Chrome para o Selenium.

## Funcionalidades

1. **Interface Gráfica:** O usuário pode visualizar e interagir com uma tabela que mostra os dados a serem preenchidos.
2. **Leitura de Dados:** Os dados são lidos de um arquivo Excel, permitindo que o usuário preencha rapidamente o formulário com informações previamente armazenadas.
3. **Preenchimento em Massa:** O aplicativo pode preencher automaticamente o formulário da web com todos os dados carregados na tabela.
4. **Adição, Alteração e Exclusão de Dados:** O usuário pode manipular os dados na tabela antes de enviar os dados para o formulário.
5. **Preenchimento Individual:** O usuário pode preencher o formulário uma única vez com os dados inseridos na interface.

## Técnicas de RPA Aplicadas

- **RPA (Robotic Process Automation):** Este projeto é um exemplo prático de RPA, onde tarefas manuais de preenchimento de formulários são automatizadas. Isso ajuda a reduzir o erro humano, aumentar a eficiência e liberar o tempo dos usuários para tarefas mais estratégicas.
- **Automação de Navegadores:** O uso do Selenium permite simular ações humanas, como cliques e digitação, para interagir com o formulário da web de maneira programática.
- **Integração de Dados:** A combinação do Pandas para manipulação de dados e do Selenium para automação de navegação demonstra uma integração eficaz entre diferentes tecnologias.

## Estrutura do Código

O código é organizado em uma classe `FormPreenchimento`, que encapsula todas as funcionalidades necessárias. Os principais métodos incluem:

- `configurar_janela()`: Configura a janela principal da aplicação.
- `criar_treeview()`: Cria a tabela que exibe os dados.
- `preencher_em_massa()`: Preenche o formulário da web com todos os dados da tabela.
- `preencher_formulario_web(dados)`: Interage com o navegador para preencher um formulário específico.
