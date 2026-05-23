# ⚖️ Extrator e Simplificador de Pautas de Julgamento

Uma aplicação web *client-side* (Front-end) desenvolvida para automatizar a limpeza e formatação de pautas de julgamento, otimizando a rotina de trabalho em gabinetes de magistrados.

---

## 📖 Contexto e Propósito

Na rotina do gabinete de uma Desembargadora do Trabalho, o recebimento da pauta de julgamento em formato DOCX gera uma demanda considerável de trabalho manual. O documento original contém o cabeçalho da sessão, orientações institucionais e, sob o tópico **"FEITOS DE COMPETÊNCIA RECURSAL"**, uma extensa lista de processos (frequentemente ultrapassando uma centena de itens). 

Para cada processo, o sistema processual lista uma série de informações: *Desembargador(a) Relator(a), Recorrentes, Recorridos, Advogados, Peritos, Terceiros, Custos Legis*, entre outros. No entanto, para a condução dinâmica da sessão, a Desembargadora necessita apenas de um roteiro limpo e objetivo, contendo **exclusivamente o número do processo e o nome do Relator(a)**.

**O Problema:** A exclusão manual das dezenas de linhas excedentes em cada um dos mais de cem processos consumia tempo valioso da equipe, além de estar sujeita a erros e poluição visual devido a quebras de linha irregulares e nomenclaturas no plural.

**A Solução:** Esta aplicação foi criada para automatizar 100% dessa tarefa. Hospedada diretamente no GitHub Pages e rodando inteiramente no navegador do usuário (sem envio de dados sensíveis para servidores externos), a ferramenta varre o documento identificando padrões estruturais. Utilizando uma lógica de "blocos de exclusão", ela ignora cabeçalhos, descarta partes envolvidas, elimina quebras de linha residuais e entrega, em segundos, um documento perfeitamente formatado para leitura e impressão, com separações sutis entre cada feito.

---

## ✨ Funcionalidades

- **Processamento Inteligente de DOCX:** Lê e manipula o código XML por trás de arquivos Word (`.docx`) diretamente no navegador.
- **Lógica de Blocos (Anti-Falsos Positivos):** Substitui expressões regulares simples por uma máquina de estados. O sistema identifica o Processo e o Relator e apaga cirurgicamente todo o conteúdo excedente até o próximo processo, lidando perfeitamente com plurais e textos multilinhas.
- **Limpeza de Cabeçalho:** Remoção automática de todo o texto institucional que antecede o início da lista de processos.
- **Injeção de Estilo Nativo (Word):** Insere divisórias horizontais em cinza claro (`#D3D3D3`) diretamente no código do DOCX, separando os processos com elegância.
- **Exportação Dupla:**
  - **Download DOCX:** O arquivo original resumido e reformatado.
  - **Visualização/Impressão HTML:** Geração de um layout em formato de *cards*, sem quebras de página no meio do processo (`page-break-inside: avoid`), estilizado em tons de cinza e otimizado para impressão em folha A4.
- **Interface Intuitiva:** Suporte a *Drag and Drop* (arrastar e soltar), com feedbacks visuais claros e iconografia moderna.

---

## 🛠️ Tecnologias Utilizadas

O projeto foi construído focando em leveza, não dependendo de *frameworks* pesados:

- **HTML5 & CSS3:** Semântica e estilização baseada em variáveis de ambiente corporativo.
- **JavaScript (Vanilla ES6+):** Lógica assíncrona, manipulação de DOM e parsing de XML.
- **[JSZip](https://stuk.github.io/jszip/):** Biblioteca para descompactar e reempacotar arquivos `.docx` no lado do cliente.
- **[Lucide Icons](https://lucide.dev/):** Pacote de ícones SVG leves e minimalistas (sem uso de emojis na interface).

---

## 🚀 Como Utilizar (Usuário Final)

1. Acesse a página da aplicação hospedada no GitHub Pages.
2. Arraste o arquivo `.docx` original da pauta para a área tracejada (ou clique para buscar no computador).
3. Aguarde o aviso de **"Pauta processada com sucesso!"** (leva apenas alguns segundos).
4. Escolha sua opção de exportação:
   - Clique em **"Salvar DOCX Resumido"** para baixar o arquivo Word modificado.
   - Clique em **"Visualizar/Imprimir Pauta"** para gerar a versão em tela, pronta para ser impressa ou salva em PDF via atalho do navegador (`Ctrl+P` / `Cmd+P`).
5. Caso deseje limpar outro arquivo, clique em **"Processar outra pauta"**.

---

## 💻 Como Executar e Modificar o Projeto (Desenvolvedores)

Como se trata de uma aplicação estática (apenas Front-end), não é necessário instalar Node.js, dependências pesadas ou bancos de dados.

1. **Clone o repositório:**
   ```bash
   git clone https://github.com/seu-usuario/nome-do-repositorio.git
   ```
2. **Abra o diretório:**
   ```bash
   cd nome-do-repositorio
   ```
3. **Execute localmente:**
   Basta abrir o arquivo `index.html` diretamente em qualquer navegador moderno moderno (Chrome, Edge, Firefox, Safari).
   *Dica: Para testes avançados, utilize a extensão "Live Server" do VS Code.*

---

## 📂 Estrutura de Arquivos

```text
├── index.html     # Estrutura principal da página e importação dos scripts
├── style.css      # Estilos visuais, responsividade e regras de impressão (@media print)
├── app.js         # Lógica central: manipulação do ZIP/XML, filtragem de dados e geração de arquivos
└── README.md      # Documentação do projeto
```

---

## 🗺️ Roadmap (Melhorias Futuras)

- [ ] **Web Workers:** Mover o processamento do DOCX para uma *thread* separada em segundo plano, evitando qualquer travamento visual da interface em pautas excepcionalmente longas.
- [ ] **Filtro Reverso por Relator(a):** Adicionar a opção de extrair e isolar apenas os processos de responsabilidade de um magistrado específico.
- [ ] **Estatísticas Rápidas:** Exibir na tela de sucesso um resumo quantitativo (ex: "Foram localizados e extraídos 146 processos").

---

**Desenvolvido para otimização de rotinas no Poder Judiciário Trabalhista.** ⚖️
```
