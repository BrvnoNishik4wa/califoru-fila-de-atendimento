# CaliforU — Sistema de Lista da Vez

![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=flat-square&logo=googleappsscript&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=flat-square&logo=javascript&logoColor=black)
![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=flat-square&logo=html5&logoColor=white)
![CSS3](https://img.shields.io/badge/CSS3-1572B6?style=flat-square&logo=css3&logoColor=white)
![Google Sheets](https://img.shields.io/badge/Google_Sheets-34A853?style=flat-square&logo=googlesheets&logoColor=white)

> Sistema de controle de fila de atendimento (lista da vez) para as lojas de shopping da CaliforU, com registro automático de resultado de venda e relatórios de conversão por vendedor.

## O problema

Nas lojas de shopping, a ordem de quem atende o próximo cliente era controlada de um jeito manual e redundante: cada vendedora assinava numa prancheta de papel ao final do atendimento pra marcar sua posição na fila, e o mesmo registro era lançado de novo numa planilha do Excel — além de precisar imprimir a prancheta com frequência. Era trabalho duplicado por vendedora, por atendimento, todos os dias.

Além do retrabalho, esse processo não gerava nada além do registro em si: para saber a taxa de conversão de cada vendedora, ou os motivos mais comuns de perda de venda, alguém precisava tabular tudo manualmente depois.

## O que o sistema faz

- **Fila digital em tempo real**, por loja — a vendedora entra na fila com um toque, sem papel.
- **Regras de ausência justas**: uma pausa pro banheiro mantém a posição exata na fila; pausa pessoal ou intervalo manda a vendedora pro fim da fila. A regra fica explícita no código, refletindo o critério de justiça que o time definiu.
- **Reset automático da fila à meia-noite** — corrige um problema real do processo manual, em que a fila do dia anterior podia ser confundida com a de hoje.
- **Registro do resultado do atendimento** (venda, perda ou troca/ajuste), com motivo da perda e um campo de texto livre para o produto específico que o cliente procurava quando aplicável — funcionalidade que saiu direto de uma conversa com o time da loja.
- **Relatório completo em um clique**: gera uma planilha à parte com resumo geral, ranking de vendedoras por taxa de conversão, detalhamento por vendedora e agregação dos motivos de perda mais comuns — o que antes exigia tabular tudo à mão.
- **Cadastro de vendedoras pela própria administradora**, direto pela tela do sistema, sem precisar editar planilha ou código.

## Desenvolvido com quem usa

Antes de fechar o desenho do sistema, apresentei o protótipo para o time das lojas de shopping numa reunião. Algumas funcionalidades — como o campo de produto específico no motivo de perda — vieram direto dessa conversa. O sistema também já foi validado e aprovado pelo diretor/dono da CaliforU e pela administradora responsável pelas lojas antes de entrar em operação.

## Como foi construído

Mesma base do [sistema de ajustes](../califoru-sistema-ajustes): **Google Apps Script** como backend, **Google Sheets** como banco de dados. A arquitetura segue o mesmo padrão de segurança:

- Toda função pública resolve permissão e loja **no servidor**, nunca confiando no que vem do cliente — a mesma lição aplicada de novo, agora em outro sistema.
- **Controle de concorrência** com `LockService` em toda escrita na fila, para evitar que duas ações simultâneas corrompam o estado.
- **Migração de coluna idempotente** — mesma disciplina do sistema de ajustes: segura para rodar mais de uma vez, sem sobrescrever dado existente.
- **Gatilho de tempo automático** (`resetFilasDiario`) rodando à meia-noite sem depender de nenhum usuário logado.

## Stack

`Google Apps Script` · `HTML / CSS / JavaScript` · `Google Sheets` (banco de dados)

## Estrutura do repositório

```
califoru-lista-da-vez/
├── backend/
│   └── Codigo.gs          # lógica de servidor: permissões, fila, relatórios
├── frontend/
│   ├── Index.html         # estrutura da página
│   ├── JavaScript.html    # lógica de front-end
│   └── Styles.html        # estilos
├── docs/
│   └── screenshots/       # capturas de tela e vídeo de demonstração
└── README.md
```

## Status

Sistema aprovado e pronto para entrar em operação nas lojas em setembro de 2026. A expectativa é de 15 atendimentos ou mais por dia — ainda não há uma base de uso real, já que o sistema não está no ar até o momento desta publicação.

## Demonstração

🎥 *Vídeo em breve.*

## Sobre este repositório

Nomes de vendedoras e e-mails internos foram substituídos por valores genéricos antes da publicação. A lógica, a arquitetura e as regras de negócio são exatamente as que vão rodar em produção na CaliforU.
