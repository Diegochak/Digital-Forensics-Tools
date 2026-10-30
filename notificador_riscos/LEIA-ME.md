# Notificador de software não homologado

Lê duas planilhas Excel e monta um e-mail por achado, com base no modelo Word da criticidade. Cria os e-mails como **rascunhos no Outlook clássico**. Nada é enviado automaticamente.

Só os achados **EMERGENCY, CRITICAL e HIGH** viram e-mail ao gestor (`criticidade › notificar` no config). Os demais níveis continuam na planilha, para controle e dashboards. Eles aparecem na prévia como "fica só no controle".

## Instalação (uma vez)

1. Instale o Python 3.10 ou superior pelo python.org e marque "Add python.exe to PATH".
2. Dê dois cliques em `instalar.bat`.

## Uso

| Passo | Arquivo | O que faz |
|---|---|---|
| 1 | `1_mapear_colunas.bat` | Lista letra, cabeçalho e um exemplo de cada coluna das duas planilhas |
| 2 | `2_gerar_previa.bat` | Gera os e-mails em HTML e abre um índice no navegador. Não usa o Outlook |
| 3 | `3_criar_rascunhos.bat` | Pede confirmação e cria os rascunhos em **Rascunhos › Notificações GRC** |

Opções de linha de comando:

- `--limite 3`: gera só os 3 primeiros (bom para o primeiro teste)
- `--forcar`: recria rascunhos que já constam no log
- `--sim`: não pede confirmação
- `--config outro.yaml`: usa outra configuração

## Como apontar as suas planilhas

Tudo fica no `config.yaml`:

- **`planilha_analise`** é a planilha 2. Cada linha é um achado e gera um e-mail.
- **`planilha_inventario`** é a planilha 1. É opcional e completa o achado. A ligação entre as duas é feita por `ligacao` (por padrão, endpoint + nome do software).
- **`colunas`** usa o formato `nome_do_campo: coluna`. A coluna pode ser a letra (`"E"`) ou o cabeçalho (`"Risk Level"`). Indicar pelo cabeçalho resiste a colunas inseridas no meio.
- **`aba`** aceita um nome ou uma lista (`["Week 43", "Week 44"]`).
- **`linha_cabecalho` / `primeira_linha_dados` / `ultima_linha_dados`** delimitam a área lida.

## Como preparar os seus modelos Word

1. Copie o seu modelo de cada criticidade para a pasta `modelos/`.
2. No texto, troque os dados variáveis por `{{ nome_do_campo }}`, usando os nomes definidos em `colunas`. Por exemplo: `Prezado(a) {{ gestor }}`, ou `O software {{ aplicacao }} foi encontrado no endpoint {{ endpoint }}`.
3. Aponte cada arquivo em `criticidade › modelos`.

Campos extras disponíveis em todos os modelos:

- `{{ criticidade }}`: nível padronizado
- `{{ criticidade_original }}`: valor como está na planilha
- `{{ hoje }}`
- `{{ linha }}`

Para usar um valor alternativo quando a célula estiver vazia: `{{ alternativa | default("a definir", true) }}`

**Sobre a formatação:** a conversão do `.docx` preserva negrito, itálico, tabelas, listas, links e imagens (como o logo). Cores de fonte e fontes específicas se perdem. Se o visual exato importa, abra o modelo no Word e use **Salvar como › Página da Web, Filtrada (.htm)**. Depois aponte o `.htm` no config, que mantém tudo.

**Se um marcador não for preenchido:** apague o marcador inteiro e digite-o de novo de uma vez só. O Word às vezes quebra o texto em pedaços com formatações diferentes.

## Ressalvas

Se um achado notificável estiver sem endpoint, sem usuário ou sem gestor (nome ou e-mail), ele não gera e-mail. Vira **RESSALVA** e aparece em `saida/ressalvas.csv`, que é refeito a cada execução com o que ainda está pendente. Quando o dado for completado na planilha, a próxima execução cria o rascunho normalmente.

Os campos exigidos ficam em `ressalva › campos_obrigatorios`, no config. Os `filtros` são aplicados antes: uma linha excluída por filtro não vira ressalva.

## Controles

- **Prévia antes de tudo.** O índice mostra quem recebe, o assunto e os alertas de cada linha:
  - destinatário sem e-mail
  - achado sem correspondência no inventário
  - criticidade sem modelo
- **Sem duplicidade.** O log `saida/log_notificacoes.csv` registra cada rascunho pela chave (`campo_chave`, por padrão o ticket). Uma nova execução pula o que já foi criado. Na mesma execução, uma chave repetida (por exemplo, o mesmo ticket em duas abas) gera um e-mail só. As outras linhas aparecem como "chave repetida".
- **Evidência.** O mesmo log guarda data, destinatário, assunto e status. Serve de trilha para auditoria.
- **Destinatário sem "@"** é resolvido pelo catálogo do Outlook. Se não resolver, o rascunho é criado com status `RASCUNHO_COM_ALERTA`.
- **Uma linha com erro** não interrompe o lote.
- **Fórmulas:** a leitura usa o último valor calculado. Salve a planilha no Excel antes de rodar.

## Estrutura

```
notificador.py   fluxo principal e linha de comando
planilhas.py     leitura das planilhas e mapeamento de colunas
modelos.py       modelos Word/HTML e preenchimento dos campos
remetentes.py    prévia HTML e rascunhos no Outlook
registro.py      log CSV
```

## Migração para o Google

Só `remetentes.py` conhece o Outlook. Para o Gmail, cria-se uma classe `RemetenteGmail` com o mesmo método `entregar(email)`, que gera rascunhos pela Gmail API. Leitura, modelos, filtros e log continuam iguais.
