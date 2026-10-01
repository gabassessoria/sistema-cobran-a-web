# Sistema de Controle de Demandas — Web

Versão web da aplicação desktop de controle de demandas.

## O que ela faz

- recebe a planilha Excel pelo navegador;
- carrega dinamicamente Andamentos e Interessados;
- gera o relatório filtrado em `.xlsx`;
- gera o acompanhamento semanal por Superintendência;
- cria SA, SD, SF, SH, SR e Conselhos Diretores (`CD-DIRETOR-N`);
- preserva a Área Técnica completa;
- cria `FORA_DA_DISTRIBUICAO.xlsx`;
- entrega os arquivos de acompanhamento em um único `.zip`.

### Destaques de prazo no acompanhamento

**PRIORITÁRIOS**
- vencido: vermelho;
- a vencer em até 5 dias: amarelo;
- prazo administrativo: cinza.

**NÃO PRIORITÁRIOS**
- prazo administrativo: cinza;
- demais prazos: sem destaque.

## Rodar localmente

No terminal, dentro da pasta do projeto:

```bash
python -m venv .venv
```

No Windows:

```bash
.venv\Scripts\activate
```

Instale:

```bash
pip install -r requirements.txt
```

Rode:

```bash
uvicorn api.index:app --reload
```

Acesse:

```text
http://127.0.0.1:8000
```

## Publicar no Vercel

1. Coloque esta pasta em um repositório GitHub.
2. No Vercel, crie um novo projeto a partir do repositório.
3. Faça o deploy.
4. Não é necessário configurar banco de dados nesta primeira versão.

A aplicação é stateless: a planilha enviada é processada em memória e não é salva de forma permanente no servidor.

## SharePoint

Nesta primeira versão web, o resultado é baixado pelo navegador.

Salvar diretamente no SharePoint é uma etapa separada, porque uma aplicação hospedada no Vercel não enxerga a pasta OneDrive sincronizada do computador. Para gravação direta no SharePoint será necessário integrar com Microsoft Graph/SharePoint com autenticação e permissões da organização.
