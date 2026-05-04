# FinanceOne

Aplicativo responsivo para **controle financeiro pessoal completo**, com experiência moderna para **web e mobile**.

## Funcionalidades implementadas nesta versão

- Dashboard com KPIs de saldo, receitas, despesas e economia mensal.
- Cadastro de movimentações (receitas/despesas) com atualização em tempo real.
- Simulador de parcelamento inteligente (valor da parcela + cronograma de vencimentos).
- Metas financeiras com barra de progresso.
- Navegação por módulos: Dashboard, Movimentações, Parcelamentos, Metas e Relatórios.
- Tema escuro/claro.
- Layout responsivo com adaptação para telas menores.

## Executar localmente

Como é uma SPA estática, basta abrir o `index.html` no navegador.

Opcional com servidor local:

```bash
python3 -m http.server 8080
```

Depois acesse `http://localhost:8080`.

## Estrutura do projeto

- `index.html`: estrutura da aplicação e seções principais.
- `styles.css`: design system, responsividade e temas light/dark.
- `app.js`: estado da aplicação, renderização dos módulos e regras de negócio do frontend.

## Próximos passos (evolução full stack)

- Backend Node.js (NestJS) com API REST e autenticação JWT/OAuth Google.
- PostgreSQL para persistência de transações, cartões, contas e metas.
- Exportação de relatórios (PDF, Excel, CSV).
- Alertas (vencimentos, limite de cartão e gasto excessivo).
- IA para previsão de gastos e sugestões automáticas de economia.
