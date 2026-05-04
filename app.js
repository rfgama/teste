const state = {
  transactions: [
    { data: '2026-05-01', descricao: 'Salário', tipo: 'Receita', categoria: 'Salário', valor: 8200 },
    { data: '2026-05-02', descricao: 'Mercado', tipo: 'Despesa', categoria: 'Alimentação', valor: 640 },
    { data: '2026-05-03', descricao: 'Uber', tipo: 'Despesa', categoria: 'Transporte', valor: 58 }
  ],
  goals: [
    { nome: 'Reserva de emergência', atual: 12500, alvo: 30000 },
    { nome: 'Viagem', atual: 3500, alvo: 12000 }
  ]
};

const views = document.querySelectorAll('.view');
const navLinks = document.querySelectorAll('.nav-link');
const viewTitle = document.getElementById('view-title');

const currency = v => v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

function renderDashboard() {
  const receitas = state.transactions.filter(t => t.tipo === 'Receita').reduce((a, t) => a + t.valor, 0);
  const despesas = state.transactions.filter(t => t.tipo === 'Despesa').reduce((a, t) => a + t.valor, 0);
  const saldo = receitas - despesas;
  const cards = [
    ['Saldo atual', currency(saldo)], ['Receitas', currency(receitas)], ['Despesas', currency(despesas)],
    ['Contas pendentes', '3'], ['Faturas futuras', currency(1420)], ['Economia do mês', `${Math.max(0, (saldo / Math.max(receitas, 1)) * 100).toFixed(1)}%`]
  ];
  document.getElementById('cards').innerHTML = cards.map(([k, v]) => `<article class="card"><h4>${k}</h4><strong>${v}</strong></article>`).join('');

  const byCategory = {};
  state.transactions.filter(t => t.tipo === 'Despesa').forEach(t => byCategory[t.categoria] = (byCategory[t.categoria] || 0) + t.valor);
  const max = Math.max(...Object.values(byCategory), 1);
  document.getElementById('category-bars').innerHTML = Object.entries(byCategory)
    .map(([cat, val]) => `<div><small>${cat} - ${currency(val)}</small><div class="bar"><span style="width:${(val / max) * 100}%"></span></div></div>`).join('');

  const topCat = Object.entries(byCategory).sort((a,b)=>b[1]-a[1])[0]?.[0] || 'N/A';
  document.getElementById('insights').innerHTML = `
    <li>Maior gasto: <strong>${topCat}</strong></li>
    <li>Média de despesas: <strong>${currency(despesas / Math.max(1, state.transactions.filter(t => t.tipo === 'Despesa').length))}</strong></li>
    <li>Tendência: <strong>${saldo > 0 ? 'superávit' : 'atenção ao orçamento'}</strong></li>`;
}

function renderTransactions() {
  document.getElementById('transaction-table').innerHTML = state.transactions.map(t =>
    `<tr><td>${t.data}</td><td>${t.descricao}</td><td>${t.tipo}</td><td>${t.categoria}</td><td>${currency(t.valor)}</td></tr>`).join('');
}

function renderGoals() {
  document.getElementById('goals').innerHTML = state.goals.map(g => {
    const pct = Math.min(100, (g.atual / g.alvo) * 100).toFixed(1);
    return `<div class="goal"><strong>${g.nome}</strong><p>${currency(g.atual)} de ${currency(g.alvo)} (${pct}%)</p><div class="bar"><span style="width:${pct}%"></span></div></div>`;
  }).join('');
}

document.getElementById('transaction-form').addEventListener('submit', (e) => {
  e.preventDefault();
  const form = new FormData(e.target);
  state.transactions.unshift({
    descricao: form.get('descricao'), valor: Number(form.get('valor')), tipo: form.get('tipo'), categoria: form.get('categoria'), data: form.get('data')
  });
  e.target.reset();
  renderTransactions();
  renderDashboard();
});

document.getElementById('installment-form').addEventListener('submit', (e) => {
  e.preventDefault();
  const form = new FormData(e.target);
  const total = Number(form.get('total'));
  const count = Number(form.get('count'));
  const installment = total / count;
  const date = new Date(form.get('firstDue'));
  let html = `<p><strong>Parcela:</strong> ${currency(installment)} (${count}x)</p><ol>`;
  for (let i = 0; i < count; i++) {
    const due = new Date(date); due.setMonth(date.getMonth() + i);
    html += `<li>${due.toLocaleDateString('pt-BR')} - ${currency(installment)}</li>`;
  }
  html += '</ol>';
  document.getElementById('installment-result').innerHTML = html;
});

navLinks.forEach(btn => btn.addEventListener('click', () => {
  navLinks.forEach(b => b.classList.remove('active')); btn.classList.add('active');
  views.forEach(v => v.classList.remove('active'));
  const id = btn.dataset.view; document.getElementById(id).classList.add('active');
  viewTitle.textContent = btn.textContent;
}));

document.getElementById('theme-toggle').addEventListener('click', () => document.body.classList.toggle('light'));

renderTransactions();
renderDashboard();
renderGoals();
