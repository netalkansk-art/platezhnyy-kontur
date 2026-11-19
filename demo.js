(() => {
  const storageKey = 'platezhnyy-kontur-public-demo-v1';
  const companyNames = ['ООО «Северный контур»', 'ООО «Практика Плюс»'];
  const initial = [
    [1, 0, 'Городские сети', 18750, 'Интернет и телефония за ноябрь', 24, 'review', null],
    [2, 1, 'Линия офиса', 36400, 'Бумага и картриджи', 25, 'review', null],
    [3, 0, 'Квартал', 128000, 'Аренда офиса за декабрь', 27, 'approval', 'Аренда'],
    [4, 1, 'Счёт и право', 52000, 'Сопровождение договора', 23, 'approval', 'Услуги'],
    [5, 0, 'Папирус', 9450, 'Расходные материалы', 20, 'approved', 'Материалы'],
  ].map(([id, company, name, amount, purpose, day, stage, article]) => ({
    id, company: companyNames[company], recipient: 'ООО «' + name + '»', amount, purpose,
    due: '2026-11-' + day, stage, article, step: Number(stage === 'approved'), history: [],
    department: ['Администрация','Операции','ИТ'][(id - 1) % 3],
  }));
  function snapshot() {
    const stored = localStorage.getItem(storageKey);
    return stored ? JSON.parse(stored) : {revision:0, requests:initial};
  }
  async function exportWorkbook(selected) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Реестр заявок');
    sheet.columns = [
      {header:'Номер', width:10}, {header:'Заявка', width:36}, {header:'Компания', width:30},
      {header:'Подразделение', width:20}, {header:'Получатель', width:30},
      {header:'Сумма, ₽', width:16}, {header:'Срок', width:16}, {header:'Статус', width:24},
    ];
    sheet.getColumn(6).numFmt = '#,##0.00';
    sheet.getColumn(7).numFmt = 'dd.mm.yyyy';
    sheet.views = [{state:'frozen', ySplit:1}];
    sheet.autoFilter = 'A1:H1';
    for (const request of selected) {
      const status = request.returned ? 'На исправлении' : stages.find(([id]) => id === request.stage)[1];
      sheet.addRow([request.id, request.purpose, request.company, request.department,
        request.recipient, request.amount, new Date(request.due + 'T00:00:00Z'), status]);
    }
    return new Response(await workbook.xlsx.writeBuffer(), {
      headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},
    });
  }
  window.publicDemoRequest = async (path, options = {}) => {
    const state = snapshot();
    if (path === '/api/requests' && !options.method) {
      return Response.json(state);
    }
    const body = JSON.parse(options.body);
    if (body.revision !== state.revision) return new Response(null,{status:409});
    if (path === '/api/requests' && options.method === 'PUT') {
      const next = {revision:state.revision + 1, requests:body.requests};
      localStorage.setItem(storageKey, JSON.stringify(next));
      return Response.json(next);
    }
    if (path === '/api/registry.xlsx' && options.method === 'POST') {
      const ids = new Set(body.ids);
      return exportWorkbook(state.requests.filter(request => ids.has(request.id)));
    }
    return new Response(null,{status:404});
  };
})();
