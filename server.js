const express = require('express');
const path = require('path');
const https = require('https');

const app = express();
const PORT = process.env.PORT || 3000;

const FEISHU = {
  APP_ID: 'cli_a920565c54f8dcba',
  APP_SECRET: 'tb9ieqwsib2ZIRyEhDqKxsmgPmonb8ep',
  BASE_TOKEN: 'XdxKb0YgJa5nM6sT3saciIX6nue',
  TABLE_ID: 'tblYDHF3ObdeMybZ',
};

const LEADERS = [
  { name: '赵总', role: '董事长', color: '#E4393C' },
  { name: '刚总', role: '总经理', color: '#D48806' },
  { name: '许书记', role: '党委副书记', color: '#1677FF' },
  { name: '肖总', role: '副总经理', color: '#389E0D' },
  { name: '孟总', role: '副总经理', color: '#722ED1' },
  { name: 'X书记', role: '纪委书记', color: '#C41D7F' },
  { name: '闫总', role: '副总经理', color: '#13A8A8' },
];

let tokenCache = { token: '', expire: 0 };

function feishuRequest(method, apiPath, data) {
  return new Promise((resolve, reject) => {
    const body = data ? Buffer.from(JSON.stringify(data), 'utf-8') : null;
    const options = {
      hostname: 'open.feishu.cn', path: '/open-apis' + apiPath, method,
      headers: { 'Content-Type': 'application/json; charset=utf-8', 'Authorization': 'Bearer ' + tokenCache.token },
    };
    if (body) options.headers['Content-Length'] = body.length;
    const req = https.request(options, res => {
      let chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => { try { resolve(JSON.parse(Buffer.concat(chunks).toString('utf-8'))); } catch(e) { reject(e); } });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

async function getToken() {
  if (tokenCache.token && Date.now() < tokenCache.expire) return tokenCache.token;
  const r = await feishuRequest('POST', '/auth/v3/tenant_access_token/internal', { app_id: FEISHU.APP_ID, app_secret: FEISHU.APP_SECRET });
  tokenCache = { token: r.tenant_access_token, expire: Date.now() + (r.expire - 300) * 1000 };
  return tokenCache.token;
}

async function feishuAPI(method, apiPath, data) {
  await getToken();
  let r = await feishuRequest(method, apiPath, data);
  if (r.code === 99991663 || r.code === 99991661) { tokenCache.expire = 0; await getToken(); r = await feishuRequest(method, apiPath, data); }
  return r;
}

function parseTime(t) {
  if (!t || t === '全天') return null;
  let m = t.match(/(\d{1,2})[点時](\d{1,2}|半)?/);
  if (m) return parseInt(m[1]) * 60 + (m[2] === '半' ? 30 : (m[2] ? parseInt(m[2]) : 0));
  m = t.match(/(\d{1,2})[：:](\d{2})/);
  if (m) return parseInt(m[1]) * 60 + parseInt(m[2]);
  return null;
}

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// GET schedules
app.get('/api/schedules', async (req, res) => {
  try {
    const basePath = `/bitable/v1/apps/${FEISHU.BASE_TOKEN}/tables/${FEISHU.TABLE_ID}/records`;
    let all = [], pt = '';
    do {
      const r = await feishuAPI('GET', basePath + `?page_size=100${pt ? '&page_token=' + pt : ''}`);
      if (r.code !== 0) return res.status(500).json({ error: r.msg });
      if (r.data.items) all = all.concat(r.data.items);
      pt = r.data.has_more ? r.data.page_token : '';
    } while (pt);

    const schedules = all.map(rec => {
      const f = rec.fields;
      let dateStr = '';
      if (f['日期']) { const d = new Date(f['日期']); dateStr = d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
      const sm = parseTime(f['开始时间']), em = parseTime(f['结束时间']);
      return {
        id: rec.record_id, leader: f['领导姓名']||'', date: dateStr,
        start_time: f['开始时间']||'', end_time: f['结束时间']||'',
        start_min: sm, end_min: em||(sm?sm+90:null),
        title: f['事项']||'', location: f['地点']||'',
        participants: f['参与人']||'', importance: f['重要程度']||'中',
        notes: f['备注']||'', is_allday: f['开始时间']==='全天',
        companion: f['陪同人']||'', preparation: f['筹备内容']||'',
        accommodation: f['食宿安排']||'', vehicle: f['车辆安排']||'',
      };
    }).filter(s => s.date && s.leader);
    schedules.sort((a,b) => a.date!==b.date ? a.date.localeCompare(b.date) : (a.start_min||0)-(b.start_min||0));
    res.json(schedules);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// POST create schedule
app.post('/api/schedules', async (req, res) => {
  try {
    const s = req.body;
    const dateTs = new Date(s.date + 'T00:00:00+08:00').getTime();
    const fields = {
      '事项': s.title, '领导姓名': s.leader, '日期': dateTs,
      '开始时间': s.start_time, '结束时间': s.end_time||'',
      '地点': s.location||'', '参与人': s.participants||'',
      '重要程度': s.importance||'中', '备注': s.notes||'',
      '陪同人': s.companion||'', '筹备内容': s.preparation||'',
      '食宿安排': s.accommodation||'', '车辆安排': s.vehicle||'',
    };
    const r = await feishuAPI('POST', `/bitable/v1/apps/${FEISHU.BASE_TOKEN}/tables/${FEISHU.TABLE_ID}/records`, { fields });
    if (r.code !== 0) return res.status(500).json({ error: r.msg });
    res.json({ success: true, id: r.data.record.record_id });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// PUT update schedule
app.put('/api/schedules/:id', async (req, res) => {
  try {
    const s = req.body;
    const fields = {};
    if (s.title) fields['事项'] = s.title;
    if (s.leader) fields['领导姓名'] = s.leader;
    if (s.date) fields['日期'] = new Date(s.date + 'T00:00:00+08:00').getTime();
    if (s.start_time) fields['开始时间'] = s.start_time;
    if (s.end_time !== undefined) fields['结束时间'] = s.end_time;
    if (s.location !== undefined) fields['地点'] = s.location;
    if (s.participants !== undefined) fields['参与人'] = s.participants;
    if (s.importance) fields['重要程度'] = s.importance;
    if (s.notes !== undefined) fields['备注'] = s.notes;
    if (s.companion !== undefined) fields['陪同人'] = s.companion;
    if (s.preparation !== undefined) fields['筹备内容'] = s.preparation;
    if (s.accommodation !== undefined) fields['食宿安排'] = s.accommodation;
    if (s.vehicle !== undefined) fields['车辆安排'] = s.vehicle;
    const r = await feishuAPI('PUT', `/bitable/v1/apps/${FEISHU.BASE_TOKEN}/tables/${FEISHU.TABLE_ID}/records/${req.params.id}`, { fields });
    if (r.code !== 0) return res.status(500).json({ error: r.msg });
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// DELETE schedule
app.delete('/api/schedules/:id', async (req, res) => {
  try {
    const r = await feishuAPI('DELETE', `/bitable/v1/apps/${FEISHU.BASE_TOKEN}/tables/${FEISHU.TABLE_ID}/records/${req.params.id}`);
    if (r.code !== 0) return res.status(500).json({ error: r.msg });
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/leaders', (req, res) => res.json(LEADERS));
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

app.listen(PORT, '0.0.0.0', () => { console.log(`\n  西影集团领导行程系统: http://localhost:${PORT}\n`); });
