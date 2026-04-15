const express = require('express');
const path = require('path');
const https = require('https');

const app = express();
const PORT = process.env.PORT || 3000;

// ---- 飞书配置 ----
const FEISHU = {
  APP_ID: 'cli_a920565c54f8dcba',
  APP_SECRET: 'tb9ieqwsib2ZIRyEhDqKxsmgPmonb8ep',
  BASE_TOKEN: 'XdxKb0YgJa5nM6sT3saciIX6nue',
  TABLE_ID: 'tblYDHF3ObdeMybZ',
};

// 领导列表（排序）
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

// ---- 飞书 API ----
function feishuRequest(method, apiPath, data) {
  return new Promise((resolve, reject) => {
    const body = data ? Buffer.from(JSON.stringify(data), 'utf-8') : null;
    const options = {
      hostname: 'open.feishu.cn',
      path: '/open-apis' + apiPath,
      method,
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Authorization': 'Bearer ' + tokenCache.token,
      },
    };
    if (body) options.headers['Content-Length'] = body.length;

    const req = https.request(options, (res) => {
      let chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        try { resolve(JSON.parse(Buffer.concat(chunks).toString('utf-8'))); }
        catch (e) { reject(e); }
      });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

async function getToken() {
  if (tokenCache.token && Date.now() < tokenCache.expire) return tokenCache.token;
  const r = await feishuRequest('POST', '/auth/v3/tenant_access_token/internal', {
    app_id: FEISHU.APP_ID, app_secret: FEISHU.APP_SECRET,
  });
  tokenCache = { token: r.tenant_access_token, expire: Date.now() + (r.expire - 300) * 1000 };
  return tokenCache.token;
}

async function feishuAPI(method, apiPath, data) {
  await getToken();
  let r = await feishuRequest(method, apiPath, data);
  if (r.code === 99991663 || r.code === 99991661) {
    tokenCache.expire = 0;
    await getToken();
    r = await feishuRequest(method, apiPath, data);
  }
  return r;
}

// ---- 时间解析工具 ----
function parseTime(t) {
  if (!t || t === '全天') return null;
  let m = t.match(/(\d{1,2})[点時](\d{1,2}|半)?/);
  if (m) {
    const h = parseInt(m[1]);
    const mins = m[2] === '半' ? 30 : (m[2] ? parseInt(m[2]) : 0);
    return h * 60 + mins;
  }
  m = t.match(/(\d{1,2})[：:](\d{2})/);
  if (m) return parseInt(m[1]) * 60 + parseInt(m[2]);
  return null;
}

function minToTime(m) {
  return `${String(Math.floor(m/60)).padStart(2,'0')}:${String(m%60).padStart(2,'0')}`;
}

// ---- 中间件 ----
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ---- API: 获取所有行程 ----
app.get('/api/schedules', async (req, res) => {
  try {
    const basePath = `/bitable/v1/apps/${FEISHU.BASE_TOKEN}/tables/${FEISHU.TABLE_ID}/records`;
    let allRecords = [];
    let pageToken = '';
    do {
      const url = basePath + `?page_size=100${pageToken ? '&page_token=' + pageToken : ''}`;
      const r = await feishuAPI('GET', url);
      if (r.code !== 0) return res.status(500).json({ error: r.msg });
      if (r.data.items) allRecords = allRecords.concat(r.data.items);
      pageToken = r.data.has_more ? r.data.page_token : '';
    } while (pageToken);

    const schedules = allRecords.map(rec => {
      const f = rec.fields;
      let dateStr = '';
      if (f['日期']) {
        const d = new Date(f['日期']);
        dateStr = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
      }
      const startMin = parseTime(f['开始时间']);
      const endMin = parseTime(f['结束时间']);
      return {
        id: rec.record_id,
        leader: f['领导姓名'] || '',
        date: dateStr,
        start_time: f['开始时间'] || '',
        end_time: f['结束时间'] || '',
        start_min: startMin,
        end_min: endMin || (startMin ? startMin + 90 : null),
        title: f['事项'] || '',
        location: f['地点'] || '',
        participants: f['参与人'] || '',
        importance: f['重要程度'] || '中',
        notes: f['备注'] || '',
        is_allday: f['开始时间'] === '全天',
      };
    }).filter(s => s.date && s.leader);

    schedules.sort((a, b) => a.date !== b.date ? a.date.localeCompare(b.date) : (a.start_min||0) - (b.start_min||0));
    res.json(schedules);
  } catch (e) {
    console.error('获取行程失败:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ---- API: 领导列表 ----
app.get('/api/leaders', (req, res) => {
  res.json(LEADERS);
});

// ---- 首页 ----
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n  西影集团领导行程系统已启动: http://localhost:${PORT}\n`);
});
