const fetch = require('node-fetch');
const fs = require('fs');
const path = require('path');

// 飞书富文本单元格解扁平 —— 可能是 string / number / Array<{text:...}> / {text:...}
function cellToString(cell) {
  if (cell === null || cell === undefined) return '';
  if (typeof cell === 'string') return cell;
  if (typeof cell === 'number' || typeof cell === 'boolean') return String(cell);
  if (Array.isArray(cell)) {
    return cell.map(cellToString).join('');
  }
  if (typeof cell === 'object') {
    if (typeof cell.text === 'string') return cell.text;
    if (typeof cell.name === 'string') return cell.name; // 人员字段
    if (typeof cell.value === 'string') return cell.value;
    return '';
  }
  return String(cell);
}

// 加载配置（支持环境变量覆盖，用于 GitHub Actions）
// 本地优先读 config.json（含密钥），找不到则读 config.example.json（仅静态配置，密钥从环境变量取）
function loadConfig() {
  const localPath = path.join(__dirname, 'config.json');
  const examplePath = path.join(__dirname, 'config.example.json');
  const configPath = fs.existsSync(localPath) ? localPath : examplePath;
  const config = JSON.parse(fs.readFileSync(configPath, 'utf-8'));
  return {
    app_id: process.env.FEISHU_APP_ID || config.feishu.app_id,
    app_secret: process.env.FEISHU_APP_SECRET || config.feishu.app_secret,
    spreadsheet_token: process.env.FEISHU_SPREADSHEET_TOKEN || config.feishu.spreadsheet_token,
    sheet_id: process.env.FEISHU_SHEET_ID || config.feishu.sheet_id,
    departments: config.departments,
    categories: config.categories,
    statuses: config.statuses,
    budget: config.budget,
  };
}

// 获取飞书 tenant_access_token
async function getTenantToken(appId, appSecret) {
  const res = await fetch('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ app_id: appId, app_secret: appSecret }),
  });
  const data = await res.json();
  if (data.code !== 0) throw new Error(`获取 token 失败: ${data.msg}`);
  return data.tenant_access_token;
}

// 获取所有工作表信息
async function getAllSheets(token, spreadsheetToken) {
  const url = `https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/${spreadsheetToken}/sheets/query`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const data = await res.json();
  if (data.code !== 0) throw new Error(`获取工作表列表失败: ${data.msg}`);
  const sheets = data.data.sheets;
  if (!sheets || sheets.length === 0) throw new Error('表格中没有工作表');
  console.log(`   找到 ${sheets.length} 个工作表: ${sheets.map(s => `${s.title}(${s.sheet_id})`).join(', ')}`);
  return sheets;
}

// 从工作表列表中找到预算表（按名称匹配）
function findBudgetSheet(sheets) {
  const budgetKeywords = ['预算', 'budget', 'Budget'];
  return sheets.find(s => budgetKeywords.some(kw => s.title.includes(kw)));
}

// 从预算工作表解析各部门预算（支持逐项预算，按部门汇总）
// 表头格式: 月份 | 部门 | 费用分类 | 分类细项 | 预估金额
function parseBudgetRows(rows, targetMonth) {
  if (!rows || rows.length < 2) return null;

  const header = rows[0];
  let monthCol = -1, deptCol = -1, amountCol = -1, categoryCol = -1, descCol = -1;

  if (header && Array.isArray(header)) {
    for (let i = 0; i < header.length; i++) {
      const h = cellToString(header[i]).trim();
      if (['月份', '月', 'month'].some(kw => h.includes(kw))) monthCol = i;
      else if (['部门', '团队'].some(kw => h.includes(kw))) deptCol = i;
      else if (['预估金额', '预算金额', '金额', '预算'].some(kw => h.includes(kw))) amountCol = i;
      else if (['费用分类', '分类'].some(kw => h === kw || h.startsWith(kw))) categoryCol = i;
      else if (['分类细项', '细项', '说明', '备注'].some(kw => h.includes(kw))) descCol = i;
    }
  }

  // 如果没找到关键列，用默认位置
  if (deptCol === -1) deptCol = 1;
  if (amountCol === -1) amountCol = 4;

  console.log(`   预算表列映射: 月份=${monthCol}, 部门=${deptCol}, 金额=${amountCol}`);

  // 按部门汇总预算
  const budget = {};
  // 同时记录逐项明细（可选，用于看板展示）
  const budgetDetails = [];
  let matchedRows = 0;

  const [targetYear, targetMon] = targetMonth ? targetMonth.split('-').map(Number) : [0, 0];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row[deptCol]) continue;

    const dept = cellToString(row[deptCol]).trim();
    const amount = parseFloat(cellToString(row[amountCol]).replace(/[,，]/g, '')) || 0;
    if (!dept || isNaN(amount)) continue;

    // 解析月份数字
    let budgetMonth = null;
    let isTargetMonth = false;
    if (monthCol !== -1 && row[monthCol]) {
      let mStr = cellToString(row[monthCol]).trim();
      if (/^\d{4,5}$/.test(mStr)) {
        const md = new Date((parseInt(mStr) - 25569) * 86400000);
        budgetMonth = md.getUTCMonth() + 1;
        mStr = `${md.getUTCFullYear()}年${budgetMonth}月`;
      } else {
        const mm = mStr.match(/(\d+)/);
        if (mm) budgetMonth = parseInt(mm[1]);
      }
      if (targetMonth) {
        const hasYear = mStr.includes(String(targetYear));
        const hasMonth = mStr.includes(`${targetMon}月`) || mStr.includes(`-${String(targetMon).padStart(2, '0')}`);
        isTargetMonth = hasYear && hasMonth;
      }
    } else {
      isTargetMonth = true;
    }

    // 所有行都加入明细
    budgetDetails.push({
      budgetMonth,
      department: dept,
      category: categoryCol !== -1 ? cellToString(row[categoryCol]).trim() : '',
      description: descCol !== -1 ? cellToString(row[descCol]).trim() : '',
      amount,
    });

    // 只有目标月份的行参与汇总
    if (isTargetMonth) {
      matchedRows++;
      budget[dept] = (budget[dept] || 0) + amount;
    }
  }

  console.log(`   匹配到 ${matchedRows} 条预算记录`);
  if (Object.keys(budget).length > 0) {
    for (const [dept, total] of Object.entries(budget)) {
      console.log(`   ${dept}: ¥${total}`);
    }
  }

  return Object.keys(budget).length > 0 ? { budget, budgetDetails } : null;
}

// 读取飞书表格数据
async function readSheet(token, spreadsheetToken, sheetId) {
  // 用范围格式读取整个工作表
  const range = `${sheetId}`;
  const url = `https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values/${range}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const data = await res.json();
  if (data.code !== 0) throw new Error(`读取表格失败: ${data.msg}`);
  const values = data.data && data.data.valueRange && data.data.valueRange.values;
  return values || [];
}

// 解析表格数据
// 根据截图推断表头: 日期 | 部门 | 费用分类 | 费用说明 | 金额 | 报销状态
// 如果实际表头不同，修改 COLUMN_MAP 即可
const COLUMN_MAP = {
  date: 0,        // 日期
  department: 1,  // 部门
  category: 2,    // 费用分类
  description: 3, // 费用说明/备注
  amount: 4,      // 金额
  status: 5,      // 报销状态
};

function parseRows(rows, config) {
  const header = rows[0];
  const dataRows = rows.slice(1);

  // 自动检测列索引（如果表头存在则按表头匹配）
  const colMap = { ...COLUMN_MAP };
  let budgetMonthCol = -1; // 预算月份列
  const headerKeywords = {
    date: ['支出日期', '日期', '时间', 'date'],
    department: ['部门', '团队', 'department'],
    category: ['费用分类', '分类', '类别', 'category'],
    description: ['分类细项', '细项', '说明', '备注', '描述', '费用说明', 'description'],
    amount: ['金额', '花费', 'amount'],
    status: ['报销状态', '状态', 'status'],
  };

  if (header && Array.isArray(header)) {
    for (const [key, keywords] of Object.entries(headerKeywords)) {
      const idx = header.findIndex(h =>
        h && keywords.some(kw => cellToString(h).includes(kw))
      );
      if (idx !== -1) colMap[key] = idx;
    }
    // 检测预算月份列
    const bmIdx = header.findIndex(h => h && ['预算月份'].some(kw => cellToString(h).includes(kw)));
    if (bmIdx !== -1) budgetMonthCol = bmIdx;
  }

  const records = [];
  for (const row of dataRows) {
    if (!row || !row[colMap.amount]) continue;
    const amount = parseFloat(cellToString(row[colMap.amount]).replace(/[,，]/g, ''));
    if (isNaN(amount)) continue;

    // 解析日期
    let dateStr = cellToString(row[colMap.date]);
    // 飞书日期可能是数字（Excel序列号）或字符串
    if (/^\d{4,5}$/.test(dateStr)) {
      const d = new Date((parseInt(dateStr) - 25569) * 86400000);
      dateStr = d.toISOString().split('T')[0];
    }
    // 处理 2026/3/4 格式 → 2026-03-04
    if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(dateStr)) {
      const parts = dateStr.split('/');
      dateStr = `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
    }

    // 解析预算月份（如 "3月" → 3）
    let budgetMonth = null;
    if (budgetMonthCol !== -1 && row[budgetMonthCol]) {
      const bmStr = cellToString(row[budgetMonthCol]).trim();
      const match = bmStr.match(/(\d+)/);
      if (match) budgetMonth = parseInt(match[1]);
    }

    const deptStr = cellToString(row[colMap.department]).trim();
    const catStr = cellToString(row[colMap.category]).trim();
    const statusStr = cellToString(row[colMap.status]).trim();
    records.push({
      date: dateStr,
      budgetMonth,
      department: deptStr || '未知',
      category: catStr || '其他',
      description: cellToString(row[colMap.description]).trim(),
      amount,
      status: statusStr || '未知',
    });
  }
  return records;
}

// 聚合统计数据
function aggregate(records, config) {
  const now = new Date();
  const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

  // 当月记录：优先按预算月份筛选，否则按日期筛选
  const currentMonNum = now.getMonth() + 1;
  const hasBudgetMonth = records.some(r => r.budgetMonth !== null);
  let monthRecords;
  if (hasBudgetMonth) {
    monthRecords = records.filter(r => r.budgetMonth === currentMonNum);
  } else {
    monthRecords = records.filter(r => r.date.startsWith(currentMonth));
  }
  // 如果当月没数据就用所有数据
  const data = monthRecords.length > 0 ? monthRecords : records;

  const totalBudget = Object.values(config.budget).reduce((s, v) => s + v, 0);
  const totalSpent = data.reduce((s, r) => s + r.amount, 0);

  // 按部门聚合
  const byDept = {};
  for (const dept of config.departments) {
    byDept[dept] = { budget: config.budget[dept] || 0, spent: 0, count: 0 };
  }
  for (const r of data) {
    if (!byDept[r.department]) {
      byDept[r.department] = { budget: 0, spent: 0, count: 0 };
    }
    byDept[r.department].spent += r.amount;
    byDept[r.department].count += 1;
  }

  // 按分类聚合（总体 + 各部门）
  const byCategory = {};
  const byCategoryDept = {};
  for (const r of data) {
    byCategory[r.category] = (byCategory[r.category] || 0) + r.amount;
    if (!byCategoryDept[r.department]) byCategoryDept[r.department] = {};
    byCategoryDept[r.department][r.category] = (byCategoryDept[r.department][r.category] || 0) + r.amount;
  }

  // 按报销状态
  const byStatus = {};
  for (const s of config.statuses) {
    byStatus[s] = { count: 0, amount: 0 };
  }
  for (const r of data) {
    if (!byStatus[r.status]) byStatus[r.status] = { count: 0, amount: 0 };
    byStatus[r.status].count += 1;
    byStatus[r.status].amount += r.amount;
  }

  // 按周聚合（最近8周）
  const weeklyData = {};
  for (const r of records) {
    if (!r.date) continue;
    const d = new Date(r.date);
    if (isNaN(d.getTime())) continue;
    // ISO 周一为一周开始
    const day = d.getDay() || 7;
    const monday = new Date(d);
    monday.setDate(d.getDate() - day + 1);
    const weekKey = monday.toISOString().split('T')[0];
    weeklyData[weekKey] = (weeklyData[weekKey] || 0) + r.amount;
  }
  // 排序取最近8周
  const weekKeys = Object.keys(weeklyData).sort().slice(-8);
  const weekly = weekKeys.map(k => ({ week: k, amount: weeklyData[k] }));

  return {
    updatedAt: now.toISOString(),
    currentMonth,
    overview: {
      totalBudget,
      totalSpent,
      remaining: totalBudget - totalSpent,
      usageRate: totalBudget > 0 ? ((totalSpent / totalBudget) * 100).toFixed(1) : 0,
    },
    byDepartment: byDept,
    byCategory,
    byCategoryDept,
    byStatus,
    weekly,
    records,
    budgetDetails: config._budgetDetails || [],
  };
}

// 主函数
async function main() {
  const config = loadConfig();

  let records;

  if (!config.app_id || !config.spreadsheet_token) {
    console.log('⚠️  未配置飞书凭证，使用模拟数据生成看板...');
    records = generateMockData(config);
  } else {
    console.log('📊 正在获取飞书 token...');
    const token = await getTenantToken(config.app_id, config.app_secret);

    // 获取所有工作表
    console.log('🔍 正在获取工作表列表...');
    const sheets = await getAllSheets(token, config.spreadsheet_token);

    // 找花销表（第一个非预算的工作表）
    const budgetSheet = findBudgetSheet(sheets);
    const expenseSheet = sheets.find(s => s !== budgetSheet) || sheets[0];
    console.log(`   花销表: ${expenseSheet.title}(${expenseSheet.sheet_id})`);
    if (budgetSheet) {
      console.log(`   预算表: ${budgetSheet.title}(${budgetSheet.sheet_id})`);
    }

    // 读取花销数据
    console.log('📋 正在读取花销数据...');
    const rows = await readSheet(token, config.spreadsheet_token, expenseSheet.sheet_id);
    if (rows && rows.length > 0) {
      console.log(`   ✅ 读取到 ${rows.length - 1} 条花销记录`);
    } else {
      console.log('   ⚠️ 花销表暂无数据');
    }
    records = parseRows(rows || [], config);

    // 读取预算数据（如果有预算工作表）
    if (budgetSheet) {
      console.log('💰 正在读取预算数据...');
      const budgetRows = await readSheet(token, config.spreadsheet_token, budgetSheet.sheet_id);
      if (budgetRows && budgetRows.length > 0) {
        const now = new Date();
        const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
        const budgetResult = parseBudgetRows(budgetRows, currentMonth);
        if (budgetResult) {
          console.log(`   ✅ 读取到 ${Object.keys(budgetResult.budget).length} 个部门预算`);
          // 用飞书表格的预算覆盖 config 里的默认值
          config.budget = { ...config.budget, ...budgetResult.budget };
          config._budgetDetails = budgetResult.budgetDetails;
          // 同时更新部门列表（如果预算表里有新部门）
          for (const dept of Object.keys(budgetResult.budget)) {
            if (!config.departments.includes(dept)) {
              config.departments.push(dept);
            }
          }
        } else {
          console.log(`   ⚠️ 未找到当前月份(${currentMonth})的预算数据，使用默认预算`);
        }
      } else {
        console.log('   ⚠️ 预算表暂无数据，使用默认预算');
      }
    }
  }

  const stats = aggregate(records, config);

  const outPath = path.join(__dirname, 'data.json');
  fs.writeFileSync(outPath, JSON.stringify(stats, null, 2), 'utf-8');
  console.log(`💾 数据已保存到 ${outPath}`);

  return stats;
}

// 模拟数据（未配置凭证时使用，方便预览看板）
function generateMockData(config) {
  const records = [];
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();

  const descriptions = {
    '会员类': ['ChatGPT Plus', 'Midjourney订阅', 'Claude Pro', 'Notion团队版', 'Figma团队版'],
    '兼职类': ['视频剪辑兼职', '文案撰写', '翻译费用', '设计外包', '数据标注'],
    '网络类': ['VPN年费', '云服务器', '域名续费', 'CDN流量', '企业邮箱'],
    '硬件类': ['机械键盘', '显示器', '摄像头', '麦克风', '硬盘'],
    '账号类': ['Adobe全家桶', '企业微信', '飞书高级版', '石墨文档', '蓝湖'],
    '其他': ['快递费', '打车费', '办公用品', '团建餐费', '书籍'],
  };

  for (let week = 0; week < 8; week++) {
    const weekDate = new Date(year, month, -week * 7 + now.getDate());
    for (const dept of config.departments) {
      const numItems = Math.floor(Math.random() * 3) + 1;
      for (let i = 0; i < numItems; i++) {
        const cat = config.categories[Math.floor(Math.random() * config.categories.length)];
        const descs = descriptions[cat] || ['其他费用'];
        const status = config.statuses[Math.floor(Math.random() * config.statuses.length)];
        const dayOffset = Math.floor(Math.random() * 7);
        const d = new Date(weekDate);
        d.setDate(d.getDate() - dayOffset);

        records.push({
          date: d.toISOString().split('T')[0],
          department: dept,
          category: cat,
          description: descs[Math.floor(Math.random() * descs.length)],
          amount: Math.round((Math.random() * 2000 + 50) * 100) / 100,
          status,
        });
      }
    }
  }

  return records;
}

module.exports = { main, loadConfig, aggregate };

if (require.main === module) {
  main().catch(err => {
    console.error('❌ 错误:', err.message);
    process.exit(1);
  });
}
