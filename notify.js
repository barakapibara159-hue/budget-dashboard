const fetch = require('node-fetch');
const fs = require('fs');
const path = require('path');

function loadConfig() {
  const config = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf-8'));
  return {
    webhook_url: process.env.FEISHU_WEBHOOK_URL || config.notify.webhook_url,
    dashboard_url: process.env.DASHBOARD_URL || config.dashboard_url,
  };
}

function fmt(n) { return Math.round(n).toLocaleString('zh-CN'); }

// 发送飞书机器人消息（Webhook）
async function sendFeishuNotification(stats) {
  const config = loadConfig();
  if (!config.webhook_url) {
    console.log('⚠️  未配置飞书 Webhook URL，跳过通知');
    return;
  }

  const o = stats.overview;
  const usageRate = parseFloat(o.usageRate);
  const statusEmoji = usageRate > 100 ? '🔴' : usageRate > 80 ? '🟡' : '🟢';

  // 找最新预算月份
  const latestMonth = stats.records.reduce((max, r) => (r.budgetMonth || 0) > max ? r.budgetMonth : max, 0);

  // ===== 1. 最新预算月份花费总览 =====
  const overviewLines = [
    `${statusEmoji} **总预算：** ¥${fmt(o.totalBudget)}`,
    `💰 **${latestMonth}月已花费：** ¥${fmt(o.totalSpent)}（使用率 ${o.usageRate}%）`,
    `📝 **费用笔数：** ${stats.records.length}笔`,
  ];

  // ===== 2. 报销状态 =====
  const reimbursedInfo = stats.byStatus['已报销完'] || { count: 0, amount: 0 };
  const pendingStatuses = Object.entries(stats.byStatus)
    .filter(([name, info]) => name !== '已报销完' && info.count > 0);
  const pendingTotal = pendingStatuses.reduce((s, [, info]) => s + info.count, 0);
  const pendingAmount = pendingStatuses.reduce((s, [, info]) => s + info.amount, 0);

  const statusLines = [
    `✅ **${latestMonth}月已报销完：** ${reimbursedInfo.count}笔　¥${fmt(reimbursedInfo.amount)}`,
    `⏳ **待报销累计：** ${pendingTotal}笔　¥${fmt(pendingAmount)}`,
  ];
  if (pendingStatuses.length > 0) {
    pendingStatuses.forEach(([name, info]) => {
      statusLines.push(`　　· ${name}：${info.count}笔　¥${fmt(info.amount)}`);
    });
  }

  // ===== 3. 各部门花销排名 =====
  const deptRanking = Object.entries(stats.byDepartment)
    .filter(([, info]) => info.spent > 0)
    .sort((a, b) => b[1].spent - a[1].spent)
    .map(([name, info], idx) => {
      const pct = o.totalSpent > 0 ? ((info.spent / o.totalSpent) * 100).toFixed(1) : '0';
      const medal = idx === 0 ? '🥇' : idx === 1 ? '🥈' : idx === 2 ? '🥉' : `${idx + 1}.`;
      return `${medal} ${name}：¥${fmt(info.spent)}（${pct}%）`;
    });

  // ===== 4. 分类花费占比 =====
  const totalSpent = o.totalSpent || 1;
  const categoryLines = Object.entries(stats.byCategory)
    .sort((a, b) => b[1] - a[1])
    .map(([name, amount]) => {
      const pct = ((amount / totalSpent) * 100).toFixed(1);
      return `　${name}：¥${fmt(amount)}（${pct}%）`;
    });

  // ===== 构建卡片 =====
  const elements = [];

  // 花费总览
  elements.push({
    tag: 'div',
    text: { tag: 'lark_md', content: overviewLines.join('\n') },
  });

  elements.push({ tag: 'hr' });

  // 报销状态
  elements.push({
    tag: 'div',
    text: { tag: 'lark_md', content: `📋 **报销进度**\n${statusLines.join('\n')}` },
  });

  elements.push({ tag: 'hr' });

  // 各部门花销排名
  elements.push({
    tag: 'div',
    text: { tag: 'lark_md', content: `🏢 **各部门花销排名**\n${deptRanking.join('\n')}` },
  });

  elements.push({ tag: 'hr' });

  // 分类花费占比
  elements.push({
    tag: 'div',
    text: { tag: 'lark_md', content: `🍩 **分类花费占比**\n${categoryLines.join('\n')}` },
  });

  elements.push({ tag: 'hr' });

  // 看板链接
  if (config.dashboard_url) {
    elements.push({
      tag: 'action',
      actions: [{
        tag: 'button',
        text: { tag: 'plain_text', content: '📊 查看完整看板' },
        url: config.dashboard_url,
        type: 'primary',
      }],
    });
  }

  // 更新时间
  elements.push({
    tag: 'note',
    elements: [{
      tag: 'plain_text',
      content: `数据更新于 ${new Date(stats.updatedAt).toLocaleString('zh-CN')}`,
    }],
  });

  const msg = {
    msg_type: 'interactive',
    card: {
      header: {
        title: { tag: 'plain_text', content: `${statusEmoji} ${latestMonth}月预算看板日报` },
        template: usageRate > 100 ? 'red' : usageRate > 80 ? 'orange' : 'green',
      },
      elements,
    },
  };

  const res = await fetch(config.webhook_url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(msg),
  });

  const data = await res.json();
  if (data.code === 0 || data.StatusCode === 0) {
    console.log('✅ 飞书通知发送成功');
  } else {
    console.error('❌ 飞书通知发送失败:', JSON.stringify(data));
  }
}

module.exports = { sendFeishuNotification };

if (require.main === module) {
  const dataPath = path.join(__dirname, 'data.json');
  if (!fs.existsSync(dataPath)) {
    console.error('❌ data.json 不存在，请先运行 fetch_data.js');
    process.exit(1);
  }
  const stats = JSON.parse(fs.readFileSync(dataPath, 'utf-8'));
  sendFeishuNotification(stats).catch(err => {
    console.error('❌ 错误:', err.message);
    process.exit(1);
  });
}
