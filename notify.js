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

// 生成进度条（文字版）
function progressBar(rate, len) {
  len = len || 10;
  const filled = Math.min(Math.round(rate / 100 * len), len);
  return '█'.repeat(filled) + '░'.repeat(len - filled);
}

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

  // ===== 1. 各部门花费排名 + 3. 使用率 =====
  const deptRanking = Object.entries(stats.byDepartment)
    .filter(([, info]) => info.spent > 0 || info.budget > 0)
    .sort((a, b) => b[1].spent - a[1].spent)
    .map(([name, info], idx) => {
      const rate = info.budget > 0 ? ((info.spent / info.budget) * 100).toFixed(1) : '-';
      const emoji = rate > 100 ? '🔴' : rate > 80 ? '🟡' : '🟢';
      const bar = info.budget > 0 ? progressBar(parseFloat(rate), 8) : '--------';
      return `${emoji} **${name}**\n　　¥${fmt(info.spent)} / ¥${fmt(info.budget)}　${bar} ${rate}%`;
    });

  // ===== 2. 超预算部门（醒目版）=====
  const overBudgetDepts = Object.entries(stats.byDepartment)
    .filter(([, info]) => info.budget > 0 && info.spent > info.budget)
    .sort((a, b) => (b[1].spent - b[1].budget) - (a[1].spent - a[1].budget))
    .map(([name, info]) => {
      const overAmount = Math.round(info.spent - info.budget);
      const rate = ((info.spent / info.budget) * 100).toFixed(1);
      return `🚨 **${name}** 超支 ¥${fmt(overAmount)}（使用率 ${rate}%）`;
    });

  // ===== 4. 分类花费占比 =====
  const totalSpent = o.totalSpent || 1;
  const categoryLines = Object.entries(stats.byCategory)
    .sort((a, b) => b[1] - a[1])
    .map(([name, amount]) => {
      const pct = ((amount / totalSpent) * 100).toFixed(1);
      return `　${name}：¥${fmt(amount)}（${pct}%）`;
    });

  // ===== 5. 本周新增花费 =====
  let weekCompare = '';
  if (stats.weekly && stats.weekly.length >= 2) {
    const thisWeek = stats.weekly[stats.weekly.length - 1];
    const lastWeek = stats.weekly[stats.weekly.length - 2];
    const diff = thisWeek.amount - lastWeek.amount;
    const diffPct = lastWeek.amount > 0 ? ((diff / lastWeek.amount) * 100).toFixed(1) : '-';
    const arrow = diff > 0 ? '📈 ↑' : diff < 0 ? '📉 ↓' : '➡️';
    weekCompare = `**本周花费：** ¥${fmt(thisWeek.amount)}（${thisWeek.week}）\n**上周花费：** ¥${fmt(lastWeek.amount)}（${lastWeek.week}）\n**环比：** ${arrow} ${diff > 0 ? '+' : ''}¥${fmt(diff)}（${diffPct}%）`;
  } else if (stats.weekly && stats.weekly.length === 1) {
    weekCompare = `**本周花费：** ¥${fmt(stats.weekly[0].amount)}（${stats.weekly[0].week}）`;
  } else {
    weekCompare = '暂无周数据';
  }

  // ===== 6. 报销状态 =====
  const statusLines = Object.entries(stats.byStatus)
    .filter(([, info]) => info.count > 0)
    .map(([name, info]) => `　${name}：${info.count}笔　¥${fmt(info.amount)}`);
  const hasStatus = statusLines.length > 0;

  // ===== 构建卡片 =====
  const elements = [];

  // 总览
  elements.push({
    tag: 'div',
    text: {
      tag: 'lark_md',
      content: [
        `${statusEmoji} **总预算：** ¥${fmt(o.totalBudget)}`,
        `💰 **已花费：** ¥${fmt(o.totalSpent)}（${o.usageRate}%）`,
        `💵 **剩余：** ¥${fmt(o.remaining)}`,
        `📝 **费用笔数：** ${stats.records.length}笔`,
      ].join('\n'),
    },
  });

  elements.push({ tag: 'hr' });

  // 超预算提醒（醒目）
  if (overBudgetDepts.length > 0) {
    elements.push({
      tag: 'div',
      text: {
        tag: 'lark_md',
        content: `⚠️ **【超预算警告】${overBudgetDepts.length}个部门超支**\n${overBudgetDepts.join('\n')}`,
      },
    });
    elements.push({ tag: 'hr' });
  }

  // 各部门花费排名 + 使用率
  elements.push({
    tag: 'div',
    text: {
      tag: 'lark_md',
      content: `🏢 **各部门预算使用率**（按花费排序）\n\n${deptRanking.join('\n')}`,
    },
  });

  elements.push({ tag: 'hr' });

  // 分类花费占比
  elements.push({
    tag: 'div',
    text: {
      tag: 'lark_md',
      content: `🍩 **分类花费占比**\n${categoryLines.join('\n')}`,
    },
  });

  elements.push({ tag: 'hr' });

  // 本周 vs 上周
  elements.push({
    tag: 'div',
    text: {
      tag: 'lark_md',
      content: `📊 **周花费趋势**\n${weekCompare}`,
    },
  });

  // 报销状态
  if (hasStatus) {
    elements.push({ tag: 'hr' });
    elements.push({
      tag: 'div',
      text: {
        tag: 'lark_md',
        content: `📋 **报销进度**\n${statusLines.join('\n')}`,
      },
    });
  }

  elements.push({ tag: 'hr' });

  // 查看看板按钮
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
        title: { tag: 'plain_text', content: `${statusEmoji} 预算看板日报 · ${stats.currentMonth}` },
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
