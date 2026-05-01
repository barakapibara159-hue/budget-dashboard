const { main: fetchData } = require('./fetch_data');
const { sendFeishuNotification } = require('./notify');

// 飞书推送节奏：北京时间 day-of-year % 3 === 0 时才推送（约每 3 天一次）
// 设 FORCE_NOTIFY=1 可强制推送（手动触发时用）
function shouldNotifyToday() {
  if (process.env.FORCE_NOTIFY === '1') return true;
  const now = new Date();
  const beijing = new Date(now.getTime() + 8 * 60 * 60 * 1000);
  const start = Date.UTC(beijing.getUTCFullYear(), 0, 0);
  const dayOfYear = Math.floor((beijing.getTime() - start) / 86400000);
  return dayOfYear % 3 === 0;
}

async function run() {
  console.log('🚀 预算看板更新流程开始\n');
  console.log('═'.repeat(40));

  // Step 1: 拉取数据
  console.log('\n📊 Step 1: 拉取飞书表格数据...');
  const stats = await fetchData();
  console.log(`   ✅ 共 ${stats.records.length} 条记录，总花费 ¥${Math.round(stats.overview.totalSpent)}`);

  // Step 2: 看板 HTML 直接读取 data.json，无需额外生成
  console.log('\n📄 Step 2: data.json 已就绪，index.html 将自动加载');

  // Step 3: 发送飞书通知（每 3 天一次）
  if (shouldNotifyToday()) {
    console.log('\n🔔 Step 3: 发送飞书通知...');
    await sendFeishuNotification(stats);
  } else {
    console.log('\n⏭️  Step 3: 今天跳过飞书推送（每 3 天推一次，FORCE_NOTIFY=1 可强制）');
  }

  console.log('\n═'.repeat(40));
  console.log('🎉 全部完成！');
  console.log(`   看板地址: 本地预览 → npx serve . -p 3000`);
}

run().catch(err => {
  console.error('\n❌ 运行失败:', err.message);
  process.exit(1);
});
