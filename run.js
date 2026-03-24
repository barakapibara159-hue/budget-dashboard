const { main: fetchData } = require('./fetch_data');
const { sendFeishuNotification } = require('./notify');

async function run() {
  console.log('🚀 预算看板更新流程开始\n');
  console.log('═'.repeat(40));

  // Step 1: 拉取数据
  console.log('\n📊 Step 1: 拉取飞书表格数据...');
  const stats = await fetchData();
  console.log(`   ✅ 共 ${stats.records.length} 条记录，总花费 ¥${Math.round(stats.overview.totalSpent)}`);

  // Step 2: 看板 HTML 直接读取 data.json，无需额外生成
  console.log('\n📄 Step 2: data.json 已就绪，index.html 将自动加载');

  // Step 3: 发送飞书通知
  console.log('\n🔔 Step 3: 发送飞书通知...');
  await sendFeishuNotification(stats);

  console.log('\n═'.repeat(40));
  console.log('🎉 全部完成！');
  console.log(`   看板地址: 本地预览 → npx serve . -p 3000`);
}

run().catch(err => {
  console.error('\n❌ 运行失败:', err.message);
  process.exit(1);
});
