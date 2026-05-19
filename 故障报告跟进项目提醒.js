/**
 * 故障报告跟进项目定期提醒
 * Failure Report Follow-up Item Daily Reminder
 *
 * 触发逻辑（每天执行一次）：
 * - 验证列含"已通过"              → 跳过
 * - 验证列含"未通过" 或 "NA"      → 提醒责任人：完成时间 ≤2天提醒临期，已逾期提醒逾期
 * - 验证列含"未验证"              → 每天提醒验证人进行验证
 */

/**
 * 主函数：每日执行故障报告跟进项目提醒
 */
function dailyFollowUpReminder() {
  try {
    console.log('=== 开始执行故障报告跟进项目提醒 ===');

    const followUpData = getFollowUpData();
    if (!followUpData || followUpData.length === 0) {
      console.log('没有找到跟进项目数据');
      logSystemActivity('跟进项目提醒', '没有找到跟进项目数据，跳过执行');
      return;
    }

    // ownerMap[email] = { dueSoon: [], overdue: [] }
    const ownerMap = {};
    // verifierMap[email] = []
    const verifierMap = {};

    followUpData.forEach(item => {
      const verify = item.status || '';

      if (verify.includes('已通过')) return;

      if (verify.includes('未通过') || verify.includes('NA')) {
        const email = extractEmailFromPersonField(item.owner);
        if (!email) return;
        const days = calcDaysUntilDue(item.dueDate);
        if (days < 0) {
          if (!ownerMap[email]) ownerMap[email] = { dueSoon: [], overdue: [] };
          ownerMap[email].overdue.push(item);
        } else if (days <= 2) {
          if (!ownerMap[email]) ownerMap[email] = { dueSoon: [], overdue: [] };
          ownerMap[email].dueSoon.push(item);
        }
        // days > 2: 不提醒

      } else if (verify.includes('未验证')) {
        const email = extractEmailFromPersonField(item.verifier);
        if (!email) return;
        if (!verifierMap[email]) verifierMap[email] = [];
        verifierMap[email].push(item);
      }
    });

    let ownerSent = 0;
    for (const email in ownerMap) {
      const { dueSoon, overdue } = ownerMap[email];
      if (dueSoon.length === 0 && overdue.length === 0) continue;
      const subject = overdue.length > 0
        ? '【逾期提醒】故障报告跟进项目逾期 / Follow-up Items Overdue'
        : '【临期提醒】故障报告跟进项目即将到期 / Follow-up Items Due Soon';
      GmailApp.sendEmail(email, subject, '请使用支持HTML的邮件客户端查看此邮件。', {
        htmlBody: generateFollowUpOwnerEmailContent(dueSoon, overdue),
        name: '故障报告提醒系统'
      });
      ownerSent++;
      console.log(`✅ 责任人提醒 → ${email} (临期:${dueSoon.length} 逾期:${overdue.length})`);
    }

    let verifierSent = 0;
    for (const email in verifierMap) {
      const items = verifierMap[email];
      if (items.length === 0) continue;
      GmailApp.sendEmail(email,
        '【验证提醒】故障报告跟进项目待验证 / Follow-up Items Pending Verification',
        '请使用支持HTML的邮件客户端查看此邮件。', {
          htmlBody: generateFollowUpVerifierEmailContent(items),
          name: '故障报告提醒系统'
        });
      verifierSent++;
      console.log(`✅ 验证人提醒 → ${email} (${items.length}条待验证)`);
    }

    logSystemActivity('跟进项目提醒', `成功执行，责任人提醒 ${ownerSent} 封，验证人提醒 ${verifierSent} 封，共处理 ${followUpData.length} 条跟进项目`);
    console.log('=== 故障报告跟进项目提醒执行完成 ===');

  } catch (error) {
    console.error('故障报告跟进项目提醒执行出错:', error);
    logSystemActivity('跟进项目提醒', `执行出错: ${error.message}`, 'ERROR');
  }
}

/**
 * 读取跟进项目数据
 */
function getFollowUpData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
      .getSheetByName(CONFIG.FOLLOWUP_SHEET_NAME);
    if (!sheet) { console.error('跟进项目表未找到'); return []; }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const idx = getFieldIndexes(data[0]);
    const result = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[idx['序号 / followup_id']]) continue;
      result.push({
        id:          row[idx['序号 / followup_id']]               || '',
        reportNo:    row[idx['故障报告编号 / failure_report_no']]  || '',
        paType:      row[idx['行动类型 / pa_type']]                || '',
        paPlan:      row[idx['预防行动 / pa_plan']]                || '',
        owner:   String(row[idx['责任人 / pa_who']]                || ''),
        dueDate:     row[idx['完成时间 / pa_when']]                || '',
        verifier:String(row[idx['验证人 / pa_verifier']]           || ''),
        status:      row[idx['状态 / status']]                     || '',
        notes:       row[idx['跟进内容 / follow_up_notes']]        || ''
      });
    }

    console.log(`✅ 读取到 ${result.length} 条跟进项目数据`);
    return result;
  } catch (error) {
    console.error('读取跟进项目数据时出错:', error);
    return [];
  }
}

/**
 * 从"姓名【邮箱】"格式中提取邮箱
 */
function extractEmailFromPersonField(field) {
  if (!field) return '';
  const match = field.match(/【(.+?)】/);
  return match ? match[1].trim() : '';
}

/**
 * 从"姓名【邮箱】"格式中提取姓名
 */
function extractNameFromPersonField(field) {
  if (!field) return '';
  return field.replace(/【.+?】/, '').trim();
}

/**
 * 计算距离截止日期的天数（负数=已逾期，0=今天到期）
 */
function calcDaysUntilDue(dueDate) {
  try {
    const due = (dueDate instanceof Date) ? new Date(dueDate) : new Date(dueDate);
    if (isNaN(due.getTime())) return 999;
    due.setHours(0, 0, 0, 0);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return Math.round((due - today) / (1000 * 3600 * 24));
  } catch (e) {
    return 999;
  }
}

/**
 * 格式化日期值（兼容字符串和 Date 对象）
 */
function formatFollowUpDate(value) {
  if (!value) return '';
  const d = (value instanceof Date) ? value : new Date(value);
  return isNaN(d.getTime()) ? String(value) : formatDate(d);
}

/**
 * 生成责任人跟进提醒邮件内容
 */
function generateFollowUpOwnerEmailContent(dueSoonItems, overdueItems) {
  const hasOverdue = overdueItems.length > 0;
  const today = formatDate(new Date());
  const ownerName = extractNameFromPersonField((overdueItems[0] || dueSoonItems[0] || {}).owner || '');
  const accentColor = hasOverdue ? '#f44336' : '#f39c12';
  const darkColor   = hasOverdue ? '#d32f2f' : '#e65100';
  const bgColor     = hasOverdue ? '#ffebee' : '#fff8e1';

  const buildTable = (items, isOverdue) => {
    const headerGrad = isOverdue
      ? 'linear-gradient(135deg,#f44336,#d32f2f)'
      : 'linear-gradient(135deg,#f39c12,#e67e22)';
    const rowBgAlt = isOverdue ? '#fff5f5' : '#fffbf0';

    let rows = '';
    items.forEach((item, i) => {
      const days = calcDaysUntilDue(item.dueDate);
      const badge = isOverdue
        ? `<div style="background:linear-gradient(135deg,#f44336,#d32f2f);color:white;padding:6px 12px;border-radius:16px;font-size:12px;font-weight:600;display:inline-block;min-width:80px;text-align:center;"><span style="display:block;">[逾期] ${Math.abs(days)}天</span><span style="display:block;font-size:10px;opacity:0.9;">Days Overdue</span></div>`
        : `<div style="background:linear-gradient(135deg,#f39c12,#e67e22);color:white;padding:6px 12px;border-radius:16px;font-size:12px;font-weight:600;display:inline-block;min-width:80px;text-align:center;"><span style="display:block;">还剩 ${days}天</span><span style="display:block;font-size:10px;opacity:0.9;">Days Left</span></div>`;
      rows += `
        <tr style="background-color:${i % 2 === 0 ? rowBgAlt : '#ffffff'};">
          <td style="padding:12px;border-bottom:1px solid #e9ecef;font-weight:500;color:#2c3e50;">${item.reportNo}</td>
          <td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;max-width:220px;word-wrap:break-word;">${item.paPlan}</td>
          <td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">${extractNameFromPersonField(item.owner)}</td>
          <td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;font-family:monospace;">${formatFollowUpDate(item.dueDate)}</td>
          <td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">${item.status}</td>
          <td style="padding:12px;border-bottom:1px solid #e9ecef;text-align:center;">${badge}</td>
        </tr>`;
    });

    return `
      <div style="overflow-x:auto;margin-bottom:20px;">
        <table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.1);">
          <thead>
            <tr style="background:${headerGrad};color:white;">
              <th style="padding:12px;text-align:left;font-weight:600;">故障报告编号<br><span style="font-size:0.8em;opacity:0.9;">Report No.</span></th>
              <th style="padding:12px;text-align:left;font-weight:600;">预防行动<br><span style="font-size:0.8em;opacity:0.9;">Action Plan</span></th>
              <th style="padding:12px;text-align:left;font-weight:600;">责任人<br><span style="font-size:0.8em;opacity:0.9;">Owner</span></th>
              <th style="padding:12px;text-align:left;font-weight:600;">期限<br><span style="font-size:0.8em;opacity:0.9;">Due Date</span></th>
              <th style="padding:12px;text-align:left;font-weight:600;">状态<br><span style="font-size:0.8em;opacity:0.9;">Status</span></th>
              <th style="padding:12px;text-align:center;font-weight:600;">${isOverdue ? '逾期天数<br><span style="font-size:0.8em;opacity:0.9;">Overdue Days</span>' : '剩余天数<br><span style="font-size:0.8em;opacity:0.9;">Days Left</span>'}</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
  };

  let body = `
    <div style="font-family:Arial,sans-serif;max-width:900px;margin:0 auto;background-color:#f8f9fa;padding:20px;">
      <div style="background:${bgColor};border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;border-left:5px solid ${accentColor};">
        <h2 style="color:${darkColor};text-align:center;margin-bottom:20px;border-bottom:3px solid ${accentColor};padding-bottom:10px;">
          ${hasOverdue ? '[逾期提醒] 故障报告跟进项目逾期' : '[临期提醒] 故障报告跟进项目即将到期'}<br>
          <span style="font-size:0.8em;">${hasOverdue ? 'Follow-up Items Overdue Reminder' : 'Follow-up Items Due Soon Reminder'}</span>
        </h2>
        <p style="font-size:16px;line-height:1.6;color:${darkColor};">
          您好${ownerName ? ' ' + ownerName : ''}！（${today}）以下故障报告跟进项目需要您处理：<br>
          <span style="font-size:0.9em;opacity:0.85;">Hello${ownerName ? ' ' + ownerName : ''}! The following follow-up items require your attention (${today}):</span>
        </p>
      </div>`;

  if (overdueItems.length > 0) {
    body += `
      <div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;">
        <h3 style="color:#d32f2f;border-bottom:2px solid #f44336;padding-bottom:10px;margin-bottom:20px;">
          [逾期] 已逾期跟进项目 Overdue Items (${overdueItems.length}条)
        </h3>
        ${buildTable(overdueItems, true)}
      </div>`;
  }

  if (dueSoonItems.length > 0) {
    body += `
      <div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;">
        <h3 style="color:#e65100;border-bottom:2px solid #f39c12;padding-bottom:10px;margin-bottom:20px;">
          [临期] 即将到期跟进项目 Due Soon Items (${dueSoonItems.length}条)
        </h3>
        ${buildTable(dueSoonItems, false)}
      </div>`;
  }

  body += `
      <div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;">
        <div style="text-align:center;color:${darkColor};font-size:14px;line-height:1.6;">
          <p style="margin-bottom:10px;font-weight:600;">请及时处理以上跟进项目！<br>
          <span style="font-size:0.9em;opacity:0.85;">Please handle the above follow-up items promptly!</span></p>
          <p style="margin:0;font-style:italic;">此邮件由系统自动发送，请勿回复。<br>
          <span style="font-size:0.8em;opacity:0.7;">This email is automatically sent by the system, please do not reply.</span></p>
        </div>
      </div>
    </div>`;

  return body;
}

/**
 * 生成验证人验证提醒邮件内容
 */
function generateFollowUpVerifierEmailContent(items) {
  const today = formatDate(new Date());
  const verifierName = extractNameFromPersonField((items[0] || {}).verifier || '');

  let rows = '';
  items.forEach((item, i) => {
    rows += `
      <tr style="background-color:${i % 2 === 0 ? '#f0f4ff' : '#ffffff'};">
        <td style="padding:12px;border-bottom:1px solid #e9ecef;font-weight:500;color:#2c3e50;">${item.id}</td>
        <td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">${item.reportNo}</td>
        <td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">${item.paType}</td>
        <td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;max-width:220px;word-wrap:break-word;">${item.paPlan}</td>
        <td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;max-width:200px;word-wrap:break-word;">${item.notes}</td>
        <td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;font-family:monospace;">${formatFollowUpDate(item.dueDate)}</td>
        <td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">${extractNameFromPersonField(item.owner)}</td>
        <td style="padding:12px;border-bottom:1px solid #e9ecef;color:#34495e;">${item.status}</td>
      </tr>`;
  });

  return `
    <div style="font-family:Arial,sans-serif;max-width:900px;margin:0 auto;background-color:#f8f9fa;padding:20px;">
      <div style="background:#e8eaf6;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;border-left:5px solid #3f51b5;">
        <h2 style="color:#283593;text-align:center;margin-bottom:20px;border-bottom:3px solid #3f51b5;padding-bottom:10px;">
          [验证提醒] 故障报告跟进项目待验证<br>
          <span style="font-size:0.8em;">Follow-up Items Pending Verification</span>
        </h2>
        <p style="font-size:16px;line-height:1.6;color:#283593;">
          您好${verifierName ? ' ' + verifierName : ''}！（${today}）以下跟进项目等待您的验证：<br>
          <span style="font-size:0.9em;opacity:0.85;">Hello${verifierName ? ' ' + verifierName : ''}! The following follow-up items are pending your verification (${today}):</span>
        </p>
      </div>
      <div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;margin-bottom:20px;">
        <h3 style="color:#283593;border-bottom:2px solid #3f51b5;padding-bottom:10px;margin-bottom:20px;">
          [详情] 待验证跟进项目 Pending Verification Items (${items.length}条)
        </h3>
        <div style="overflow-x:auto;">
          <table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.1);">
            <thead>
              <tr style="background:linear-gradient(135deg,#3f51b5,#283593);color:white;">
                <th style="padding:12px;text-align:left;font-weight:600;">跟进编号<br><span style="font-size:0.8em;opacity:0.9;">Follow-up ID</span></th>
                <th style="padding:12px;text-align:left;font-weight:600;">故障报告编号<br><span style="font-size:0.8em;opacity:0.9;">Report No.</span></th>
                <th style="padding:12px;text-align:left;font-weight:600;">行动类型<br><span style="font-size:0.8em;opacity:0.9;">Type</span></th>
                <th style="padding:12px;text-align:left;font-weight:600;">预防行动<br><span style="font-size:0.8em;opacity:0.9;">Action Plan</span></th>
                <th style="padding:12px;text-align:left;font-weight:600;">跟进内容<br><span style="font-size:0.8em;opacity:0.9;">Follow-up Notes</span></th>
                <th style="padding:12px;text-align:left;font-weight:600;">完成时间<br><span style="font-size:0.8em;opacity:0.9;">Due Date</span></th>
                <th style="padding:12px;text-align:left;font-weight:600;">责任人<br><span style="font-size:0.8em;opacity:0.9;">Owner</span></th>
                <th style="padding:12px;text-align:left;font-weight:600;">状态<br><span style="font-size:0.8em;opacity:0.9;">Status</span></th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </div>
      <div style="background:#ffffff;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);padding:30px;">
        <div style="text-align:center;color:#283593;font-size:14px;line-height:1.6;">
          <p style="margin-bottom:10px;font-weight:600;">请及时对以上项目进行验证！<br>
          <span style="font-size:0.9em;opacity:0.85;">Please verify the above items promptly!</span></p>
          <p style="margin:0;font-style:italic;">此邮件由系统自动发送，请勿回复。<br>
          <span style="font-size:0.8em;opacity:0.7;">This email is automatically sent by the system, please do not reply.</span></p>
        </div>
      </div>
    </div>`;
}

/**
 * 测试函数：手动执行故障报告跟进项目提醒（收件人固定为测试邮箱）
 */
function testFollowUpReminder() {
  dailyFollowUpReminder();
}
