/**
 * 故障报告邮件自动提醒系统
 * Failure Report Email Auto-Reminder System
 * 
 * 功能说明：
 * 1. 基于车间和工序筛选需要提醒的邮箱
 * 2. 每周定期提醒所有故障报告
 * 3. 每天提醒超期故障报告（超过7天未提交）
 * 4. 采用卡片式邮件模板，中英文双语显示
 */

// 全局配置
const CONFIG = {
  SPREADSHEET_ID: '1YAPdZKVEOHgCGIJRQwWTQBmwaWIS4yd1SQKJJfRCtAU',
  FAILURE_SHEET_NAME: 'Failure_Database',
  NOTIFICATION_SHEET_NAME: '通知清单',
  FOLLOWUP_SHEET_NAME: 'Failure_Report_followup',
  OVERDUE_DAYS: 7, // 超期天数阈值
  LOG_SHEET_NAME: '系统日志' // 日志记录表
};

/**
 * 主函数：每周执行故障报告提醒
 */
function weeklyFailureReportReminder() {
  try {
    console.log('=== 开始执行每周故障报告提醒 ===');
    
    // 获取需要提醒的故障报告数据
    const failureData = getFailureReportData();
    if (!failureData || failureData.length === 0) {
      console.log('没有找到需要提醒的故障报告数据');
      return;
    }
    
    // 获取通知邮箱列表
    const notificationEmails = getNotificationEmails();
    if (!notificationEmails || notificationEmails.length === 0) {
      console.log('没有找到通知邮箱地址');
      return;
    }
    
    // 生成邮件内容
    const emailContent = generateWeeklyEmailContent(failureData);
    
    // 发送邮件
    sendFailureReportEmails(notificationEmails, '每周故障报告提醒', emailContent);
    
    // 记录日志
    logSystemActivity('每周提醒', `成功发送给 ${notificationEmails.length} 个邮箱，包含 ${failureData.length} 条故障报告`);
    
    console.log('=== 每周故障报告提醒执行完成 ===');
    
  } catch (error) {
    console.error('每周故障报告提醒执行出错:', error);
    logSystemActivity('每周提醒', `执行出错: ${error.message}`, 'ERROR');
  }
}

/**
 * 主函数：每日执行超期故障报告提醒
 */
function dailyOverdueFailureReportReminder() {
  try {
    console.log('=== 开始执行每日超期故障报告提醒 ===');
    
    // 获取超期故障报告数据
    const overdueData = getOverdueFailureReportData();
    if (!overdueData || overdueData.length === 0) {
      console.log('没有找到超期故障报告数据');
      return;
    }
    
    // 获取通知邮箱列表
    const notificationEmails = getNotificationEmails();
    if (!notificationEmails || notificationEmails.length === 0) {
      console.log('没有找到通知邮箱地址');
      return;
    }
    
    // 生成邮件内容
    const emailContent = generateOverdueEmailContent(overdueData);
    
    // 发送邮件
    sendFailureReportEmails(notificationEmails, '超期故障报告紧急提醒', emailContent);
    
    // 记录日志
    logSystemActivity('每日超期提醒', `成功发送给 ${notificationEmails.length} 个邮箱，包含 ${overdueData.length} 条超期故障报告`);
    
    console.log('=== 每日超期故障报告提醒执行完成 ===');
    
  } catch (error) {
    console.error('每日超期故障报告提醒执行出错:', error);
    logSystemActivity('每日超期提醒', `执行出错: ${error.message}`, 'ERROR');
  }
}

/**
 * 获取故障报告数据
 */
function getFailureReportData() {
  try {
    console.log('开始获取故障报告数据...');
    
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.FAILURE_SHEET_NAME);
    
    if (!sheet) {
      console.error('故障报告数据表未找到');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.log('故障报告数据表为空');
      return [];
    }
    
    // 获取表头
    const headers = data[0];
    console.log('🔍 表头字段:', headers);
    
    const fieldIndexes = getFieldIndexes(headers);
    console.log('🔍 字段索引映射:', fieldIndexes);
    
    // 处理数据行（跳过表头）
    const failureData = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // 调试：显示原始行数据
      if (i <= 3) { // 只显示前3行用于调试
        console.log(`📊 第${i}行原始数据:`, row);
      }
      
      const failureRecord = {
        id: row[fieldIndexes.编号] || '',
        machineId: row[fieldIndexes.机台号] || '',
        description: row[fieldIndexes.问题描述] || '',
        submitDate: row[fieldIndexes.提交日期] || '',
        workshop: row[fieldIndexes.车间] || '',
        process: row[fieldIndexes.工序] || '',
        reportNumber: row[fieldIndexes.故障报告编号] || '',
        assignDate: row[fieldIndexes.分配日期] || '',
        uploadDate: row[fieldIndexes.上传日期] || '',
        attachment: row[fieldIndexes.附件] || '',
        overdueDays: calculateOverdueDays(row[fieldIndexes.分配日期])
      };
      
      // 调试：显示处理后的记录
      if (i <= 3) {
        console.log(`📋 第${i}行处理后记录:`, failureRecord);
      }
      
      failureData.push(failureRecord);
    }
    
    console.log(`✅ 成功获取 ${failureData.length} 条故障报告数据`);
    return failureData;
    
  } catch (error) {
    console.error('获取故障报告数据时出错:', error);
    return [];
  }
}

/**
 * 获取超期故障报告数据
 * 只返回超期天数大于等于7天且未上传的故障报告
 */
function getOverdueFailureReportData() {
  try {
    const allData = getFailureReportData();
    
    // 筛选条件：超期天数 >= 7天 且 未上传
    const overdueData = allData.filter(record => {
      const isOverdue = record.overdueDays >= CONFIG.OVERDUE_DAYS;
      const isUnuploaded = !isFailureReportUploaded(record).isUploaded;
      
      return isOverdue && isUnuploaded;
    });
    
    console.log(`找到 ${overdueData.length} 条超期且未上传的故障报告（超期天数 >= ${CONFIG.OVERDUE_DAYS}天）`);
    
    // 按超期天数排序，超期天数多的排在前面
    overdueData.sort((a, b) => b.overdueDays - a.overdueDays);
    
    // 显示超期天数分布
    const overdueDistribution = {};
    overdueData.forEach(record => {
      const days = record.overdueDays;
      overdueDistribution[days] = (overdueDistribution[days] || 0) + 1;
    });
    
    console.log('📊 超期天数分布:');
    Object.keys(overdueDistribution).sort((a, b) => b - a).forEach(days => {
      console.log(`  ${days}天: ${overdueDistribution[days]}条`);
    });
    
    // 统计筛选前后的数据
    const totalOverdue = allData.filter(record => record.overdueDays >= CONFIG.OVERDUE_DAYS);
    const uploadedOverdue = totalOverdue.filter(record => isFailureReportUploaded(record).isUploaded);
    
    console.log(`📊 筛选统计:`);
    console.log(`  总超期记录: ${totalOverdue.length} 条`);
    console.log(`  已上传超期: ${uploadedOverdue.length} 条 (已排除)`);
    console.log(`  未上传超期: ${overdueData.length} 条 (需要提醒)`);
    
    return overdueData;
    
  } catch (error) {
    console.error('获取超期故障报告数据时出错:', error);
    return [];
  }
}

/**
 * 获取通知邮箱列表
 */
function getNotificationEmails() {
  try {
    console.log('开始获取通知邮箱地址...');
    
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.NOTIFICATION_SHEET_NAME);
    
    if (!sheet) {
      console.error('通知清单表未找到');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.log('通知清单表为空');
      return [];
    }
    
    // 获取表头
    const headers = data[0];
    const workshopIndex = headers.findIndex(h => h === '车间');
    const processIndex = headers.findIndex(h => h === '工序');
    const emailIndex = headers.findIndex(h => h === '邮箱');
    
    if (workshopIndex === -1 || processIndex === -1 || emailIndex === -1) {
      console.error('通知清单表缺少必要字段：车间、工序、邮箱');
      return [];
    }
    
    // 获取所有有效的邮箱地址
    const emails = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const email = String(row[emailIndex] || '').trim();
      const workshop = String(row[workshopIndex] || '').trim();
      const process = String(row[processIndex] || '').trim();
      
      if (email && email.includes('@') && workshop && process) {
        emails.push({
          email: email,
          workshop: workshop,
          process: process
        });
        console.log(`✅ 找到有效邮箱: ${email} (车间: ${workshop}, 工序: ${process})`);
      }
    }
    
    console.log(`成功获取 ${emails.length} 个通知邮箱地址`);
    return emails;
    
  } catch (error) {
    console.error('获取通知邮箱地址时出错:', error);
    return [];
  }
}

/**
 * 生成每周提醒邮件内容
 * 只显示未上传的故障报告
 */
function generateWeeklyEmailContent(failureData) {
  try {
    const currentDate = new Date();
    const formattedDate = formatDate(currentDate);
    
    // 过滤出未上传的故障报告
    const unuploadedData = failureData.filter(record => {
      const uploadStatus = isFailureReportUploaded(record);
      return !uploadStatus.isUploaded;
    });
    
    let emailBody = `
      <div style="font-family: Arial, sans-serif; max-width: 900px; margin: 0 auto; background-color: #f8f9fa; padding: 20px;">
        <style>
          @keyframes pulse {
            0% { transform: scale(1); opacity: 1; }
            50% { transform: scale(1.2); opacity: 0.7; }
            100% { transform: scale(1); opacity: 1; }
          }
        </style>
        
        <div style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px;">
          <h2 style="color: #2c3e50; text-align: center; margin-bottom: 20px; border-bottom: 3px solid #3498db; padding-bottom: 10px;">
            [提醒] 故障报告定期提醒<br>
            <span style="font-size: 0.8em; color: #7f8c8d;">Failure Report Regular Reminder</span>
          </h2>
          <p style="color: #34495e; font-size: 16px; line-height: 1.6; margin-bottom: 20px;">
            您好！以下是本周（${formattedDate}）<strong>未上传</strong>的故障报告汇总情况：<br>
            <span style="font-size: 0.9em; color: #7f8c8d;">Hello! Below is the summary of this week's (${formattedDate}) <strong>unuploaded</strong> failure reports:</span>
          </p>
        </div>
    `;
    
    if (unuploadedData && unuploadedData.length > 0) {
      emailBody += `
        <div style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px;">
          <h3 style="color: #2c3e50; margin-bottom: 20px; display: flex; align-items: center;">
            [详情] 未上传故障报告详情<br>
            <span style="font-size: 0.8em; color: #2c3e50; margin-left: 10px;">Unuploaded Failure Report Details</span>
          </h3>
          <div style="overflow-x: auto;">
            <table style="width: 100%; border-collapse: collapse; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
              <thead>
                <tr style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white;">
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">编号<br><span style="font-size: 0.8em; opacity: 0.9;">ID</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">机台号<br><span style="font-size: 0.8em; opacity: 0.9;">Machine ID</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">问题描述<br><span style="font-size: 0.8em; opacity: 0.9;">Description</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">车间<br><span style="font-size: 0.8em; opacity: 0.9;">Workshop</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">工序<br><span style="font-size: 0.8em; opacity: 0.9;">Process</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">分配日期<br><span style="font-size: 0.8em; opacity: 0.9;">Assign Date</span></th>
                                     <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">超期天数<br><span style="font-size: 0.8em; opacity: 0.9;">Overdue Days</span></th>
                </tr>
              </thead>
              <tbody>
      `;
      
      for (let i = 0; i < unuploadedData.length; i++) {
        const record = unuploadedData[i];
        const rowStyle = i % 2 === 0 ? 'background-color: #f8f9fa;' : 'background-color: #ffffff;';
        
                 // 超期天数显示
         let overdueDisplay = '';
         if (record.overdueDays > CONFIG.OVERDUE_DAYS) {
           overdueDisplay = `
             <div style="display: inline-block; text-align: center;">
               <div style="background: linear-gradient(135deg, #e74c3c, #c0392b); color: white; padding: 6px 12px; border-radius: 16px; font-weight: 600; font-size: 12px; box-shadow: 0 2px 6px rgba(231, 76, 60, 0.3); display: inline-block; min-width: 80px;">
                 <span style="display: block;">[超期] ${record.overdueDays}天</span>
                 <span style="display: block; font-size: 10px; opacity: 0.9; font-weight: 400;">Days</span>
               </div>
             </div>
           `;
         } else {
           overdueDisplay = `
             <div style="display: inline-block; text-align: center;">
               <div style="background: linear-gradient(135deg, #f39c12, #e67e22); color: white; padding: 6px 12px; border-radius: 16px; font-weight: 600; font-size: 12px; box-shadow: 0 2px 6px rgba(243, 156, 18, 0.3); display: inline-block; min-width: 80px;">
                 <span style="display: block;">${record.overdueDays}天</span>
                 <span style="display: block; font-size: 10px; opacity: 0.9; font-weight: 400;">Days</span>
               </div>
             </div>
           `;
         }
         
         emailBody += `
           <tr style="${rowStyle}">
             <td style="padding: 12px; border-bottom: 1px solid #e9ecef; font-weight: 500; color: #2c3e50;">${record.id || ''}</td>
             <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; font-family: monospace;">${record.machineId || ''}</td>
             <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; max-width: 200px; word-wrap: break-word;">${record.description || ''}</td>
             <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e;">${record.workshop || ''}</td>
             <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e;">${record.process || ''}</td>
             <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; font-family: monospace;">${formatDate(record.assignDate) || ''}</td>
             <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; text-align: center;">
               ${overdueDisplay}
             </td>
           </tr>
         `;
      }
      
      emailBody += `
              </tbody>
            </table>
          </div>
        </div>
      `;
    } else {
      emailBody += `
        <div style="background-color: #d4edda; border: 1px solid #c3e6cb; border-radius: 8px; padding: 20px; text-align: center; margin-bottom: 20px;">
          <p style="color: #155724; font-size: 16px; margin: 0; font-weight: 500;">
            [完成] 本周所有故障报告都已上传完成！<br>
            <span style="font-size: 0.9em; color: #155724; opacity: 0.8;">All failure reports for this week have been uploaded!</span>
          </p>
        </div>
      `;
    }
    
    emailBody += `
        <div style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px;">
          <div style="text-align: center; color: #7f8c8d; font-size: 14px; line-height: 1.6;">
            <p style="margin-bottom: 10px;">请及时查看并处理相关故障报告。<br>
            <span style="font-size: 0.9em; opacity: 0.8;">Please review and handle related failure reports promptly.</span></p>
            <p style="margin: 0; font-style: italic;">此邮件由系统自动发送，请勿回复。<br>
            <span style="font-size: 0.8em; opacity: 0.8;">This email is automatically sent by the system, please do not reply.</span></p>
          </div>
        </div>
      </div>
    `;
    
    return emailBody;
    
  } catch (error) {
    console.error('生成每周提醒邮件内容时出错:', error);
    return '<p>邮件内容生成失败，请检查系统日志。</p>';
  }
}

/**
 * 生成超期提醒邮件内容
 * 只显示未上传的超期故障报告
 */
function generateOverdueEmailContent(overdueData) {
  try {
    const currentDate = new Date();
    const formattedDate = formatDate(currentDate);
    
    // overdueData 已经是从 getOverdueFailureReportData() 获取的超期且未上传的故障报告
    // 不需要再次过滤
    
    let emailBody = `
      <div style="font-family: Arial, sans-serif; max-width: 900px; margin: 0 auto; background-color: #f8f9fa; padding: 20px;">
        <style>
          @keyframes pulse {
            0% { transform: scale(1); opacity: 1; }
            50% { transform: scale(1.2); opacity: 0.7; }
            100% { transform: scale(1); opacity: 1; }
          }
        </style>
        
        <div style="background-color: #ffebee; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px; border-left: 5px solid #f44336;">
                     <h2 style="color: #d32f2f; text-align: center; margin-bottom: 20px; border-bottom: 3px solid #f44336; padding-bottom: 10px;">
             [紧急] 超期未上传故障报告紧急提醒 (≥${CONFIG.OVERDUE_DAYS}天)<br>
             <span style="font-size: 0.8em; color: #d32f2f;">Overdue Unuploaded Failure Report Urgent Reminder (≥${CONFIG.OVERDUE_DAYS} Days)</span>
           </h2>
           <p style="color: #d32f2f; font-size: 16px; line-height: 1.6; margin-bottom: 20px;">
             紧急通知！以下<strong>未上传</strong>的故障报告已超期${CONFIG.OVERDUE_DAYS}天或以上，请立即上传处理：<br>
             <span style="font-size: 0.9em; color: #d32f2f; opacity: 0.8;">Urgent Notice! The following <strong>unuploaded</strong> failure reports have been overdue for ${CONFIG.OVERDUE_DAYS} days or more, please upload and handle immediately:</span>
           </p>
        </div>
    `;
    
    if (overdueData && overdueData.length > 0) {
      emailBody += `
        <div style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px;">
                     <h3 style="color: #d32f2f; margin-bottom: 20px; display: flex; align-items: center;">
             [详情] 超期未上传故障报告详情 (≥${CONFIG.OVERDUE_DAYS}天)<br>
             <span style="font-size: 0.8em; color: #d32f2f; margin-left: 10px;">Overdue Unuploaded Failure Report Details (≥${CONFIG.OVERDUE_DAYS} Days)</span>
           </h3>
          <div style="overflow-x: auto;">
            <table style="width: 100%; border-collapse: collapse; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
              <thead>
                <tr style="background: linear-gradient(135deg, #f44336 0%, #d32f2f 100%); color: white;">
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">编号<br><span style="font-size: 0.8em; opacity: 0.9;">ID</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">机台号<br><span style="font-size: 0.8em; opacity: 0.9;">Machine ID</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">问题描述<br><span style="font-size: 0.8em; opacity: 0.9;">Description</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">车间<br><span style="font-size: 0.8em; opacity: 0.9;">Workshop</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">工序<br><span style="font-size: 0.8em; opacity: 0.9;">Process</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">分配日期<br><span style="font-size: 0.8em; opacity: 0.9;">Assign Date</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">超期天数<br><span style="font-size: 0.8em; opacity: 0.9;">Overdue Days</span></th>
                </tr>
              </thead>
              <tbody>
      `;
      
      for (let i = 0; i < overdueData.length; i++) {
        const record = overdueData[i];
        const rowStyle = i % 2 === 0 ? 'background-color: #fff5f5;' : 'background-color: #ffffff;';
        
        emailBody += `
          <tr style="${rowStyle}">
            <td style="padding: 12px; border-bottom: 1px solid #e9ecef; font-weight: 500; color: #2c3e50;">${record.id || ''}</td>
            <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; font-family: monospace;">${record.machineId || ''}</td>
            <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; max-width: 200px; word-wrap: break-word;">${record.description || ''}</td>
            <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e;">${record.workshop || ''}</td>
            <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e;">${record.process || ''}</td>
            <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; font-family: monospace;">${formatDate(record.assignDate) || ''}</td>
            <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #d32f2f; font-weight: 600; text-align: center;">
              <div style="background: linear-gradient(135deg, #f44336, #d32f2f); color: white; padding: 6px 12px; border-radius: 16px; font-weight: 600; font-size: 12px; box-shadow: 0 2px 6px rgba(244, 67, 54, 0.3); display: inline-block; min-width: 80px;">
                <span style="display: block;">[超期] ${record.overdueDays}天</span>
                <span style="display: block; font-size: 10px; opacity: 0.9; font-weight: 400;">Days</span>
              </div>
            </td>
          </tr>
        `;
      }
      
      emailBody += `
              </tbody>
            </table>
          </div>
        </div>
      `;
    } else {
      emailBody += `
        <div style="background-color: #d4edda; border: 1px solid #c3e6cb; border-radius: 8px; padding: 20px; text-align: center; margin-bottom: 20px;">
          <p style="color: #155724; font-size: 16px; margin: 0; font-weight: 500;">
            [完成] 所有超期故障报告都已上传完成！<br>
            <span style="font-size: 0.8em; color: #155724; opacity: 0.8;">All overdue failure reports have been uploaded!</span>
          </p>
        </div>
      `;
    }
    
    emailBody += `
        <div style="background-color: #ffebee; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px;">
          <div style="text-align: center; color: #d32f2f; font-size: 14px; line-height: 1.6;">
            <p style="margin-bottom: 10px; font-weight: 600;">[重要] 请立即处理这些超期故障报告！<br>
            <span style="font-size: 0.9em; opacity: 0.8;">Please handle these overdue failure reports immediately!</span></p>
            <p style="margin: 0; font-style: italic;">此邮件由系统自动发送，请勿回复。<br>
            <span style="font-size: 0.9em; opacity: 0.8;">This email is automatically sent by the system, please do not reply.</span></p>
          </div>
        </div>
      </div>
    `;
    
    return emailBody;
    
  } catch (error) {
    console.error('生成超期提醒邮件内容时出错:', error);
    return '<p>邮件内容生成失败，请检查系统日志。</p>';
  }
}

/**
 * 发送故障报告提醒邮件
 * 将多个邮箱地址合并到同一封邮件的收件人中，用最少的邮件数量发送
 */
function sendFailureReportEmails(notificationEmails, subject, emailContent) {
  try {
    console.log(`📧 开始发送邮件，主题: ${subject}`);
    console.log(`📋 通知邮箱总数: ${notificationEmails.length}`);
    
    // 收集所有有效的邮箱地址
    const allEmails = [];
    const emailDetails = [];
    
    for (const notification of notificationEmails) {
      try {
        console.log(`\n🔍 处理通知: ${notification.email} (车间: ${notification.workshop}, 工序: ${notification.process})`);
        
        // 根据车间和工序筛选故障报告
        // 如果是超期提醒，则只返回超期且未上传的故障报告
        const isOverdueReminder = subject.includes('超期');
        const filteredData = filterFailureDataByWorkshopAndProcess(notification.workshop, notification.process, isOverdueReminder);
        
        if (filteredData && filteredData.length > 0) {
          // 收集邮箱地址和详细信息
          allEmails.push(notification.email);
          emailDetails.push({
            email: notification.email,
            workshop: notification.workshop,
            process: notification.process,
            data: filteredData
          });
          
          console.log(`✅ 收集邮箱: ${notification.email} (车间: ${notification.workshop}, 工序: ${notification.process})`);
          console.log(`📧 包含 ${filteredData.length} 条故障报告`);
        } else {
          console.log(`⏭️ 跳过邮箱: ${notification.email}，没有匹配的故障报告`);
        }
        
      } catch (error) {
        console.error(`❌ 处理邮箱 ${notification.email} 时出错:`, error);
      }
    }
    
    // 如果有有效的邮箱地址，发送合并邮件
    if (allEmails.length > 0) {
      try {
        // 生成合并的邮件内容
        const mergedContent = generateMergedEmailContent(emailDetails, subject);
        
        // 发送合并邮件给所有邮箱
        GmailApp.sendEmail(
          allEmails.join(','), // 用逗号分隔多个邮箱地址
          subject,
          '此邮件包含HTML内容，请在支持HTML的邮件客户端中查看。',
          {
            htmlBody: mergedContent,
            name: '故障报告提醒系统'
          }
        );
        
        console.log(`\n📧 成功发送合并邮件:`);
        console.log(`📧 收件人: ${allEmails.join(', ')}`);
        console.log(`📧 邮件主题: ${subject}`);
        console.log(`📧 包含 ${emailDetails.length} 个邮箱的故障报告数据`);
        
        // 统计各邮箱的故障报告数量
        emailDetails.forEach(detail => {
          console.log(`  📧 ${detail.email} (${detail.workshop}/${detail.process}): ${detail.data.length} 条故障报告`);
        });
        
      } catch (error) {
        console.error('❌ 发送合并邮件失败:', error);
      }
    } else {
      console.log('⚠️ 没有找到需要发送邮件的有效邮箱地址');
    }
    
    console.log(`\n📊 邮件发送完成统计:`);
    console.log(`📧 有效邮箱: ${allEmails.length} 个`);
    console.log(`📧 发送邮件数量: ${allEmails.length > 0 ? 1 : 0} 封`);
    console.log(`📧 邮件主题: ${subject}`);
    
  } catch (error) {
    console.error('发送故障报告提醒邮件时出错:', error);
  }
}

/**
 * 解析车间字段，支持多车间格式（如：TB1/TB2）
 */
function parseWorkshopField(workshopField) {
  try {
    if (!workshopField || typeof workshopField !== 'string') {
      return [];
    }
    
    // 解析多车间格式（如：TB1/TB2 → ["TB1", "TB2"]）
    const workshopList = workshopField.split('/').map(w => w.trim()).filter(w => w);
    
    console.log(`🔧 车间字段解析: "${workshopField}" → [${workshopList.join(', ')}]`);
    
    return workshopList;
    
  } catch (error) {
    console.error('解析车间字段时出错:', error);
    return [];
  }
}

/**
 * 根据车间和工序筛选故障报告数据
 * 支持多车间格式（如：TB1/TB2）
 * 只返回未上传的故障报告
 * @param {string} workshop - 车间字段
 * @param {string} process - 工序字段
 * @param {boolean} isOverdueOnly - 是否只返回超期故障报告（默认false）
 */
function filterFailureDataByWorkshopAndProcess(workshop, process, isOverdueOnly = false) {
  try {
    const allData = getFailureReportData();
    
    // 解析多车间格式
    const workshopList = parseWorkshopField(workshop);
    
    if (workshopList.length === 0) {
      console.log(`⚠️ 警告: 车间字段 "${workshop}" 解析失败或为空`);
      return [];
    }
    
    console.log(`🔍 筛选条件 - 车间: ${workshop} (解析为: [${workshopList.join(', ')}]), 工序: ${process}, 仅超期: ${isOverdueOnly}`);
    
    // 使用"与"的关系筛选，支持多车间匹配，并且只返回未上传的故障报告
    const filteredData = allData.filter(record => {
      const workshopMatch = workshopList.includes(record.workshop);
      const processMatch = record.process === process;
      
      // 检查上传状态
      const uploadStatus = isFailureReportUploaded(record);
      const isUnuploaded = !uploadStatus.isUploaded;
      
      // 如果要求只返回超期故障报告，则检查超期天数
      const isOverdue = isOverdueOnly ? record.overdueDays >= CONFIG.OVERDUE_DAYS : true;
      
      if (workshopMatch && processMatch) {
        if (isUnuploaded) {
          if (isOverdueOnly && !isOverdue) {
            console.log(`⏭️ 跳过未超期: 故障报告 ${record.id} (车间: ${record.workshop}, 工序: ${record.process}) - ❌ 未上传但未超期${CONFIG.OVERDUE_DAYS}天`);
          } else {
            console.log(`✅ 匹配成功: 故障报告 ${record.id} (车间: ${record.workshop}, 工序: ${record.process}) - ❌ 未上传${isOverdueOnly ? '且超期' : ''}`);
          }
        } else {
          console.log(`⏭️ 跳过已上传: 故障报告 ${record.id} (车间: ${record.workshop}, 工序: ${record.process}) - ✅ 已上传`);
        }
      }
      
      // 返回匹配、未上传且满足超期条件的故障报告
      return workshopMatch && processMatch && isUnuploaded && isOverdue;
    });
    
    console.log(`📊 筛选结果: 找到 ${filteredData.length} 条匹配且未上传${isOverdueOnly ? '且超期' : ''}的故障报告`);
    
    return filteredData;
    
  } catch (error) {
    console.error('筛选故障报告数据时出错:', error);
    return [];
  }
}

/**
 * 生成个性化邮件内容
 */
function generatePersonalizedEmailContent(filteredData, notification, subject) {
  try {
    if (subject.includes('超期')) {
      return generateOverdueEmailContent(filteredData);
    } else {
      return generateWeeklyEmailContent(filteredData);
    }
  } catch (error) {
    console.error('生成个性化邮件内容时出错:', error);
    return generateWeeklyEmailContent(filteredData);
  }
}

/**
 * 生成合并的邮件内容
 * 将多个邮箱的故障报告数据合并到一封邮件中
 * @param {Array} emailDetails - 包含邮箱、车间、工序和故障报告数据的数组
 * @param {string} subject - 邮件主题
 * @returns {string} 合并后的HTML邮件内容
 */
function generateMergedEmailContent(emailDetails, subject) {
  try {
    const currentDate = new Date();
    const formattedDate = formatDate(currentDate);
    
    // 判断邮件类型
    const isOverdueReminder = subject.includes('超期');
    
    // 生成邮件头部
    let emailBody = `
      <div style="font-family: Arial, sans-serif; max-width: 900px; margin: 0 auto; background-color: #f8f9fa; padding: 20px;">
        <style>
          @keyframes pulse {
            0% { transform: scale(1); opacity: 1; }
            50% { transform: scale(1.2); opacity: 0.7; }
            100% { transform: scale(1); opacity: 1; }
          }
        </style>
    `;
    
    if (isOverdueReminder) {
      // 超期提醒邮件头部
      emailBody += `
        <div style="background-color: #ffebee; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px; border-left: 5px solid #f44336;">
          <h2 style="color: #d32f2f; text-align: center; margin-bottom: 20px; border-bottom: 3px solid #f44336; padding-bottom: 10px;">
            [紧急] 超期未上传故障报告紧急提醒 (≥${CONFIG.OVERDUE_DAYS}天)<br>
            <span style="font-size: 0.8em; color: #d32f2f;">Overdue Unuploaded Failure Report Urgent Reminder (≥${CONFIG.OVERDUE_DAYS} Days)</span>
          </h2>
          <p style="color: #d32f2f; font-size: 16px; line-height: 1.6; margin-bottom: 20px;">
            紧急通知！以下<strong>未上传</strong>的故障报告已超期${CONFIG.OVERDUE_DAYS}天或以上，请立即上传处理：<br>
            <span style="font-size: 0.9em; color: #d32f2f; opacity: 0.8;">Urgent Notice! The following <strong>unuploaded</strong> failure reports have been overdue for ${CONFIG.OVERDUE_DAYS} days or more, please upload and handle immediately:</span>
          </p>
        </div>
      `;
    } else {
      // 每周提醒邮件头部
      emailBody += `
        <div style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px;">
          <h2 style="color: #2c3e50; text-align: center; margin-bottom: 20px; border-bottom: 3px solid #3498db; padding-bottom: 10px;">
            [提醒] 故障报告定期提醒<br>
            <span style="font-size: 0.8em; color: #7f8c8d;">Failure Report Regular Reminder</span>
          </h2>
          <p style="color: #34495e; font-size: 16px; line-height: 1.6; margin-bottom: 20px;">
            您好！以下是本周（${formattedDate}）<strong>未上传</strong>的故障报告汇总情况：<br>
            <span style="font-size: 0.9em; color: #7f8c8d;">Hello! Below is the summary of this week's (${formattedDate}) <strong>unuploaded</strong> failure reports:</span>
          </p>
        </div>
      `;
    }
    
        // 根据邮件类型生成不同的内容结构
    let totalReports = 0;
    
    if (isOverdueReminder) {
      // 日报：按车间/工序分组，但相同车间/工序组合只显示一个表格
      // 按车间/工序分组数据，避免重复
      const workshopProcessGroups = {};
      const processedRecords = new Set(); // 用于跟踪已处理的记录
      
      emailDetails.forEach(detail => {
        if (detail.data && detail.data.length > 0) {
          // 只添加未处理的记录，避免重复
          detail.data.forEach(record => {
            const recordKey = `${record.id}_${record.workshop}_${record.process}`;
            if (!processedRecords.has(recordKey)) {
              // 使用实际的车间和工序作为分组键
              const key = `${record.workshop}/${record.process}`;
              if (!workshopProcessGroups[key]) {
                workshopProcessGroups[key] = [];
              }
              workshopProcessGroups[key].push(record);
              processedRecords.add(recordKey);
              totalReports++;
            }
          });
        }
      });
      
      // 为每个车间/工序组合生成一个表格
      Object.keys(workshopProcessGroups).forEach(key => {
        const [workshop, process] = key.split('/');
        const groupData = workshopProcessGroups[key];
        
        emailBody += `
          <div style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px;">
            <h3 style="color: #d32f2f; margin-bottom: 20px; display: flex; align-items: center; border-bottom: 2px solid #f44336; padding-bottom: 10px;">
              [详情] ${workshop}/${process} - 未上传故障报告详情<br>
              <span style="font-size: 0.8em; color: #d32f2f; margin-left: 10px;">${workshop}/${process} - Unuploaded Failure Report Details</span>
            </h3>
        `;
        
        // 生成表格
        emailBody += `
          <div style="overflow-x: auto;">
            <table style="width: 100%; border-collapse: collapse; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
              <thead>
                <tr style="background: linear-gradient(135deg, #f44336 0%, #d32f2f 100%); color: white;">
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">编号<br><span style="font-size: 0.8em; opacity: 0.9;">ID</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">机台号<br><span style="font-size: 0.8em; opacity: 0.9;">Machine ID</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">问题描述<br><span style="font-size: 0.8em; opacity: 0.9;">Description</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">分配日期<br><span style="font-size: 0.8em; opacity: 0.9;">Assign Date</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">超期天数<br><span style="font-size: 0.8em; opacity: 0.9;">Overdue Days</span></th>
                </tr>
              </thead>
              <tbody>
        `;
        
        // 添加数据行
        groupData.forEach((record, recordIndex) => {
          const rowStyle = recordIndex % 2 === 0 ? 'background-color: #fff5f5;' : 'background-color: #ffffff;';
          
          // 超期提醒邮件：所有记录都是超期的，显示红色标签
          const overdueDisplay = `
            <div style="display: inline-block; text-align: center;">
              <div style="background: linear-gradient(135deg, #f44336, #d32f2f); color: white; padding: 6px 12px; border-radius: 16px; font-weight: 600; font-size: 12px; box-shadow: 0 2px 6px rgba(244, 67, 54, 0.3); display: inline-block; min-width: 80px;">
                <span style="display: block;">[超期] ${record.overdueDays}天</span>
                <span style="display: block; font-size: 10px; opacity: 0.9; font-weight: 400;">Days</span>
              </div>
            </div>
          `;
          
          emailBody += `
            <tr style="${rowStyle}">
              <td style="padding: 12px; border-bottom: 1px solid #e9ecef; font-weight: 500; color: #2c3e50;">${record.id || ''}</td>
              <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; font-family: monospace;">${record.machineId || ''}</td>
              <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; max-width: 200px; word-wrap: break-word;">${record.description || ''}</td>
              <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; font-family: monospace;">${formatDate(record.assignDate) || ''}</td>
              <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; text-align: center;">
                ${overdueDisplay}
              </td>
            </tr>
          `;
        });
        
        emailBody += `
              </tbody>
            </table>
          </div>
          <div style="margin-top: 15px; text-align: right;">
            <span style="color: #7f8c8d; font-size: 14px;">
              共 ${groupData.length} 条故障报告 | Total: ${groupData.length} failure reports
            </span>
          </div>
        </div>
        `;
      });
    } else {
      // 周报：按工序分组，每个工序只显示一个表格
      // 按工序分组数据，避免重复
      const processGroups = {};
      const processedRecords = new Set(); // 用于跟踪已处理的记录
      
      emailDetails.forEach(detail => {
        if (detail.data && detail.data.length > 0) {
          if (!processGroups[detail.process]) {
            processGroups[detail.process] = [];
          }
          
          // 只添加未处理的记录，避免重复
          detail.data.forEach(record => {
            const recordKey = `${record.id}_${record.workshop}_${record.process}`;
            if (!processedRecords.has(recordKey)) {
              processGroups[detail.process].push(record);
              processedRecords.add(recordKey);
              totalReports++;
            }
          });
        }
      });
      
      // 为每个工序生成一个表格
      Object.keys(processGroups).forEach(process => {
        const processData = processGroups[process];
        
        emailBody += `
          <div style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px; margin-bottom: 20px;">
            <h3 style="color: #2c3e50; margin-bottom: 20px; display: flex; align-items: center; border-bottom: 2px solid #3498db; padding-bottom: 10px;">
              [详情] ${process} 工序 - 未上传故障报告详情<br>
              <span style="font-size: 0.8em; color: #2c3e50; margin-left: 10px;">${process} Process - Unuploaded Failure Report Details</span>
            </h3>
        `;
        
        // 生成表格
        emailBody += `
          <div style="overflow-x: auto;">
            <table style="width: 100%; border-collapse: collapse; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
              <thead>
                <tr style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white;">
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">编号<br><span style="font-size: 0.8em; opacity: 0.9;">ID</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">机台号<br><span style="font-size: 0.8em; opacity: 0.9;">Machine ID</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">问题描述<br><span style="font-size: 0.8em; opacity: 0.9;">Description</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">车间<br><span style="font-size: 0.8em; opacity: 0.9;">Workshop</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">分配日期<br><span style="font-size: 0.8em; opacity: 0.9;">Assign Date</span></th>
                  <th style="padding: 12px; text-align: left; border: none; font-weight: 600;">超期天数<br><span style="font-size: 0.8em; opacity: 0.9;">Overdue Days</span></th>
                </tr>
              </thead>
              <tbody>
        `;
        
        // 添加数据行
        processData.forEach((record, recordIndex) => {
          const rowStyle = recordIndex % 2 === 0 ? 'background-color: #f8f9fa;' : 'background-color: #ffffff;';
          
          // 根据超期天数显示不同颜色
          let overdueDisplay = '';
          if (record.overdueDays >= CONFIG.OVERDUE_DAYS) {
            overdueDisplay = `
              <div style="display: inline-block; text-align: center;">
                <div style="background: linear-gradient(135deg, #e74c3c, #c0392b); color: white; padding: 6px 12px; border-radius: 16px; font-weight: 600; font-size: 12px; box-shadow: 0 2px 6px rgba(231, 76, 60, 0.3); display: inline-block; min-width: 80px;">
                  <span style="display: block;">[超期] ${record.overdueDays}天</span>
                  <span style="display: block; font-size: 10px; opacity: 0.9; font-weight: 400;">Days</span>
                </div>
              </div>
            `;
          } else {
            overdueDisplay = `
              <div style="display: inline-block; text-align: center;">
                <div style="display: inline-block; text-align: center;">
                  <div style="background: linear-gradient(135deg, #f39c12, #e67e22); color: white; padding: 6px 12px; border-radius: 16px; font-weight: 600; font-size: 12px; box-shadow: 0 2px 6px rgba(243, 156, 18, 0.3); display: inline-block; min-width: 80px;">
                    <span style="display: block;">${record.overdueDays}天</span>
                    <span style="display: block; font-size: 10px; opacity: 0.9; font-weight: 400;">Days</span>
                  </div>
                </div>
              </div>
            `;
          }
          
          emailBody += `
            <tr style="${rowStyle}">
              <td style="padding: 12px; border-bottom: 1px solid #e9ecef; font-weight: 500; color: #2c3e50;">${record.id || ''}</td>
              <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; font-family: monospace;">${record.machineId || ''}</td>
              <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; max-width: 200px; word-wrap: break-word;">${record.description || ''}</td>
              <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; font-family: monospace;">${record.workshop || ''}</td>
              <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; font-family: monospace;">${formatDate(record.assignDate) || ''}</td>
              <td style="padding: 12px; border-bottom: 1px solid #e9ecef; color: #34495e; text-align: center;">
                ${overdueDisplay}
              </td>
            </tr>
          `;
        });
        
        emailBody += `
              </tbody>
            </table>
          </div>
          <div style="margin-top: 15px; text-align: right;">
            <span style="color: #7f8c8d; font-size: 14px;">
              共 ${processData.length} 条故障报告 | Total: ${processData.length} failure reports
            </span>
          </div>
        </div>
        `;
      });
    }
    
    // 如果没有故障报告，显示完成信息
    if (totalReports === 0) {
      const successColor = isOverdueReminder ? '#d4edda' : '#d4edda';
      const successBorderColor = isOverdueReminder ? '#c3e6cb' : '#c3e6cb';
      const successTextColor = isOverdueReminder ? '#155724' : '#155724';
      
      emailBody += `
        <div style="background-color: ${successColor}; border: 1px solid ${successBorderColor}; border-radius: 8px; padding: 20px; text-align: center; margin-bottom: 20px;">
          <p style="color: ${successTextColor}; font-size: 16px; margin: 0; font-weight: 500;">
            [完成] 所有故障报告都已上传完成！<br>
            <span style="font-size: 0.9em; color: ${successTextColor}; opacity: 0.8;">All failure reports have been uploaded!</span>
          </p>
        </div>
      `;
    }
    
    // 邮件底部
    const footerColor = isOverdueReminder ? '#d32f2f' : '#7f8c8d';
    const footerBgColor = isOverdueReminder ? '#ffebee' : '#ffffff';
    const footerBorderColor = isOverdueReminder ? '#f44336' : '#3498db';
    
    emailBody += `
      <div style="background-color: ${footerBgColor}; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); padding: 30px;">
        <div style="text-align: center; color: ${footerColor}; font-size: 14px; line-height: 1.6;">
          <p style="margin-bottom: 10px; font-weight: 600;">
            ${isOverdueReminder ? '[重要] 请立即处理这些超期故障报告！' : '请及时查看并处理相关故障报告。'}<br>
            <span style="font-size: 0.9em; opacity: 0.8;">
              ${isOverdueReminder ? 'Please handle these overdue failure reports immediately!' : 'Please review and handle related failure reports promptly.'}
            </span>
          </p>
          <p style="margin: 0; font-style: italic;">
            此邮件由系统自动发送，请勿回复。<br>
            <span style="font-size: 0.8em; opacity: 0.7;">This email is automatically sent by the system, please do not reply.</span>
          </p>
        </div>
      </div>
    </div>
    `;
    
    return emailBody;
    
  } catch (error) {
    console.error('生成合并邮件内容时出错:', error);
    return '<p>邮件内容生成失败，请检查系统日志。</p>';
  }
}

/**
 * 获取字段索引
 */
function getFieldIndexes(headers) {
  const indexes = {};
  headers.forEach((header, index) => {
    indexes[header] = index;
  });
  return indexes;
}

/**
 * 计算超期天数
 */
function calculateOverdueDays(assignDate) {
  try {
    if (!assignDate || !(assignDate instanceof Date)) {
      return 0;
    }
    
    const today = new Date();
    const timeDiff = today.getTime() - assignDate.getTime();
    const daysDiff = Math.ceil(timeDiff / (1000 * 3600 * 24));
    
    return Math.max(0, daysDiff);
    
  } catch (error) {
    console.error('计算超期天数时出错:', error);
    return 0;
  }
}

/**
 * 格式化日期
 */
function formatDate(date) {
  try {
    if (!date || !(date instanceof Date)) {
      return '';
    }
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    return `${year}-${month}-${day}`;
    
  } catch (error) {
    console.error('格式化日期时出错:', error);
    return '';
  }
}

/**
 * 记录系统日志
 */
function logSystemActivity(action, message, level = 'INFO') {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let logSheet = spreadsheet.getSheetByName(CONFIG.LOG_SHEET_NAME);
    
    // 如果日志表不存在，创建它
    if (!logSheet) {
      logSheet = spreadsheet.insertSheet(CONFIG.LOG_SHEET_NAME);
      logSheet.appendRow(['时间', '操作', '消息', '级别']);
    }
    
    const timestamp = new Date();
    logSheet.appendRow([timestamp, action, message, level]);
    
  } catch (error) {
    console.error('记录系统日志时出错:', error);
  }
}

/**
 * 测试函数：手动执行每周提醒
 */
function testWeeklyReminder() {
  weeklyFailureReportReminder();
}

/**
 * 测试函数：手动执行每日超期提醒
 */
function testDailyOverdueReminder() {
  dailyOverdueFailureReportReminder();
}

/**
 * 测试函数：获取故障报告数据
 */
function testGetFailureData() {
  const data = getFailureReportData();
  console.log('故障报告数据:', data);
  return data;
}

/**
 * 测试函数：获取通知邮箱
 */
function testGetNotificationEmails() {
  const emails = getNotificationEmails();
  console.log('通知邮箱:', emails);
  return emails;
}

/**
 * 测试函数：验证数据映射
 */
function testDataMapping() {
  try {
    console.log('=== 开始测试数据映射 ===');
    
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(CONFIG.FAILURE_SHEET_NAME);
    
    if (!sheet) {
      console.error('❌ 故障报告数据表未找到');
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.log('⚠️ 故障报告数据表为空');
      return;
    }
    
    // 显示表头
    const headers = data[0];
    console.log('📋 表头字段:', headers);
    
    // 显示字段索引
    const fieldIndexes = getFieldIndexes(headers);
    console.log('🔍 字段索引映射:', fieldIndexes);
    
    // 验证关键字段是否存在
    const requiredFields = ['编号', '机台号', '问题描述', '提交日期', '车间', '工序', '故障报告编号', '分配日期', '上传日期', '附件'];
    const missingFields = requiredFields.filter(field => fieldIndexes[field] === undefined);
    
    if (missingFields.length > 0) {
      console.error('❌ 缺少必要字段:', missingFields);
    } else {
      console.log('✅ 所有必要字段都存在');
    }
    
    // 显示前3行数据用于验证
    console.log('\n📊 前3行数据验证:');
    for (let i = 1; i <= Math.min(3, data.length - 1); i++) {
      const row = data[i];
      console.log(`\n第${i}行:`);
      console.log('  编号:', row[fieldIndexes.编号]);
      console.log('  机台号:', row[fieldIndexes.机台号]);
      console.log('  问题描述:', row[fieldIndexes.问题描述]);
      console.log('  车间:', row[fieldIndexes.车间]);
      console.log('  工序:', row[fieldIndexes.工序]);
      console.log('  分配日期:', row[fieldIndexes.分配日期]);
    }
    
    console.log('\n=== 数据映射测试完成 ===');
    
  } catch (error) {
    console.error('❌ 测试数据映射时出错:', error);
  }
}

/**
 * 判断故障报告是否已上传
 * @param {Object} record - 故障报告记录
 * @returns {Object} 包含上传状态和详细信息
 */
function isFailureReportUploaded(record) {
  try {
    if (!record) {
      return {
        isUploaded: false,
        reason: '记录为空',
        uploadDate: null,
        attachment: null
      };
    }
    
    const hasUploadDate = record.uploadDate && record.uploadDate.toString().trim() !== '';
    const hasAttachment = record.attachment && record.attachment.toString().trim() !== '';
    
    // 判断上传状态
    let isUploaded = false;
    let reason = '';
    
    if (hasUploadDate && hasAttachment) {
      isUploaded = true;
      reason = '已上传 - 有上传日期和附件';
    } else if (hasUploadDate && !hasAttachment) {
      isUploaded = true;
      reason = '已上传 - 有上传日期但无附件';
    } else if (!hasUploadDate && hasAttachment) {
      isUploaded = true;
      reason = '已上传 - 有附件但无上传日期';
    } else {
      isUploaded = false;
      reason = '未上传 - 无上传日期和附件';
    }
    
    return {
      isUploaded: isUploaded,
      reason: reason,
      uploadDate: record.uploadDate || null,
      attachment: record.attachment || null
    };
    
  } catch (error) {
    console.error('判断故障报告上传状态时出错:', error);
    return {
      isUploaded: false,
      reason: '判断出错',
      uploadDate: null,
      attachment: null
    };
  }
}

/**
 * 获取未上传的故障报告
 * @param {Array} failureData - 故障报告数据数组
 * @returns {Array} 未上传的故障报告数组
 */
function getUnuploadedFailureReports(failureData) {
  try {
    if (!failureData || !Array.isArray(failureData)) {
      return [];
    }
    
    const unuploadedReports = failureData.filter(record => {
      const uploadStatus = isFailureReportUploaded(record);
      return !uploadStatus.isUploaded;
    });
    
    console.log(`📊 找到 ${unuploadedReports.length} 条未上传的故障报告`);
    return unuploadedReports;
    
  } catch (error) {
    console.error('获取未上传故障报告时出错:', error);
    return [];
  }
}

/**
 * 获取已上传的故障报告
 * @param {Array} failureData - 故障报告数据数组
 * @returns {Array} 已上传的故障报告数组
 */
function getUploadedFailureReports(failureData) {
  try {
    if (!failureData || !Array.isArray(failureData)) {
      return [];
    }
    
    const uploadedReports = failureData.filter(record => {
      const uploadStatus = isFailureReportUploaded(record);
      return uploadStatus.isUploaded;
    });
    
    console.log(`📊 找到 ${uploadedReports.length} 条已上传的故障报告`);
    return uploadedReports;
    
  } catch (error) {
    console.error('获取已上传故障报告时出错:', error);
    return [];
  }
}

/**
 * 测试函数：测试故障报告上传状态判断
 */
function testUploadStatus() {
  try {
    console.log('=== 开始测试故障报告上传状态判断 ===');
    
    // 获取故障报告数据
    const failureData = getFailureReportData();
    if (!failureData || failureData.length === 0) {
      console.log('⚠️ 没有故障报告数据');
      return;
    }
    
    console.log(`📊 总共有 ${failureData.length} 条故障报告`);
    
    // 分析每条记录的上传状态
    console.log('\n📋 各记录上传状态分析:');
    failureData.forEach((record, index) => {
      const uploadStatus = isFailureReportUploaded(record);
      console.log(`\n${index + 1}. 编号: ${record.id}`);
      console.log(`   状态: ${uploadStatus.isUploaded ? '✅ 已上传' : '❌ 未上传'}`);
      console.log(`   原因: ${uploadStatus.reason}`);
      console.log(`   上传日期: ${uploadStatus.uploadDate || '无'}`);
      console.log(`   附件: ${uploadStatus.attachment || '无'}`);
    });
    
    // 统计上传状态
    const uploadedReports = getUploadedFailureReports(failureData);
    const unuploadedReports = getUnuploadedFailureReports(failureData);
    
    console.log('\n📊 上传状态统计:');
    console.log(`✅ 已上传: ${uploadedReports.length} 条`);
    console.log(`❌ 未上传: ${unuploadedReports.length} 条`);
    console.log(`📧 总计: ${failureData.length} 条`);
    
    // 显示未上传的故障报告详情
    if (unuploadedReports.length > 0) {
      console.log('\n🚨 未上传的故障报告详情:');
      unuploadedReports.forEach((record, index) => {
        console.log(`${index + 1}. ${record.id} - ${record.description} (${record.workshop}/${record.process})`);
      });
    }
    
    console.log('\n=== 故障报告上传状态判断测试完成 ===');
    
  } catch (error) {
    console.error('❌ 测试故障报告上传状态判断时出错:', error);
  }
}

/**
 * 测试函数：测试邮件内容生成（只显示未上传的故障报告）
 */
function testEmailContentGeneration() {
  try {
    console.log('=== 开始测试邮件内容生成 ===');
    
    // 获取故障报告数据
    const failureData = getFailureReportData();
    if (!failureData || failureData.length === 0) {
      console.log('⚠️ 没有故障报告数据');
      return;
    }
    
    console.log(`📊 总共有 ${failureData.length} 条故障报告`);
    
    // 测试每周提醒邮件内容生成
    console.log('\n📧 测试每周提醒邮件内容生成:');
    const weeklyContent = generateWeeklyEmailContent(failureData);
    console.log(`✅ 每周提醒邮件内容生成成功，长度: ${weeklyContent.length} 字符`);
    
    // 测试超期提醒邮件内容生成
    console.log('\n📧 测试超期提醒邮件内容生成:');
    const overdueContent = generateOverdueEmailContent(failureData);
    console.log(`✅ 超期提醒邮件内容生成成功，长度: ${overdueContent.length} 字符`);
    
    // 测试筛选功能
    console.log('\n🔍 测试筛选功能:');
    const testWorkshop = 'TB1';
    const testProcess = 'IM';
    const filteredData = filterFailureDataByWorkshopAndProcess(testWorkshop, testProcess);
    console.log(`✅ 筛选结果: 车间 ${testWorkshop}, 工序 ${testProcess} 找到 ${filteredData.length} 条未上传的故障报告`);
    
    if (filteredData.length > 0) {
      console.log('📋 筛选到的故障报告:');
      filteredData.forEach((record, index) => {
        console.log(`  ${index + 1}. ${record.id} - ${record.description}`);
      });
    }
    
    console.log('\n=== 邮件内容生成测试完成 ===');
    
  } catch (error) {
    console.error('❌ 测试邮件内容生成时出错:', error);
  }
}

/**
 * 测试函数：测试超期故障报告筛选逻辑
 */
function testOverdueFiltering() {
  try {
    console.log('=== 开始测试超期故障报告筛选逻辑 ===');
    
    // 获取故障报告数据
    const failureData = getFailureReportData();
    if (!failureData || failureData.length === 0) {
      console.log('⚠️ 没有故障报告数据');
      return;
    }
    
    console.log(`📊 总共有 ${failureData.length} 条故障报告`);
    
    // 测试超期筛选
    console.log('\n🔍 测试超期筛选逻辑:');
    const overdueData = getOverdueFailureReportData();
    console.log(`✅ 超期筛选结果: 找到 ${overdueData.length} 条超期 >= ${CONFIG.OVERDUE_DAYS} 天且未上传的故障报告`);
    
    if (overdueData.length > 0) {
      console.log('\n📋 超期未上传故障报告详情:');
      overdueData.forEach((record, index) => {
        console.log(`  ${index + 1}. ${record.id} - ${record.description}`);
        console.log(`     车间/工序: ${record.workshop}/${record.process}`);
        console.log(`     分配日期: ${record.assignDate}`);
        console.log(`     超期天数: ${record.overdueDays}天`);
        console.log(`     上传状态: ${isFailureReportUploaded(record).isUploaded ? '已上传' : '未上传'}`);
        console.log('');
      });
    }
    
    // 验证筛选逻辑
    console.log('\n🔍 验证筛选逻辑:');
    const totalOverdue = failureData.filter(record => record.overdueDays >= CONFIG.OVERDUE_DAYS);
    const uploadedOverdue = totalOverdue.filter(record => isFailureReportUploaded(record).isUploaded);
    const unuploadedOverdue = totalOverdue.filter(record => !isFailureReportUploaded(record).isUploaded);
    
    console.log(`📊 筛选逻辑验证:`);
    console.log(`  总超期记录: ${totalOverdue.length} 条`);
    console.log(`  已上传超期: ${uploadedOverdue.length} 条 (已排除)`);
    console.log(`  未上传超期: ${unuploadedOverdue.length} 条 (需要提醒)`);
    console.log(`  函数筛选结果: ${overdueData.length} 条`);
    console.log(`  筛选逻辑一致: ${unuploadedOverdue.length === overdueData.length ? '✅ 是' : '❌ 否'}`);
    
    // 显示已上传但超期的故障报告（用于验证排除逻辑）
    if (uploadedOverdue.length > 0) {
      console.log('\n📋 已上传但超期的故障报告 (已排除):');
      uploadedOverdue.forEach((record, index) => {
        console.log(`  ${index + 1}. ${record.id} - ${record.description}`);
        console.log(`     车间/工序: ${record.workshop}/${record.process}`);
        console.log(`     分配日期: ${record.assignDate}`);
        console.log(`     超期天数: ${record.overdueDays}天`);
        console.log(`     上传状态: 已上传 (已排除)`);
        console.log('');
      });
    }
    
    console.log('\n=== 超期故障报告筛选逻辑测试完成 ===');
    
  } catch (error) {
    console.error('❌ 测试超期故障报告筛选逻辑时出错:', error);
  }
}

/**
 * 测试函数：测试日报邮件内容生成（验证修复后的逻辑）
 */
function testDailyEmailGeneration() {
  try {
    console.log('=== 开始测试日报邮件内容生成 ===');
    
    // 获取超期故障报告数据
    const overdueData = getOverdueFailureReportData();
    if (!overdueData || overdueData.length === 0) {
      console.log('⚠️ 没有超期且未上传的故障报告数据');
      return;
    }
    
    console.log(`📊 超期且未上传的故障报告: ${overdueData.length} 条`);
    
    // 测试日报邮件内容生成
    console.log('\n📧 测试日报邮件内容生成:');
    const emailContent = generateOverdueEmailContent(overdueData);
    console.log(`✅ 日报邮件内容生成成功，长度: ${emailContent.length} 字符`);
    
    // 验证邮件内容中是否包含了所有超期故障报告
    console.log('\n🔍 验证邮件内容:');
    overdueData.forEach((record, index) => {
      const recordInEmail = emailContent.includes(record.id);
      console.log(`  ${index + 1}. ${record.id} - ${record.description}: ${recordInEmail ? '✅ 包含在邮件中' : '❌ 未包含在邮件中'}`);
    });
    
    // 统计邮件中的记录数量
    const tableRows = (emailContent.match(/<tr style=/g) || []).length;
    console.log(`\n📊 邮件内容统计:`);
    console.log(`  超期故障报告数量: ${overdueData.length} 条`);
    console.log(`  邮件表格行数: ${tableRows} 行 (包含表头)`);
    console.log(`  数据行数: ${tableRows - 1} 行`);
    console.log(`  数据一致性: ${overdueData.length === (tableRows - 1) ? '✅ 是' : '❌ 否'}`);
    
    console.log('\n=== 日报邮件内容生成测试完成 ===');
    
  } catch (error) {
    console.error('❌ 测试日报邮件内容生成时出错:', error);
  }
}

/**
 * 测试函数：测试超期筛选逻辑（验证修复后的筛选）
 */
function testOverdueFilteringWithWorkshop() {
  try {
    console.log('=== 开始测试车间工序超期筛选逻辑 ===');
    
    // 测试TB1/IM组合的超期筛选
    const testWorkshop = 'TB1';
    const testProcess = 'IM';
    
    console.log(`🔍 测试筛选条件: 车间 ${testWorkshop}, 工序 ${testProcess}`);
    
    // 测试普通筛选（所有未上传的）
    console.log('\n📊 测试普通筛选 (所有未上传的):');
    const allUnuploaded = filterFailureDataByWorkshopAndProcess(testWorkshop, testProcess, false);
    console.log(`✅ 普通筛选结果: 找到 ${allUnuploaded.length} 条未上传的故障报告`);
    
    if (allUnuploaded.length > 0) {
      allUnuploaded.forEach((record, index) => {
        console.log(`  ${index + 1}. ${record.id} - ${record.description}`);
        console.log(`     超期天数: ${record.overdueDays}天`);
        console.log(`     上传状态: 未上传`);
        console.log('');
      });
    }
    
    // 测试超期筛选（只返回超期>=7天的）
    console.log('\n📊 测试超期筛选 (只返回超期>=7天的):');
    const overdueOnly = filterFailureDataByWorkshopAndProcess(testWorkshop, testProcess, true);
    console.log(`✅ 超期筛选结果: 找到 ${overdueOnly.length} 条超期>=${CONFIG.OVERDUE_DAYS}天且未上传的故障报告`);
    
    if (overdueOnly.length > 0) {
      overdueOnly.forEach((record, index) => {
        console.log(`  ${index + 1}. ${record.id} - ${record.description}`);
        console.log(`     超期天数: ${record.overdueDays}天`);
        console.log(`     上传状态: 未上传`);
        console.log('');
      });
    }
    
    // 验证筛选逻辑
    console.log('\n🔍 验证筛选逻辑:');
    const expectedOverdue = allUnuploaded.filter(record => record.overdueDays >= CONFIG.OVERDUE_DAYS);
    console.log(`📊 筛选逻辑验证:`);
    console.log(`  总未上传记录: ${allUnuploaded.length} 条`);
    console.log(`  预期超期记录: ${expectedOverdue.length} 条`);
    console.log(`  实际超期筛选: ${overdueOnly.length} 条`);
    console.log(`  筛选逻辑一致: ${expectedOverdue.length === overdueOnly.length ? '✅ 是' : '❌ 否'}`);
    
    // 显示被排除的未超期记录
    const excludedRecords = allUnuploaded.filter(record => record.overdueDays < CONFIG.OVERDUE_DAYS);
    if (excludedRecords.length > 0) {
      console.log('\n📋 被排除的未超期记录:');
      excludedRecords.forEach((record, index) => {
        console.log(`  ${index + 1}. ${record.id} - ${record.description}`);
        console.log(`     超期天数: ${record.overdueDays}天 (未达到${CONFIG.OVERDUE_DAYS}天阈值)`);
        console.log(`     上传状态: 未上传 (但未超期，已排除)`);
        console.log('');
      });
    }
    
    console.log('\n=== 车间工序超期筛选逻辑测试完成 ===');
    
  } catch (error) {
    console.error('❌ 测试车间工序超期筛选逻辑时出错:', error);
  }
}

/**
 * 测试函数：测试合并邮件功能
 */
function testMergedEmailGeneration() {
  try {
    console.log('=== 开始测试合并邮件功能 ===');
    
    // 获取通知邮箱列表
    const notificationEmails = getNotificationEmails();
    if (!notificationEmails || notificationEmails.length === 0) {
      console.log('⚠️ 没有找到通知邮箱地址');
      return;
    }
    
    console.log(`📧 找到 ${notificationEmails.length} 个通知邮箱:`);
    notificationEmails.forEach((email, index) => {
      console.log(`  ${index + 1}. ${email.email} (${email.workshop}/${email.process})`);
    });
    
    // 模拟邮件发送流程（不实际发送）
    console.log('\n🔍 模拟邮件发送流程:');
    
    // 收集所有有效的邮箱地址和故障报告数据
    const allEmails = [];
    const emailDetails = [];
    
    notificationEmails.forEach((notification, index) => {
      console.log(`\n📧 处理第 ${index + 1} 个邮箱: ${notification.email}`);
      
      // 测试普通筛选（所有未上传的）
      const filteredData = filterFailureDataByWorkshopAndProcess(notification.workshop, notification.process, false);
      
      if (filteredData && filteredData.length > 0) {
        allEmails.push(notification.email);
        emailDetails.push({
          email: notification.email,
          workshop: notification.workshop,
          process: notification.process,
          data: filteredData
        });
        
        console.log(`✅ 收集成功: ${notification.email} (${notification.workshop}/${notification.process})`);
        console.log(`📊 包含 ${filteredData.length} 条故障报告`);
      } else {
        console.log(`⏭️ 跳过: ${notification.email}，没有匹配的故障报告`);
      }
    });
    
    // 显示收集结果
    console.log('\n📊 收集结果统计:');
    console.log(`📧 有效邮箱: ${allEmails.length} 个`);
    console.log(`📧 邮箱列表: ${allEmails.join(', ')}`);
    
    if (emailDetails.length > 0) {
      console.log('\n📋 各邮箱故障报告详情:');
      emailDetails.forEach((detail, index) => {
        console.log(`\n${index + 1}. ${detail.email} (${detail.workshop}/${detail.process}):`);
        detail.data.forEach((record, recordIndex) => {
          console.log(`   ${recordIndex + 1}. ${record.id} - ${record.description} (超期${record.overdueDays}天)`);
        });
      });
      
      // 测试合并邮件内容生成
      console.log('\n📧 测试合并邮件内容生成:');
      const mergedContent = generateMergedEmailContent(emailDetails, '每周故障报告提醒');
      console.log(`✅ 合并邮件内容生成成功，长度: ${mergedContent.length} 字符`);
      
      // 验证邮件内容
      console.log('\n🔍 验证合并邮件内容:');
      emailDetails.forEach((detail, index) => {
        const hasWorkshopProcess = mergedContent.includes(`${detail.workshop}/${detail.process}`);
        const hasAllRecords = detail.data.every(record => mergedContent.includes(record.id));
        console.log(`  ${index + 1}. ${detail.email} (${detail.workshop}/${detail.process}):`);
        console.log(`     车间工序标题: ${hasWorkshopProcess ? '✅ 包含' : '❌ 未包含'}`);
        console.log(`     故障报告记录: ${hasAllRecords ? '✅ 全部包含' : '❌ 部分缺失'}`);
      });
      
      console.log('\n📊 合并邮件统计:');
      console.log(`📧 收件人数量: ${allEmails.length} 个`);
      console.log(`📧 发送邮件数量: 1 封`);
      console.log(`📧 总故障报告数量: ${emailDetails.reduce((sum, detail) => sum + detail.data.length, 0)} 条`);
      console.log(`📧 邮件大小: ${Math.round(mergedContent.length / 1024 * 100) / 100} KB`);
      
    } else {
      console.log('⚠️ 没有找到需要发送邮件的有效邮箱地址');
    }
    
    console.log('\n=== 合并邮件功能测试完成 ===');
    
  } catch (error) {
    console.error('❌ 测试合并邮件功能时出错:', error);
  }
}

/**
 * 测试函数：测试日报和周报的不同显示逻辑
 */
function testDailyWeeklyDisplayLogic() {
  try {
    console.log('=== 开始测试日报和周报的不同显示逻辑 ===');
    
    // 获取通知邮箱列表
    const notificationEmails = getNotificationEmails();
    if (!notificationEmails || notificationEmails.length === 0) {
      console.log('⚠️ 没有找到通知邮箱地址');
      return;
    }
    
    console.log(`📧 找到 ${notificationEmails.length} 个通知邮箱:`);
    notificationEmails.forEach((email, index) => {
      console.log(`  ${index + 1}. ${email.email} (${email.workshop}/${email.process})`);
    });
    
    // 模拟邮件发送流程（不实际发送）
    console.log('\n🔍 模拟邮件发送流程:');
    
    // 收集所有有效的邮箱地址和故障报告数据
    const allEmails = [];
    const emailDetails = [];
    
    notificationEmails.forEach((notification, index) => {
      console.log(`\n📧 处理第 ${index + 1} 个邮箱: ${notification.email}`);
      
      // 测试普通筛选（所有未上传的）
      const filteredData = filterFailureDataByWorkshopAndProcess(notification.workshop, notification.process, false);
      
      if (filteredData && filteredData.length > 0) {
        allEmails.push(notification.email);
        emailDetails.push({
          email: notification.email,
          workshop: notification.workshop,
          process: notification.process,
          data: filteredData
        });
        
        console.log(`✅ 收集成功: ${notification.email} (${notification.workshop}/${notification.process})`);
        console.log(`📊 包含 ${filteredData.length} 条故障报告`);
      } else {
        console.log(`⏭️ 跳过: ${notification.email}，没有匹配的故障报告`);
      }
    });
    
    if (emailDetails.length > 0) {
      // 测试日报邮件内容生成（超期提醒）
      console.log('\n📧 测试日报邮件内容生成 (超期提醒):');
      const dailyContent = generateMergedEmailContent(emailDetails, '超期故障报告紧急提醒');
      console.log(`✅ 日报邮件内容生成成功，长度: ${dailyContent.length} 字符`);
      
      // 验证日报内容：应该按车间/工序分组显示
      console.log('\n🔍 验证日报内容结构:');
      const dailyWorkshopProcessCount = (dailyContent.match(/\[详情\] .*\/.* - 未上传故障报告详情/g) || []).length;
      console.log(`📊 日报车间/工序分组数量: ${dailyWorkshopProcessCount} 个`);
      console.log(`📊 预期分组数量: ${emailDetails.length} 个`);
      console.log(`📊 日报结构正确: ${dailyWorkshopProcessCount === emailDetails.length ? '✅ 是' : '❌ 否'}`);
      
      // 测试周报邮件内容生成（每周提醒）
      console.log('\n📧 测试周报邮件内容生成 (每周提醒):');
      const weeklyContent = generateMergedEmailContent(emailDetails, '每周故障报告提醒');
      console.log(`✅ 周报邮件内容生成成功，长度: ${weeklyContent.length} 字符`);
      
      // 验证周报内容：应该按工序分组显示
      console.log('\n🔍 验证周报内容结构:');
      const weeklyProcessCount = (weeklyContent.match(/\[详情\] .* 工序 - 未上传故障报告详情/g) || []).length;
      
      // 统计不同的工序数量
      const uniqueProcesses = [...new Set(emailDetails.map(detail => detail.process))];
      console.log(`📊 周报工序分组数量: ${weeklyProcessCount} 个`);
      console.log(`📊 不同工序数量: ${uniqueProcesses.length} 个`);
      console.log(`📊 工序列表: ${uniqueProcesses.join(', ')}`);
      console.log(`📊 周报结构正确: ${weeklyProcessCount === uniqueProcesses.length ? '✅ 是' : '❌ 否'}`);
      
      // 显示各工序的故障报告数量
      console.log('\n📋 各工序故障报告统计:');
      uniqueProcesses.forEach(process => {
        const processData = emailDetails.filter(detail => detail.process === process);
        const totalRecords = processData.reduce((sum, detail) => sum + detail.data.length, 0);
        console.log(`  ${process}: ${totalRecords} 条故障报告`);
      });
      
      // 统计邮件内容
      console.log('\n📊 邮件内容统计:');
      console.log(`📧 日报邮件大小: ${Math.round(dailyContent.length / 1024 * 100) / 100} KB`);
      console.log(`📧 周报邮件大小: ${Math.round(weeklyContent.length / 1024 * 100) / 100} KB`);
      console.log(`📧 日报表格数量: ${dailyWorkshopProcessCount} 个`);
      console.log(`📧 周报表格数量: ${weeklyProcessCount} 个`);
      
    } else {
      console.log('⚠️ 没有找到需要发送邮件的有效邮箱地址');
    }
    
    console.log('\n=== 日报和周报显示逻辑测试完成 ===');
    
  } catch (error) {
    console.error('❌ 测试日报和周报显示逻辑时出错:', error);
  }
}
