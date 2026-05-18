/**
 * 故障报告邮件提醒定时执行模块
 * Failure Report Email Reminder Timing Module
 * 
 * 功能说明：
 * 1. 设置每周故障报告提醒的定时任务
 * 2. 设置每日超期故障报告提醒的定时任务
 * 3. 支持灵活的定时配置
 * 4. 参考 Changeover Process Para Verification 项目的定时执行机制
 */

// 全局配置
const TIMING_CONFIG = {
  WEEKLY_REMINDER: {
    function: 'weeklyFailureReportReminder',
    description: '每周故障报告提醒',
    defaultTime: '09:00', // 默认每周一上午9点
    defaultWeek: 1 // 默认每周一执行
  },
  DAILY_OVERDUE_REMINDER: {
    function: 'dailyOverdueFailureReportReminder',
    description: '每日超期故障报告提醒',
    defaultTime: '08:30', // 默认每天上午8:30
    defaultWeek: -1 // -1表示每天执行
  }
};

/******辅助函数：计算日期是本月第几周******/
function getWeekInMonth(date) {
  try {
    // 获取日期是本月第几周
    let firstDayOfMonth = new Date(date.getFullYear(), date.getMonth(), 1);
    let firstDayWeek = firstDayOfMonth.getDay(); // 0-6 (周日-周六)
    let dayOfMonth = date.getDate();
    
    // 计算是本月第几周
    let weekInMonth = Math.ceil((dayOfMonth + firstDayWeek) / 7);
    return weekInMonth;
  } catch (error) {
    console.error('计算周内日期时出错:', error);
    return 1; // 默认返回第1周
  }
}

/******辅助函数：格式化日期时间为 HH:MM:SS 格式******/
function formatVariableAsDateHms(date) {
  try {
    if (!date || !(date instanceof Date)) {
      return '无效日期';
    }
    
    let hours = date.getHours().toString().padStart(2, '0');
    let minutes = date.getMinutes().toString().padStart(2, '0');
    let seconds = date.getSeconds().toString().padStart(2, '0');
    
    return `${hours}:${minutes}:${seconds}`;
  } catch (error) {
    console.error('格式化日期时间时出错:', error);
    return '格式化错误';
  }
}

/******取消所有定时任务******/
function cancelAllTimeDrivenTriggers(){
  let triggers=ScriptApp.getProjectTriggers();
  console.log(triggers.length);
  for(let i=triggers.length-1;i>-1;i--){
    if (triggers[i].getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    // ScriptApp.deleteTrigger(triggers[i]);
  }
}

/******每天运行一次精准时间触发任务******/
function timeExec() {
  cancelAllTimeDrivenTriggers();
  let setHour=12;
  let setMinute=30;
  ScriptApp.newTrigger('createSpecialTimeDrivenTriggers')
    .timeBased()
    .everyDays(1) // 每天触发
    .atHour(setHour) // 指定在每天的设定小时触发
    .nearMinute(setMinute) // 尽可能接近每天设定小时:分钟
    .create();
  createSpecialTimeDrivenTriggers();
}

function createSpecialTimeDrivenTriggers() {
  let sbnTimingSetLr=sbnTimingSet.getLastRow();
  let arrSetTime=[];
  if(sbnTimingSetLr>2){arrSetTime=sbnTimingSet.getRange(3,1,sbnTimingSetLr-2,7).getDisplayValues().filter(v=>{return v[0]!=""&&v[1]!=""&&v[2]!=""&&v[3]!=""&&v[4]!=""&&v[5]!=""&&v[6]!=""})}
  if(arrSetTime.length>0){
    // 获取当前日期，但不对时间进行设置
    let now = new Date();
    let nowTime=now.getTime();
    let nowYear = now.getFullYear();
    let nowMonth = now.getMonth();
    let nowDay = now.getDate();
    let strDes="";
    let strFunction="";
    let arrSetTimeRow=[];
    let hourSetTimeRow=[];
    let minuteSetTimeRow=[];
    let specificTime = new Date(nowTime);
    let nowOver30Minute=new Date(nowTime+0.5*3600*1000);
    let arrMonth=[];
    let arrWeek=[];
    let arrWeekInMonth=[];
    let arrDay=[];
    let setExec=new Date(nowTime);
    let month_setExec=setExec.getMonth()+1;
    let week_setExec=setExec.getDay();
    let weekInMonth_setExec=getWeekInMonth(setExec);
    let day_setExec=setExec.getDate();
    for(let i=0;i<arrSetTime.length;i++){
      arrMonth=JSON.parse(arrSetTime[i][0]);
      arrWeek=JSON.parse(arrSetTime[i][1]);
      arrWeekInMonth=JSON.parse(arrSetTime[i][2]);
      arrDay=JSON.parse(arrSetTime[i][3]);
      strDes=arrSetTime[i][4];
      strFunction=arrSetTime[i][6];
      // 取消之前的定时任务
      cancelAssignstrFunction(strFunction);
      arrSetTimeRow=arrSetTime[i][4].split("|");
      for(let j=0;j<arrSetTimeRow.length;j++){
        hourSetTimeRow=Number(arrSetTimeRow[j].split(":")[0]);
        minuteSetTimeRow=Number(arrSetTimeRow[j].split(":")[1]);
        if(hourSetTimeRow>=0&&minuteSetTimeRow>=0){
          // 对每个时间点创建全新的日期对象
          specificTime = new Date(nowYear, nowMonth, nowDay, hourSetTimeRow, minuteSetTimeRow, 0);
          if(nowOver30Minute>specificTime){
            specificTime = new Date(nowYear, nowMonth, nowDay+1, hourSetTimeRow, minuteSetTimeRow, 0);
          }
          month_setExec=specificTime.getMonth()+1;
          week_setExec=specificTime.getDay();
          weekInMonth_setExec=getWeekInMonth(specificTime);
          day_setExec=specificTime.getDate();
          console.log({
            "当前月":(month_setExec)+"；是否在设置里："+arrMonth.indexOf(month_setExec),
            "当前星期":week_setExec+"；是否在设置里："+arrWeek.indexOf(week_setExec),
            "当前周In月":weekInMonth_setExec+"；是否在设置里："+arrWeekInMonth.indexOf(weekInMonth_setExec),
            "当前日":day_setExec+"；是否在设置里："+arrDay.indexOf(day_setExec),
            "程序名称:":strDes,
            "函数名：":strFunction,
            "触发时间：":formatVariableAsDateHms(specificTime)
          })
          if(arrMonth.indexOf(month_setExec)!=-1&&arrWeek.indexOf(week_setExec)!=-1&&arrWeekInMonth.indexOf(weekInMonth_setExec)!=-1&&arrDay.indexOf(day_setExec)!=-1){
            // 创建定时器
            ScriptApp.newTrigger(strFunction)
                .timeBased()
                .at(specificTime)
                .create();
          }
        }
      }
    }
  }
}

/******取消指定定时任务******/
function cancelAssignstrFunction(strFunction){
  let triggers=ScriptApp.getProjectTriggers();
  for(let i=triggers.length-1;i>-1;i--){
    console.log(triggers[i].getHandlerFunction())
    if(triggers[i].getHandlerFunction()==strFunction){
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * 设置故障报告邮件提醒的定时任务
 */
function setupFailureReportTiming() {
  try {
    console.log('=== 开始设置故障报告邮件提醒定时任务 ===');
    
    // 取消之前的定时任务
    cancelAllTimeDrivenTriggers();
    
    // 设置每周提醒定时任务
    setupWeeklyReminderTrigger();
    
    // 设置每日超期提醒定时任务
    setupDailyOverdueReminderTrigger();
    
    console.log('=== 故障报告邮件提醒定时任务设置完成 ===');
    
  } catch (error) {
    console.error('设置故障报告邮件提醒定时任务时出错:', error);
  }
}

/**
 * 设置每周提醒定时任务
 */
function setupWeeklyReminderTrigger() {
  try {
    const config = TIMING_CONFIG.WEEKLY_REMINDER;
    const timeParts = config.defaultTime.split(':');
    const hour = parseInt(timeParts[0]);
    const minute = parseInt(timeParts[1]);
    
    // 创建每周一执行的定时任务
    ScriptApp.newTrigger(config.function)
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY) // 每周一
      .atHour(hour)
      .nearMinute(minute)
      .create();
    
    console.log(`✅ 已设置每周提醒定时任务: ${config.description} - 每周一 ${config.defaultTime}`);
    
  } catch (error) {
    console.error('设置每周提醒定时任务时出错:', error);
  }
}

/**
 * 设置每日超期提醒定时任务
 */
function setupDailyOverdueReminderTrigger() {
  try {
    const config = TIMING_CONFIG.DAILY_OVERDUE_REMINDER;
    const timeParts = config.defaultTime.split(':');
    const hour = parseInt(timeParts[0]);
    const minute = parseInt(timeParts[1]);
    
    // 创建每日执行的定时任务
    ScriptApp.newTrigger(config.function)
      .timeBased()
      .everyDays(1) // 每天执行
      .atHour(hour)
      .nearMinute(minute)
      .create();
    
    console.log(`✅ 已设置每日超期提醒定时任务: ${config.description} - 每天 ${config.defaultTime}`);
    
  } catch (error) {
    console.error('设置每日超期提醒定时任务时出错:', error);
  }
}

/**
 * 高级定时配置设置
 * 支持按月份、星期、周内日期、具体日期等条件设置
 */
function setupAdvancedTiming() {
  try {
    console.log('=== 开始设置高级定时配置 ===');
    
    // 取消之前的定时任务
    cancelAllTimeDrivenTriggers();
    
    // 获取当前时间
    const now = new Date();
    const currentTime = now.getTime();
    const currentYear = now.getFullYear();
    const currentMonth = now.getMonth();
    const currentDay = now.getDate();
    
    // 设置每周提醒（每周一上午9点）
    const weeklyTime = new Date(currentYear, currentMonth, currentDay, 9, 0, 0);
    if (weeklyTime.getTime() <= currentTime) {
      weeklyTime.setDate(weeklyTime.getDate() + 7); // 如果当前时间已过，设置为下周
    }
    
    ScriptApp.newTrigger('weeklyFailureReportReminder')
      .timeBased()
      .at(weeklyTime)
      .create();
    
    console.log(`✅ 已设置每周提醒定时任务: ${formatVariableAsDateHms(weeklyTime)}`);
    
    // 设置每日超期提醒（每天上午8:30）
    const dailyTime = new Date(currentYear, currentMonth, currentDay, 8, 30, 0);
    if (dailyTime.getTime() <= currentTime) {
      dailyTime.setDate(dailyTime.getDate() + 1); // 如果当前时间已过，设置为明天
    }
    
    ScriptApp.newTrigger('dailyOverdueFailureReportReminder')
      .timeBased()
      .at(dailyTime)
      .create();
    
    console.log(`✅ 已设置每日超期提醒定时任务: ${formatVariableAsDateHms(dailyTime)}`);
    
    console.log('=== 高级定时配置设置完成 ===');
    
  } catch (error) {
    console.error('设置高级定时配置时出错:', error);
  }
}

/**
 * 手动执行定时任务设置
 */
function manualSetupTiming() {
  try {
    console.log('=== 手动设置故障报告邮件提醒定时任务 ===');
    
    // 使用高级定时配置
    setupAdvancedTiming();
    
    // 显示当前定时任务状态
    showCurrentTriggers();
    
  } catch (error) {
    console.error('手动设置定时任务时出错:', error);
  }
}

/**
 * 显示当前定时任务状态
 */
function showCurrentTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    console.log(`\n=== 当前定时任务状态 ===`);
    console.log(`总数量: ${triggers.length}`);
    
    if (triggers.length === 0) {
      console.log('暂无定时任务');
      return;
    }
    
    triggers.forEach((trigger, index) => {
      if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
        console.log(`${index + 1}. 函数: ${trigger.getHandlerFunction()}`);
        console.log(`   类型: 定时任务`);
        console.log(`   创建时间: ${trigger.getDateCreated()}`);
        console.log(`   下次执行: ${trigger.getNextRun()}`);
        console.log('---');
      }
    });
    
  } catch (error) {
    console.error('显示当前定时任务状态时出错:', error);
  }
}

/**
 * 测试定时任务设置
 */
function testTimingSetup() {
  try {
    console.log('=== 测试定时任务设置 ===');
    
    // 设置定时任务
    setupAdvancedTiming();
    
    // 显示状态
    showCurrentTriggers();
    
    // 测试执行一次
    console.log('\n=== 测试执行 ===');
    console.log('测试每周提醒...');
    weeklyFailureReportReminder();
    
    console.log('测试每日超期提醒...');
    dailyOverdueFailureReportReminder();
    
    console.log('=== 测试完成 ===');
    
  } catch (error) {
    console.error('测试定时任务设置时出错:', error);
  }
}

/**
 * 重置定时任务配置
 */
function resetTimingConfig() {
  try {
    console.log('=== 重置定时任务配置 ===');
    
    // 取消所有定时任务
    cancelAllTimeDrivenTriggers();
    
    // 重新设置默认配置
    setupFailureReportTiming();
    
    console.log('定时任务配置已重置');
    
  } catch (error) {
    console.error('重置定时任务配置时出错:', error);
  }
}

/**
 * 获取定时任务统计信息
 */
function getTimingStatistics() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const clockTriggers = triggers.filter(t => t.getEventType() === ScriptApp.EventType.CLOCK);
    
    const stats = {
      totalTriggers: triggers.length,
      clockTriggers: clockTriggers.length,
      weeklyReminder: clockTriggers.filter(t => t.getHandlerFunction() === 'weeklyFailureReportReminder').length,
      dailyOverdueReminder: clockTriggers.filter(t => t.getHandlerFunction() === 'dailyOverdueFailureReportReminder').length
    };
    
    console.log('=== 定时任务统计信息 ===');
    console.log(`总任务数: ${stats.totalTriggers}`);
    console.log(`定时任务数: ${stats.clockTriggers}`);
    console.log(`每周提醒任务数: ${stats.weeklyReminder}`);
    console.log(`每日超期提醒任务数: ${stats.dailyOverdueReminder}`);
    
    return stats;
    
  } catch (error) {
    console.error('获取定时任务统计信息时出错:', error);
    return null;
  }
}

/**
 * 主函数：初始化定时任务
 */
function initializeTimingSystem() {
  try {
    console.log('=== 初始化故障报告邮件提醒定时系统 ===');
    
    // 设置基本定时任务
    setupFailureReportTiming();
    
    // 显示当前状态
    showCurrentTriggers();
    
    console.log('=== 定时系统初始化完成 ===');
    
  } catch (error) {
    console.error('初始化定时系统时出错:', error);
  }
}

/**
 * 快速设置函数：一键设置所有定时任务
 */
function quickSetup() {
  try {
    console.log('=== 快速设置故障报告邮件提醒系统 ===');
    
    // 设置定时任务
    setupAdvancedTiming();
    
    // 显示状态
    showCurrentTriggers();
    
    console.log('=== 快速设置完成 ===');
    console.log('系统将在以下时间自动执行：');
    console.log('- 每周一上午9:00：每周故障报告提醒');
    console.log('- 每天上午8:30：每日超期故障报告提醒');
    
  } catch (error) {
    console.error('快速设置时出错:', error);
  }
}
