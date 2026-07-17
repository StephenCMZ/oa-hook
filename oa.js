// ==UserScript==
// @name         OA 系统
// @namespace    https://github.com/StephenCMZ/oa-hook.git
// @version      0.8.7
// @description  OA 系统
// @author       StephenChen
// @match        http://oa.gdytw.net/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_xmlhttpRequest
// @require      https://cdnjs.cloudflare.com/ajax/libs/crypto-js/4.0.0/crypto-js.min.js
// @downloadURL  https://cdn.jsdelivr.net/gh/StephenCMZ/oa-hook@main/oa.js
// @updateURL    https://cdn.jsdelivr.net/gh/StephenCMZ/oa-hook@main/oa.js
// ==/UserScript==

(function () {
  'use strict';

  const baseUrl = 'http://oa.gdytw.net';
  const getLinkPageUrl = '/api/Portal/Content/LinkGetPage';
  const logListUrl = '/api/Workflow/FlowList/GetRequireList';
  const logDetailUrl = '/api/Workflow/FlowMan/GetDetail';
  const logContentUrl = '/api/Form/ExternalDataSource/GetDataList';
  const workFlowDetailUrl = '/api/Workflow/FlowMan/GetPrint';
  const workFlowGetPreSelUsersUrl = '/api/Workflow/FlowMan/GetPreSelUsers';
  const userVacationUrl = '/api/Attendance/UserVacation/GetPage';
  const holidayUrl = 'https://cdn.jsdelivr.net/npm/chinese-days/dist/chinese-days.json';
  const hitokotoUrl = 'https://v1.hitokoto.cn/?c=k&encode=text';

  // 表单模板 ID
  const dailyTemplateId = '592233945022595072';
  const weekTemplateId = '592231167478988800';
  const dailyVersionId = '1058193287577341952';
  const dailyVersionId2 = '1124298401509330944';
  const weekVersionId = '1058204181233405952';
  const weekVersionId2 = '1124501387795812352';
  const dailyFormPageIds = [dailyTemplateId, dailyVersionId, dailyVersionId2];
  const weekFormPageIds = [weekTemplateId, weekVersionId, weekVersionId2];

  // 组件 ID
  const nav_setting_btn_id = 'setting_btn';
  const nav_export_btn_id = 'export_btn';
  const nav_fireworks_btn_id = 'fireworks_btn';
  const nav_lot_btn_id = 'lot_btn';
  const nav_statistics_info_id = 'statistics_info';

  // 统计信息
  let statistics = {};

  // AI
  const defaultOpenAIBaseURL = 'https://opencode.ai/zen';
  const defaultOpenAIModel = 'deepseek-v4-flash-free';
  const defaultLogSystemPrompt = '你是一名助理，负责整理工作日志。请将以下日志内容进行归纳总结，提取关键工作内容，使之更清晰、有条理。保持简洁，不要遗漏重要事项，不要添加额外内容。';

  // 设置
  const defaultSettings = {
    openAIBaseURL: defaultOpenAIBaseURL, // OpenAI API 基础 URL
    openAIAPIKey: '', // OpenAI API 密钥
    openAIModel: defaultOpenAIModel, // OpenAI 模型名称
    logSystemPrompt: defaultLogSystemPrompt, // OpenAI 日志整理提示词
    weekDailyLogYear: '', // 下载周志年份，格式为 YYYY
    showDownloadWeekDailyLogBtn: false, // 显示下载周志记录按钮
    weekDailyLogStartDate: '', // 自动填充周志开始时间，格式为 YYYY-MM-DD
    weekDailyLogEndDate: '', // 自动填充周志结束时间，格式为 YYYY-MM-DD
    aiFillWeeklyLog: true, // 自动填充周报记录时，是否使用 AI 整理
    autoFillWeeklyLog: true, // 自动填充周报记录
    autoFillDailyLog: true, // 自动填充日报记录
    autoFillPlan: true, // 自动填充明日/下周工作计划
    autoSelectReviewer: true, // 自动选择日报/周报抄送人和点评人
    showStatisticsInfo: true, // 显示统计信息
    showHitokoto: true, // 显示每日一言
    showFireworks: true, // 显示假日烟花
    showFireworksBtn: false, // 显示放烟花按钮
    fireworksText: '节日快乐', // 放烟花按钮文本
    showLotBtn: false, // 显示每日一签按钮
    debug: false, // 调试模式
  };
  let settings = { ...defaultSettings, ...getConfig('settings') };
  settings.openAIBaseURL = settings.openAIBaseURL?.trim().length ? settings.openAIBaseURL : defaultOpenAIBaseURL;
  settings.openAIModel = settings.openAIModel?.trim().length ? settings.openAIModel : defaultOpenAIModel;
  settings.logSystemPrompt = settings.logSystemPrompt?.trim().length ? settings.logSystemPrompt : defaultLogSystemPrompt;
  settings.fireworksText = settings.fireworksText?.trim().length ? settings.fireworksText : defaultSettings.fireworksText;
  log('【设置】', settings);

  const pageSize = 200;
  let authorization = '';
  let userName = '';

  init();

  /** 初始化 */
  function init() {
    hookShortMenu(); // 新增周日报记录菜单
    autoSelectReviewer(); // 自动选择日报/周报抄送人和点评人
    window.addEventListener('load', function () {
      guardAddElement(addSettingBtn); // 添加导航栏设置按钮
      guardAddElement(addExportBtn); // 添加导航栏导出按钮
      guardAddElement(addFireworksBtn); // 添加放烟花按钮
      guardAddElement(addLotBtn); // 添加每日一签按钮
      guardAddElement(addStatisticsInfo); // 添加导航栏统计信息
      guardAddElement(addAIDailyLogBtn); // 添加 AI 整理日志按钮
      guardAddElement(addAIWeekLogBtn); // 添加 AI 整理周志按钮
      guardFillEditForm(autoFillFormPlan, [...dailyFormPageIds, ...weekFormPageIds]); // 自动填充明日/下周工作计划
      guardFillEditForm(autoFillFormDailyLog, dailyFormPageIds); // 自动填充日报记录
      guardFillEditForm(autoFillFormWeekLog, weekFormPageIds); // 自动填充周报记录
      checkAndShowFireworks(); // 节假日烟花
    });
  }

  /** 守卫添加元素 */
  async function guardAddElement(addFun) {
    if (!addFun) return;

    // 仅进入管理页面后才添加元素
    if (!isManagePage()) return;

    // 已经添加过，直接返回
    const isAdd = await addFun();
    if (isAdd) return;

    // 不存在，等待 100 毫秒，再次检查
    await new Promise((resolve) => setTimeout(resolve, 100));
    guardAddElement(addFun);
  }

  /** 守卫填充表单内容 */
  async function guardFillEditForm(fillFun, templateIds) {
    if (!fillFun || !templateIds || !templateIds.length) return;

    // 仅在日报/周报计划页面填充
    const formPage = templateIds.some((id) => isFormPage(id));
    if (!formPage) return;

    // 填充表单内容
    const isFilled = await fillFun();
    if (isFilled) return;

    // 填充失败，等待 100 毫秒，再次检查
    await new Promise((resolve) => setTimeout(resolve, 100));
    guardFillEditForm(fillFun, templateIds);
  }

  /** =================================== 周日报记录 快捷按钮 ============================================ */

  /** 添加 周日报记录 快捷按钮 */
  function hookShortMenu() {
    hookRequest({
      url: getLinkPageUrl,
      fun: function (res) {
        const data = res.Data.Data;
        const index = data.findIndex((item) => item.Title === '日报');
        if (index >= 0) {
          // 新增周日报记录菜单
          const _record = JSON.parse(JSON.stringify(res.Data.Data[index]));
          _record.Title = '周日报记录';
          _record.Attribute.Href = '/workflow/search';
          res.Data.Data.splice(index, 0, _record);
          res.Data.Total = parseInt(res.Data.Total) + 1 + '';
        }
        return res;
      },
    });
  }

  /** =================================== 自动填充明日/下周工作计划 ============================================ */

  /** 自动填充表单明日/下周工作计划 */
  function autoFillFormPlan() {
    return new Promise(async (resolve) => {
      // 未开启自动填充明日/下周工作计划，直接返回
      if (!settings.autoFillPlan) return resolve(true);

      // 获取计划输入框
      const planTextarea = getPageFormTextarea(['明日工作计划', '下周工作计划']);
      if (!planTextarea) return resolve(false);

      // 填入计划
      fillFormPlan(planTextarea);
      resolve(true);
    });
  }

  /** 填充表单计划内容 */
  async function fillFormPlan(planTextarea) {
    if (!planTextarea) return;

    try {
      // 获取最新日志
      const lastDailyLogRes = await getLastDailyLog();
      const lastDailyLog = (((lastDailyLogRes || {}).Data || {}).Data || [])[0];
      if (!lastDailyLog) {
        log('【表单计划】', '获取最新日报失败');
        return;
      }

      // 获取日志计划内容
      const planContent = await getDailyContent(lastDailyLog.ProcessId);
      if (!planContent || !planContent.plan) {
        log('【表单计划】', '获取最近日报计划内容失败');
        return;
      }

      // 填充计划内容
      planTextarea.value = planContent.plan;
    } catch (error) {
      log('【表单计划】', '填充计划内容失败');
    }
  }

  /** 获取最新日志 */
  function getLastDailyLog() {
    const data = {
      page: 1,
      pageSize: 1,
      sort: 'CreateTime-desc',
      filter: `TaskName~contains~'日计划'`,
      filterMode: 0,
    };
    return request({ url: logListUrl, data });
  }

  /** 获取日志内容 */
  function getDailyContent(processId) {
    return new Promise(async (resolve, reject) => {
      if (!processId) return reject('【日志详情】', '未传入 processId');
      try {
        const dailyContentRes = await request({
          url: workFlowDetailUrl,
          method: 'GET',
          data: {
            processId,
            showForm: true,
            showAttList: false,
            showHisList: false,
            t: new Date().getTime(),
          },
        });

        var doc = document.createElement('div');
        doc.innerHTML = dailyContentRes;
        const user = doc.querySelectorAll('[fsref="db.姓名"]')[0]?.getAttribute('value');
        const dept = doc.querySelectorAll('[fsref="db.所属部门"]')[0]?.getAttribute('value');
        const date = doc.querySelectorAll('[fsref="db.日期"]')[0]?.getAttribute('value');
        const time = doc.querySelectorAll('[fsref="db.时间"]')[0]?.getAttribute('value');
        const content = doc.querySelectorAll('[fsref="db.今天工作总结"]')[0]?.getAttribute('value');
        const plan = doc.querySelectorAll('[fsref="db.明日工作计划"]')[0]?.getAttribute('value');
        const experience = doc.querySelectorAll('[fsref="db.工作心得体会"]')[0]?.getAttribute('value');
        const dailyContent = { user, dept, date, time, content, plan, experience };
        log('【日志详情】', dailyContent);

        resolve(dailyContent);
      } catch (error) {
        log('【日志详情】', error);
        reject(error);
      }
    });
  }

  /** =================================== 自动填充日工作总结 ============================================ */

  /** 添加 AI 整理日志按钮 */
  function addAIDailyLogBtn() {
    return new Promise(async (resolve) => {
      // 检查是否配置了 AI 密钥
      if (!settings.openAIAPIKey || !settings.openAIAPIKey.trim().length) return resolve(true);

      // 检查是否在日志表单页面
      const formPage = dailyFormPageIds.some((id) => isFormPage(id));
      if (!formPage) return resolve(true);

      // 获取表单操作栏
      const formFooterBar = getFormFooterBar();
      if (!formFooterBar) return resolve(false);

      // 添加 AI 生成日志按钮
      const verifyOperators = formFooterBar.querySelector('#verify-operators');
      if (!verifyOperators) return resolve(false);
      const aiDailyLogBtn = document.createElement('button');
      aiDailyLogBtn.classList.add('mr-sm', 'ant-btn', 'ant-btn-primary', 'ng-star-inserted');
      aiDailyLogBtn.textContent = '润色日志';
      aiDailyLogBtn.id = 'ai-daily-log-btn';
      aiDailyLogBtn.addEventListener('click', aiReworkDailyLog);
      verifyOperators.insertBefore(aiDailyLogBtn, verifyOperators.children[0]);
    });
  }

  /** 更新 AI 整理日志按钮文本 */
  function updateAIDailyLogBtnText(text, disabled = false) {
    const aiDailyLogBtn = document.getElementById('ai-daily-log-btn');
    if (!aiDailyLogBtn) return;
    aiDailyLogBtn.textContent = text;
    aiDailyLogBtn.disabled = disabled;
  }

  /** AI 润色日志 */
  async function aiReworkDailyLog() {
    // 获取今天工作总结输入框
    const dailyLogTextarea = getPageFormTextarea(['今天工作总结']);
    if (!dailyLogTextarea) {
      toast('未找到今天工作总结输入框');
      log('【日工作总结】', '未找到今天工作总结输入框');
      return;
    }

    // 获取今日工作总结内容
    const dailyLogContent = dailyLogTextarea.value;
    if (!dailyLogContent || !dailyLogContent.trim().length) {
      toast('请先输入今天工作总结内容');
      log('【日工作总结】', '请先输入今天工作总结内容');
      return;
    }

    try {
      // 更新按钮状态
      updateAIDailyLogBtnText('润色中...', true);

      // 调用 AI 润色日志接口
      const aiLogDetails = await openAIChat(dailyLogContent);
      if (aiLogDetails && aiLogDetails.length) {
        dailyLogTextarea.value = aiLogDetails;
        toast('AI 润色成功');
      } else {
        toast('AI 润色失败');
        log('【日工作总结】', 'AI 润色日志失败');
      }
    } catch (error) {
      toast('AI 润色失败');
      log('【日工作总结】', 'AI 润色日志失败');
    } finally {
      // 更新按钮状态
      updateAIDailyLogBtnText('润色日志');
    }
  }

  /** 自动填充表单日工作总结 */
  function autoFillFormDailyLog() {
    return new Promise(async (resolve) => {
      // 检查是否开启自动填充日志
      if (!settings.autoFillDailyLog) return resolve(true);

      // 获取今天工作总结输入框
      const dailyLogTextarea = getPageFormTextarea(['今天工作总结']);
      if (!dailyLogTextarea) return resolve(false);

      // 填充日工作总结内容
      fillFormDailyLog(dailyLogTextarea);
      resolve(true);
    });
  }

  /** 填充表单日工作总结内容 */
  async function fillFormDailyLog(dailyLogTextarea) {
    if (!dailyLogTextarea) return;

    try {
      // 获取最新日志
      const lastDailyLogRes = await getLastDailyLog();
      const lastDailyLog = (((lastDailyLogRes || {}).Data || {}).Data || [])[0];
      if (!lastDailyLog) {
        log('【日工作总结】', '获取最新日报失败');
        return;
      }

      // 获取最新日志内容
      const dailyContent = await getDailyContent(lastDailyLog.ProcessId);
      if (!dailyContent || !dailyContent.content) {
        log('【日工作总结】', '获取最新日报内容失败');
        return;
      }

      // 填充日工作总结内容
      dailyLogTextarea.value = dailyContent.content;
    } catch (error) {
      log('【日工作总结】', '填充日工作总结内容失败');
    }
  }

  /** =================================== 自动填充本周工作总结 ============================================ */

  /** 添加 AI 整理周志按钮 */
  function addAIWeekLogBtn() {
    return new Promise(async (resolve) => {
      // 检查是否配置了 AI 密钥
      if (!settings.openAIAPIKey || !settings.openAIAPIKey.trim().length) return resolve(true);

      // 检查是否开启 AI 整理周志
      if (!settings.aiFillWeeklyLog) return resolve(true);

      // 检查是否在周志表单页面
      const formPage = weekFormPageIds.some((id) => isFormPage(id));
      if (!formPage) return resolve(true);

      // 获取表单操作栏
      const formFooterBar = getFormFooterBar();
      if (!formFooterBar) return resolve(false);

      // 添加 AI 生成周志按钮
      const verifyOperators = formFooterBar.querySelector('#verify-operators');
      if (!verifyOperators) return resolve(false);
      const aiWeekLogBtn = document.createElement('button');
      aiWeekLogBtn.classList.add('mr-sm', 'ant-btn', 'ant-btn-primary', 'ng-star-inserted');
      aiWeekLogBtn.textContent = '重新生成周志';
      aiWeekLogBtn.id = 'ai-week-log-btn';
      aiWeekLogBtn.addEventListener('click', autoFillFormWeekLog);
      verifyOperators.insertBefore(aiWeekLogBtn, verifyOperators.children[0]);
    });
  }

  /** 更新 AI 整理周志按钮文本 */
  function updateAIWeekLogBtnText(text, disabled = false) {
    const aiWeekLogBtn = document.getElementById('ai-week-log-btn');
    if (!aiWeekLogBtn) return;
    aiWeekLogBtn.textContent = text;
    aiWeekLogBtn.disabled = disabled;
  }

  /** 自动填充表单本周工作总结 */
  function autoFillFormWeekLog() {
    return new Promise(async (resolve) => {
      // 检查是否开启自动填充周志
      if (!settings.autoFillWeeklyLog) return resolve(true);

      // 获取本周工作总结输入框
      const weekLogTextarea = getPageFormTextarea(['本周工作总结']);
      if (!weekLogTextarea) return resolve(false);

      // 填充本周工作总结内容
      fillFormWeekLog(weekLogTextarea);
      resolve(true);
    });
  }

  /** 填充表单本周工作总结内容 */
  async function fillFormWeekLog(weekLogTextarea) {
    if (!weekLogTextarea) return;

    try {
      // 更新按钮状态
      updateAIWeekLogBtnText('AI 生成中...', true);

      // 获取本周所有日志
      const logListRes = await getWeekDailyLogList();
      const logList = ((logListRes || {}).Data || {}).Data || [];
      log('【本周日志列表】', logList);
      if (!logList.length) {
        return toast('获取本周日志列表失败');
      }

      // 获取本周所有日志详情
      const logDetails = await Promise.all(logList.map((log) => getDailyContent(log.ProcessId)));
      log('【本周日志详情】', logDetails);
      if (!logDetails || !logDetails.length) {
        return toast('获取本周日志详情失败');
      }

      // 合并本周所有日志
      let weekLogs = '';
      logDetails.reverse().forEach((log) => {
        if (!log || !log.content) return;
        weekLogs += `${log.content}\n\n`;
      });

      if (!weekLogs) {
        return toast('本周暂无无日志');
      }

      // 不开启 AI 整理，直接填充原始内容
      if (!settings.aiFillWeeklyLog) {
        weekLogTextarea.value = weekLogs;
        return;
      }

      // OpenAI API 整理内容
      const openAIAPIKey = settings.openAIAPIKey;
      if (openAIAPIKey && openAIAPIKey.length) {
        try {
          const aiLogDetails = await openAIChat(weekLogs);
          if (aiLogDetails && aiLogDetails.length) {
            weekLogs = aiLogDetails;
          }
        } catch (error) {
          toast('AI整理周志失败，直接填充原始内容');
          log('【AI整理内容失败】', error);
        }
      } else {
        log('【本周工作总结】', '尚未配置 AI 密钥，直接填充原始内容');
      }

      // 填充本周工作总结
      weekLogTextarea.value = weekLogs;
    } catch (error) {
      log('【本周工作总结】', '填充本周工作总结内容失败');
      toast('填充本周工作总结内容失败');
    } finally {
      // 更新按钮状态
      updateAIWeekLogBtnText('重新生成周志');
    }
  }

  function getWeekDailyLogList() {
    let startDate = settings.weekDailyLogStartDate || '';
    let endDate = settings.weekDailyLogEndDate || '';
    if (!startDate.length || !isDateValid(startDate)) startDate = getMonday();
    if (!endDate.length || !isDateValid(endDate)) endDate = getSunday();
    const data = {
      page: 1,
      pageSize: 20,
      sort: 'CreateTime-desc',
      filter: `(TaskName~contains~'日计划'~and~(CreateTime~gte~datetime'${startDate}T00-00-00'~and~CreateTime~lte~datetime'${endDate}T23-59-59'))`,
      filterMode: 0,
    };
    return request({ url: logListUrl, data });
  }

  // 获取本周一日期 YYYY-MM-DD
  function getMonday() {
    const today = new Date();
    const dayOfWeek = today.getDay();
    const diff = today.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
    today.setDate(diff);
    return today.toISOString().split('T')[0];
  }

  // 获取本周日日期 YYYY-MM-DD
  function getSunday() {
    const today = new Date();
    const dayOfWeek = today.getDay();
    const diff = today.getDate() + (7 - dayOfWeek);
    today.setDate(diff);
    return today.toISOString().split('T')[0];
  }

  // 判断是否为日期格式 YYYY-MM-DD
  function isDateValid(dateString) {
    const regex = /^\d{4}-\d{2}-\d{2}$/;
    return regex.test(dateString);
  }

  // 判断是否为年份格式 YYYY
  function isYearValid(yearString) {
    const regex = /^\d{4}$/;
    return regex.test(yearString);
  }

  // 格式化日期为 YYYY-MM-DD
  function formatDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  /** =================================== 自动选择日报/周报抄送人和点评人 ============================================ */

  /** 自动选择日报/周报抄送人和点评人 */
  function autoSelectReviewer() {
    if (!settings.autoSelectReviewer) return;
    var _ob = function (s) {
      return JSON.parse(
        decodeURIComponent(
          Array.prototype.map
            .call(atob(s), function (c) {
              return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
            })
            .join(''),
        ),
      );
    };
    hookRequest({
      url: workFlowGetPreSelUsersUrl,
      fun: function (res) {
        const users = res.Data.ForPreSelUsers;
        if (users && users.length === 2) {
          const toNodeName0 = (users[0] || {}).ToNodeName; // 抄送
          const toNodeName1 = (users[1] || {}).ToNodeName; // 点评人

          if (toNodeName0 === '抄送' && toNodeName1 === '点评人') {
            users[0] = {
              ...users[0],
              SelUsers: [
                _ob(
                  'eyJJZCI6IjU4OTM3NjQ3ODIwMjkyOTE1MiIsIk5hbWUiOiLlhbPkv6HkuJwiLCJDb2RlIjoiNzAwNyIsIkRlcElkIjoiMzI1NDc1Mzk2MDEzMjc3MTg0IiwiRGVwTmFtZSI6Iue7vOWQiOS6i+WKoemDqCIsIkRlcEZ1bGxOYW1lIjoi57u85ZCI5LqL5Yqh6YOoIiwiRGVwUGF0aCI6IjAzMjU0NzUzOTYwMTMyNzcxODQiLCJTb3J0Q29kZSI6Ijk5OTkiLCJBbGxvd0xvZ2luIjp0cnVlLCJBY3RpdmVkIjp0cnVlLCJQaW55aW4iOiJHdWFuWGluRG9uZyIsIkluaXRpYWxQaW55aW4iOiJHWEQiLCJ0eXBlIjowfQ',
                ),
              ],
            };
            users[1] = {
              ...users[1],
              SelUsers: [
                _ob(
                  'eyJOYW1lIjoi5a2U5paH5aiBIiwiQWNjb3VudCI6IjIwMTEiLCJDb2RlIjoiMjAxMSIsIkdlbmRlciI6dHJ1ZSwiQ29udGFjdFZpc2liaWxpdHkiOnRydWUsIk1vYmlsZSI6IjE4MDI4MTk2NTU5IiwiQ29ybmV0IjpudWxsLCJUZWxlcGhvbmUiOm51bGwsIkVtYWlsIjoiNDIyOTYzODQ1QHFxLmNvbSIsIldlQ2hhdCI6bnVsbCwiU3VwZXJpb3JJZCI6IjAiLCJBY3RpdmVkIjp0cnVlLCJBbGxvd0xvZ2luIjp0cnVlLCJJbml0aWFsUGlueWluIjoiS1dXIiwiUGlueWluIjoiS29uZ1dlbldlaSIsIklzT25saW5lIjpmYWxzZSwiSXNFeHRlcm5hbCI6ZmFsc2UsIkV4dGVuc2lvbk51bWJlciI6bnVsbCwiQWxsb3dNb2JpbGUiOnRydWUsIkluaXRpYWxXdWJpIjoiQllEIiwiQWNjb3VudFZhbGlkaXR5IjpudWxsLCJQaW4iOm51bGwsIkNyZWF0b3IiOm51bGwsIkNyZWF0b3JJZCI6IjAiLCJDcmVhdGVUaW1lIjpudWxsLCJVcGRhdGVVc2VyTmFtZSI6Ium+meWKoOmOjyIsIlVwZGF0ZVVzZXJJZCI6IjU4Mzk2Mjc2NTU1OTk1OTU1MiIsIlVwZGF0ZVRpbWUiOiIyMDI1LTEyLTMwVDEwOjQ2OjA4LjU1MTg1MTMiLCJJbmFjdGl2ZVRpbWUiOm51bGwsIkRlcGFydG1lbnQiOnsiTmFtZSI6Iui9r+S7tuS/oeaBr+mDqCIsIlBhdGgiOiIwMzI0Nzk2OTQwNDM0ODA0NzM2IiwiRnVsbE5hbWUiOiLova/ku7bkv6Hmga/pg6giLCJBY3RpdmVkIjp0cnVlLCJTb3J0Q29kZSI6IjAwMDUiLCJDYXRlZ29yaWVzIjoiIiwiV2VDaGF0V29ya0RlcElkIjo1OSwiV2VDaGF0V29ya0NvcnBJZCI6Ind3NDFmYTNmZGQzMDMxOGJlYiIsIk1hbm5pbmciOjAsIk5vTGltaXQiOmZhbHNlLCJDb250YWN0VmlzaWJpbGl0eSI6dHJ1ZSwiVXNlckNvdW50IjowLCJJZCI6IjMyNDc5Njk0MDQzNDgwNDczNiJ9LCJQb3NpdGlvbnMiOlt7IlVzZXJJZCI6IjU4Mzk2Mjc1NjQyNDc2NTQ0MCIsIlBvc2l0aW9uSWQiOiI2NjgwMjc4NTgyNDA3ODIzMzYiLCJNYWpvciI6dHJ1ZSwiU2VxdWVuY2UiOjEsIlNvcnRDb2RlIjoiOTk5OTk5IiwiUG9zaXRpb24iOnsiSm9iSWQiOiI2NDU1NzcxMjI5ODI3MDMxMDQiLCJPcmdhbml6ZUlkIjoiMzI0Nzk2OTQwNDM0ODA0NzM2IiwiTmFtZSI6Iui9r+S7tuS/oeaBr+mDqC/mioDmnK/mgLvnm5EiLCJKb2IiOnsiTmFtZSI6IuaKgOacr+aAu+ebkSIsIlN5c0J1aWxkSW4iOmZhbHNlLCJTb3J0Q29kZSI6OTk5OTk5LCJJZCI6IjY0NTU3NzEyMjk4MjcwMzEwNCJ9LCJJZCI6IjY2ODAyNzg1ODI0MDc4MjMzNiJ9LCJJZCI6IjU4Mzk2Mjc1NzI0Njg0OTAyNCJ9XSwiUm9sZXMiOltdLCJPcmdhbml6ZVJvbGVzIjpbXSwiUGFzc3dvcmQiOm51bGwsIlN1cGVyaW9yIjpudWxsLCJSZWxhdGlvbk9yZ2FuaXplcyI6bnVsbCwiVXNlckNhcmRFeHRyYSI6bnVsbCwiSWQiOiI1ODM5NjI3NTY0MjQ3NjU0NDAiLCJKb2IiOnsiTmFtZSI6IuaKgOacr+aAu+ebkSIsIlN5c0J1aWxkSW4iOmZhbHNlLCJTb3J0Q29kZSI6OTk5OTk5LCJJZCI6IjY0NTU3NzEyMjk4MjcwMzEwNCJ9LCJVc2VySWQiOiI1ODM5NjI3NTY0MjQ3NjU0NDAiLCJGb3JtZXJOYW1lIjpudWxsLCJQb3N0SWQiOiI1IiwiUG9zdCI6IuaKgOacr+WylyIsIlBvc3RJZFBhdGgiOlsiNSJdLCJUaXRsZUlkIjpudWxsLCJUaXRsZSI6bnVsbCwiVGl0bGVDb25mZXJpbmdEYXRlIjpudWxsLCJUaXRsZUdyYWRlSWQiOm51bGwsIlRpdGxlR3JhZGUiOm51bGwsIlN0YXR1cyI6IuWcqOiBjCIsIkVtcGxveWVlQ2F0ZWdvcnkiOm51bGwsIkJpcnRoZGF5IjoiMTk4Ni0wMi0xOFQwMDowMDowMCIsIklkZW50aXR5Q2FyZE51bWJlciI6IjQ0MDY4MjE5ODYwMjE4MTAzNiIsIlBhc3Nwb3J0TnVtYmVyIjpudWxsLCJOYXRpb25hbGl0eSI6IuS4reWbvSIsIkV0aG5pY0dyb3VwIjoi5rGJ5pePIiwiTmF0aXZlUGxhY2UiOiLkvZvlsbEiLCJDdXJyZW50UmVzaWRlbmNlIjoi5bm/5bee5biC55m95LqR5Yy65Lqs5rqq6Lev5LqR5pmv6Iqx5Zut5paw5LqR5qGC6IuRMTbmoIsiLCJSZWdpc3RlZFJlc2lkZW5jZSI6bnVsbCwiTWFyaXRhbFN0YXR1cyI6IuacquWpmiIsIlBvbGl0aWNhbFN0YXR1cyI6Iue+pOS8lyIsIkhpZ2hlc3REZWdyZWUiOiLnoJTnqbbnlJ8iLCJTdGFydGluZ0RhdGVPZkZpcnN0Sm9iIjoiMjAwOS0wMy0wMVQwMDowMDowMCIsIkVudGVyRGF0ZSI6IjIwMTctMDItMDZUMDA6MDA6MDAiLCJSZXNpZ25hdGlvbkRhdGUiOm51bGwsIlJldGlyZW1lbnREYXRlIjpudWxsLCJJbnN1cmFuY2UiOm51bGwsIkluc3VyZWREYXRlIjpudWxsLCJBdHRlbmRhbmNlTWFjaGluZUlkIjpudWxsLCJQb3N0Q2F0ZWdvcnkiOm51bGwsIkJpcnRoQWRkcmVzcyI6bnVsbCwiSm9pblBhcnR5RGF0ZSI6bnVsbCwiQXJjaGl2ZXNOdW1iZXIiOm51bGwsIlN0YWZmaW5nQ2F0ZWdvcnkiOm51bGwsIkFjY291bnRVbml0SWQiOm51bGwsIkFjY291bnRVbml0TmFtZSI6bnVsbCwiUGVyc29ubmVsTW9kZWwiOm51bGwsIk9yZ0pvYk1pZCI6bnVsbCwiQXZhdGFyRmlsZUlkIjpudWxsLCJJc0xlYWRlciI6bnVsbCwiUGVyc29ubmVsQ29udHJhY3QiOm51bGwsIkRlcGFydG1lbnRSZWNvcmQiOm51bGwsIkV4dHJhRmlsZSI6bnVsbCwiRW1wbG95ZWVJZCI6IjU4Mzk2Mjc1NjQ4NzY4MDAwMCIsInR5cGUiOjB9',
                ),
              ],
            };
          }
        }
        return res;
      },
    });
  }

  /** =================================== 导出全年周志 ============================================ */

  /** 导航栏添加 下载周志 按钮 */
  function addExportBtn() {
    return new Promise((resolve) => {
      // 检查是否显示下载周志记录按钮
      if (!settings.showDownloadWeekDailyLogBtn) return resolve(true);

      var { navBar, exists } = hasNavBarItem(nav_export_btn_id);
      if (!navBar || exists) return resolve(exists);

      var liItem = document.createElement('li');
      liItem.id = nav_export_btn_id;
      liItem.className = 'ng-star-inserted';
      liItem.style = 'display: inline-block; vertical-align: middle;';

      var exportButton = document.createElement('button');
      exportButton.textContent = '下载周志';
      exportButton.style.backgroundColor = 'transparent';
      exportButton.style.color = 'white';
      exportButton.style.border = 'none';
      exportButton.style.textAlign = 'center';
      exportButton.style.textDecoration = 'none';
      exportButton.style.display = 'inline-block';
      exportButton.style.fontSize = '14px';
      exportButton.style.cursor = 'pointer';

      exportButton.onmouseover = function () {
        exportButton.style.backgroundColor = 'hsla(0, 0%, 100%, .2)';
      };
      exportButton.onmouseout = function () {
        exportButton.style.backgroundColor = 'transparent';
      };

      exportButton.onclick = exportWeeklyLogs;

      liItem.appendChild(exportButton);
      navBar.insertBefore(liItem, navBar.firstChild);

      resolve(true);
    });
  }

  /** 导出全年周志 markdown 文件 */
  async function exportWeeklyLogs() {
    toast('开始导出周志，请稍后...');
    try {
      // 获取周志列表
      const logListRes = await getWeeklyLogList();
      const logList = ((logListRes || {}).Data || {}).Data || [];
      log('【周志列表】', logList);
      if (!logList.length) {
        return toast('获取周志列表失败');
      }

      // 设置用户名
      userName = logList[0].CreateUserName || '';
      if (!userName.length) {
        log('【用户名】', userName);
        return toast('获取用户名失败');
      }

      // 获取周志详情
      const logDetailRes = await Promise.all(logList.map((log) => getWeeklyLogDetail(log.ProcessId)));
      const logDetails = (logDetailRes || []).map((res) => {
        const data = (res || {}).Data || {};
        return { FormId: data.FormId, TaskName: data.TaskName };
      });
      log('【周志详情】', logDetails);

      // 获取周志内容
      const logContentRes = await getWeeklyLogContent();
      const logContents = ((logContentRes || {}).Data || {}).Data || [];
      log('【周志内容】', logContents);

      // 合并周志列表和周志内容
      const mergedLogs = logDetails.map((log) => {
        const content = logContents.find((item) => item.id === log.FormId);
        return { ...log, ...content };
      });
      log('【合并周志】', mergedLogs);

      // 整理 markdown 内容
      const mergedLogsStr = mergedLogs
        .map((log) => {
          const title = log.TaskName.split('【周计划】');
          return `# ${title[title.length - 1]}\n\n${log.jtgzzj}`;
        })
        .join('\n');
      log('【导出内容】', mergedLogsStr);

      downloadMarkdown(mergedLogsStr);
      log('【导出结果】', '导出成功');
      toast('导出周志成功!!!');
    } catch (error) {
      toast('导出周志失败!!!');
    }
  }

  /** 获取全年周计划列表 */
  function getWeeklyLogList() {
    let year = settings.weekDailyLogYear || '';
    if (!year.length || !isYearValid(year)) {
      year = new Date().getFullYear().toString();
    }
    const data = {
      page: 1,
      pageSize: pageSize,
      sort: 'CreateTime-desc',
      filter: `(TaskName~contains~'周计划'~and~(CreateTime~gte~datetime'${year}-01-01T00-00-00'~and~CreateTime~lte~datetime'${year}-12-31T23-59-59'))`,
      filterMode: 0,
    };
    return request({ url: logListUrl, data });
  }

  /** 获取周志详情 */
  function getWeeklyLogDetail(processId) {
    const data = {
      isClose: true,
      isExternal: false,
      toIndex: false,
      params: { processId },
      processId,
      processTodoId: null,
      versionId: null,
      templateId: null,
      isMonitor: false,
      isDraft: false,
      notificationId: null,
      relatedId: null,
      moduleCode: null,
      objectId: null,
      isContinue: null,
      isAnonymous: null,
      callback: null,
    };
    return request({ url: logDetailUrl, data });
  }

  /** 获取周志内容 */
  function getWeeklyLogContent() {
    const requestPara = {
      Id: '32241332-3396-4cbe-97c8-0af9cd81cfef',
      DataSourceType: 0,
      DataSourceId: '-1',
      ObjectName: 'fs_zhouj_ji_hua',
      SchemaName: 'dbo',
      ObjectType: 0,
      Sort: 'id desc',
      InputPara: `'${userName}'=shen_qing_ren and FsIsDelete=0`,
      Mapping: '["id","jtgzzj"]',
    };
    const data = {
      page: 1,
      pageSize: pageSize,
      requestPara: btoa(encodeURIComponent(JSON.stringify(requestPara))),
    };
    return request({ url: logContentUrl, data });
  }

  /** =================================== 配置信息 ============================================ */

  /** 添加设置按钮 */
  function addSettingBtn() {
    return new Promise((resolve) => {
      var { navBar, exists } = hasNavBarItem(nav_setting_btn_id);
      if (!navBar || exists) return resolve(exists);

      var liItem = document.createElement('li');
      liItem.id = nav_setting_btn_id;
      liItem.className = 'ng-star-inserted';
      liItem.style = 'display: inline-block; vertical-align: middle;';

      var settingButton = document.createElement('button');
      settingButton.textContent = '设置';
      settingButton.style.backgroundColor = 'transparent';
      settingButton.style.color = 'white';
      settingButton.style.border = 'none';
      settingButton.style.textAlign = 'center';
      settingButton.style.textDecoration = 'none';
      settingButton.style.display = 'inline-block';
      settingButton.style.fontSize = '14px';
      settingButton.style.cursor = 'pointer';

      settingButton.onmouseover = function () {
        settingButton.style.backgroundColor = 'hsla(0, 0%, 100%, .2)';
      };
      settingButton.onmouseout = function () {
        settingButton.style.backgroundColor = 'transparent';
      };

      settingButton.onclick = showSettings;

      liItem.appendChild(settingButton);
      navBar.insertBefore(liItem, navBar.firstChild);

      resolve(true);
    });
  }

  /** 创建分组标题 */
  function createGroupTitle(text) {
    const groupTitle = document.createElement('div');
    groupTitle.textContent = text;
    groupTitle.style.cssText = 'font-size:13px;font-weight:600;color:#1890ff;padding:8px 0 4px;margin-top:6px;margin-bottom:4px;border-bottom:1px solid #f0f0f0;';
    return groupTitle;
  }

  /** 创建设置行 */
  function createSettingRow(label, element) {
    const row = document.createElement('div');
    row.style.cssText = 'display:flex;align-items:center;gap:8px;padding:4px 0;';
    const labelEl = document.createElement('span');
    labelEl.textContent = label;
    labelEl.style.cssText = 'font-size:13px;color:#333;white-space:nowrap;min-width:85px;text-align:right;';
    row.appendChild(labelEl);
    element.style.flex = '1';
    row.appendChild(element);
    return row;
  }

  /** 设置弹窗 */
  function showSettings() {
    // 创建弹窗容器
    const overlay = document.createElement('div');
    overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.45);z-index:9998;display:flex;align-items:center;justify-content:center;';

    const dialog = document.createElement('div');
    dialog.style.cssText = 'background:#fff;border-radius:8px;box-shadow:0 6px 30px rgba(0,0,0,.15);padding:24px;width:580px;max-height:80vh;overflow-y:auto;';

    // 标题
    const title = document.createElement('div');
    title.style.cssText = 'font-size:18px;font-weight:600;color:#222;margin-bottom:16px;padding-bottom:12px;border-bottom:2px solid #e8e8e8;';
    title.textContent = '设置';
    dialog.appendChild(title);

    // 设置项分组
    const groups = [
      {
        title: 'AI 配置',
        items: [
          { key: 'openAIBaseURL', type: 'text', label: '请求地址', placeholder: '请输入 AI 请求地址' },
          { key: 'openAIAPIKey', type: 'text', label: 'API 密钥', placeholder: '请输入 AI 密钥' },
          { key: 'openAIModel', type: 'text', label: '模型名称', placeholder: '请输入 AI 模型名称' },
          { key: 'logSystemPrompt', type: 'text', label: '提示词', placeholder: '请输入 AI 日志整理提示词' },
        ],
      },
      {
        title: '自动填充',
        items: [
          { key: 'autoFillDailyLog', type: 'checkbox', label: '自动填充日志' },
          { key: 'autoFillWeeklyLog', type: 'checkbox', label: '自动填充周志' },
          { key: 'autoFillPlan', type: 'checkbox', label: '自动填充计划' },
          { key: 'autoSelectReviewer', type: 'checkbox', label: '自动选点评人' },
          { key: 'aiFillWeeklyLog', type: 'checkbox', label: 'AI 整理周志' },
        ],
      },
      {
        title: '填充周志时间范围',
        items: [
          { key: 'weekDailyLogStartDate', type: 'text', label: '开始时间', placeholder: '格式 YYYY-MM-DD，不填默认本周一' },
          { key: 'weekDailyLogEndDate', type: 'text', label: '结束时间', placeholder: '格式 YYYY-MM-DD，不填默认本周日' },
        ],
      },
      {
        title: '批量周志下载',
        items: [
          { key: 'showDownloadWeekDailyLogBtn', type: 'checkbox', label: '显示下载周志' },
          { key: 'weekDailyLogYear', type: 'text', label: '下载周志年份', placeholder: '格式 YYYY，不填默认当前年份' },
        ],
      },
      {
        title: '其他',
        items: [
          { key: 'showStatisticsInfo', type: 'checkbox', label: '显示统计信息' },
          { key: 'showHitokoto', type: 'checkbox', label: '显示每日一言' },
          { key: 'showFireworks', type: 'checkbox', label: '显示假日烟花' },
          { key: 'showFireworksBtn', type: 'checkbox', label: '显示放烟花' },
          { key: 'showLotBtn', type: 'checkbox', label: '显示每日一签' },
          { key: 'debug', type: 'checkbox', label: '调试模式' },
        ],
      },
    ];

    // 收集所有 formItem 引用，方便后续取值
    const allItems = groups.flatMap((g) => g.items);

    // 渲染每个分组
    groups.forEach((group) => {
      dialog.appendChild(createGroupTitle(group.title));

      group.items.forEach((item) => {
        if (item.type === 'text') {
          const { inputElement, input } = createInputElement({ placeholder: item.placeholder, value: settings[item.key] });
          const row = createSettingRow(item.label, inputElement);
          dialog.appendChild(row);
          item.formItem = input;
        } else if (item.type === 'checkbox') {
          const checkboxWrap = document.createElement('div');
          checkboxWrap.style.cssText = 'display:flex;align-items:center;padding-left:9px;';
          const checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.checked = settings[item.key] || false;
          checkboxWrap.appendChild(checkbox);
          const row = createSettingRow(item.label, checkboxWrap);
          dialog.appendChild(row);
          item.formItem = checkbox;
        }
      });
    });

    // 按钮容器
    const btnContainer = document.createElement('div');
    btnContainer.style.cssText = 'text-align:right;margin-top:20px;padding-top:12px;border-top:1px solid #f0f0f0;';

    const cancelBtn = createButtonElement({
      title: '取消',
      onClick: () => document.body.removeChild(overlay),
    });

    const confirmBtn = createButtonElement({
      title: '确认',
      type: 'primary',
      onClick: () => {
        // 校验AI请求地址
        const openAIBaseURL = allItems.find((item) => item.key === 'openAIBaseURL')?.formItem?.value || '';
        if (openAIBaseURL.length && !openAIBaseURL.startsWith('http')) {
          toast('AI请求地址格式异常');
          return;
        }
        // 校验周志开始时间
        const weekDailyLogStartDate = allItems.find((item) => item.key === 'weekDailyLogStartDate')?.formItem?.value || '';
        if (weekDailyLogStartDate.length && !isDateValid(weekDailyLogStartDate)) {
          toast('周志开始时间格式异常');
          return;
        }
        // 校验周志结束时间
        const weekDailyLogEndDate = allItems.find((item) => item.key === 'weekDailyLogEndDate')?.formItem?.value || '';
        if (weekDailyLogEndDate.length && !isDateValid(weekDailyLogEndDate)) {
          toast('周志结束时间格式异常');
          return;
        }
        // 校验下载周志年份
        const weekDailyLogYear = allItems.find((item) => item.key === 'weekDailyLogYear')?.formItem?.value || '';
        if (weekDailyLogYear.length && !isYearValid(weekDailyLogYear)) {
          toast('周志年份格式异常');
          return;
        }

        // 获取所有设置项的值
        const _settings = { ...settings };
        allItems.forEach((item) => {
          if (item.type === 'text') {
            _settings[item.key] = item.formItem.value;
          } else if (item.type === 'checkbox') {
            _settings[item.key] = item.formItem.checked;
          }
        });

        // 保存设置
        settings = _settings;
        setConfig('settings', _settings);
        log('【保存设置】', _settings);

        // 关闭弹窗
        document.body.removeChild(overlay);
        toast('保存成功, 请刷新页面生效');
      },
    });

    btnContainer.appendChild(cancelBtn);
    btnContainer.appendChild(confirmBtn);
    dialog.appendChild(btnContainer);

    overlay.appendChild(dialog);
    document.body.appendChild(overlay);
  }

  /** 创建输入框 */
  function createInputElement({ labelText, value, placeholder }) {
    const inputElement = document.createElement('div');
    inputElement.style.display = 'flex';
    inputElement.style.alignItems = 'center';

    if (labelText && labelText.trim() !== '') {
      const label = document.createElement('label');
      label.textContent = labelText;
      label.style.fontSize = '14px';
      label.style.marginRight = '8px';
      inputElement.appendChild(label);
    }

    const input = document.createElement('input');
    input.type = 'text';
    input.placeholder = placeholder;
    input.value = value || '';
    input.style.flex = 1;
    input.style.padding = '8px';
    input.style.border = '1px solid #d9d9d9';
    input.style.borderRadius = '4px';
    inputElement.appendChild(input);

    return { inputElement, input };
  }

  /** 创建按钮 */
  function createButtonElement({ title, type = 'normal', onClick }) {
    const button = document.createElement('button');
    button.textContent = title;
    button.style.margin = '4px';
    button.style.padding = '4px 15px';
    button.style.backgroundColor = type === 'primary' ? '#1890ff' : type === 'danger' ? '#ff4d4f' : '#f0f0f0';
    button.style.color = type === 'primary' ? 'white' : type === 'danger' ? 'white' : 'black';
    button.style.border = 'none';
    button.style.borderRadius = '4px';
    button.style.cursor = 'pointer';
    button.onclick = onClick;

    return button;
  }

  /** =================================== 统计信息 ============================================ */

  /** 更新统计信息 */
  function updateStatisticsInfo() {
    return new Promise(async (resolve) => {
      if (!settings.showStatisticsInfo) return;

      // 今日日期
      const today = new Date();
      const year = today.getFullYear();
      const month = String(today.getMonth() + 1).padStart(2, '0');
      const day = String(today.getDate()).padStart(2, '0');
      const weekDay = ['日', '一', '二', '三', '四', '五', '六'][today.getDay()];
      const todayDate = `${year}年${month}月${day}日 周${weekDay}`;
      statistics.todayDate = todayDate;

      // 距离周末
      if (weekDay !== '六' && weekDay !== '日') {
        const saturday = new Date(`${year}-${month}-${day}`);
        saturday.setDate(saturday.getDate() + (6 - saturday.getDay()));
        const diffDaysToWeekend = calculateDateDiff(new Date(`${year}-${month}-${day}`), saturday);
        statistics.diffDaysToWeekend = diffDaysToWeekend;
      } else {
        statistics.diffDaysToWeekend = 0;
      }

      // 距离发工资天数，每月5号发工资
      const payDay = '05';
      if (day !== payDay && '0' + day !== payDay) {
        const nextPayDate = new Date(`${year}-${month}-${payDay}`);
        if (nextPayDate < today) {
          nextPayDate.setMonth(nextPayDate.getMonth() + 1);
        }
        const diffDaysToPay = calculateDateDiff(new Date(`${year}-${month}-${day}`), nextPayDate);
        statistics.diffDaysToPay = diffDaysToPay;
      } else {
        statistics.diffDaysToPay = 0;
      }

      // 获取请假信息
      try {
        const userVacationRes = await request({
          url: userVacationUrl,
          data: {
            page: 1,
            pageSize: 10,
            isUsableDays: false,
            startDate: `${year}-01-01 00:00:00`,
            endDate: `${year}-12-31 23:59:59`,
          },
        });
        const userVacationUsableRes = await request({
          url: userVacationUrl,
          data: {
            page: 1,
            pageSize: 10,
            isUsableDays: true,
            startDate: `${year}-01-01 00:00:00`,
            endDate: `${year}-12-31 23:59:59`,
          },
        });
        const userVacations = ((userVacationRes || {}).Data || {}).Data || [];
        const userVacationsUsable = ((userVacationUsableRes || {}).Data || {}).Data || [];
        if (userVacations.length) {
          statistics.vacations = formVacations(userVacations[0] || {}, userVacationsUsable[0] || {});
        } else {
          statistics.vacations = [];
        }
      } catch (error) {}

      // 获取法定节假日
      try {
        const holidayRes = await requestGM({ url: holidayUrl, method: 'GET' });
        let holidays = JSON.parse(holidayRes || '{}').holidays || {};

        // 处理节假日数据
        holidays = Object.keys(holidays)
          .filter((key) => key.startsWith(`${year}-`)) // 过滤出当前年份的节假日
          .filter((key) => new Date(key) >= new Date(`${year}-${month}-${day}`)) // 过滤出过期的节假日
          .map((key) => ({ date: key, name: (holidays[key].split(',') || [])[1] || '' })) // 映射为 { date: 日期, name: 节假日名称 } 格式
          .filter((item, index, arr) => arr.findIndex((i) => i.name === item.name) === index) // 过滤重复节假日名称
          .map((item) => ({ ...item, diffDays: calculateDateDiff(new Date(item.date)) })); // 计算日期相差天数

        statistics.holidays = holidays || [];
      } catch (error) {}

      resolve();
    });
  }

  /** 更新每日一言 */
  function updateHitokoto() {
    return new Promise(async (resolve) => {
      // 判断是否显示每日一言
      if (!settings.showHitokoto) {
        statistics.hitokoto = '';
        return resolve('');
      }

      // 判断今天是否已经获取过每日一言
      const { hitokoto, timestamp } = getConfig('hitokoto') || {};
      const today = new Date();
      const date = timestamp ? new Date(timestamp) : null;
      const isToday = date && date.getDate() === today.getDate() && date.getMonth() === today.getMonth() && date.getFullYear() === today.getFullYear();
      if (isToday && hitokoto && hitokoto.length) {
        statistics.hitokoto = hitokoto;
        return resolve(hitokoto);
      }

      // 获取每日一言
      try {
        const hitokotoRes = await requestGM({ url: hitokotoUrl, method: 'GET' });
        statistics.hitokoto = hitokotoRes || '';
        setConfig('hitokoto', { hitokoto: hitokotoRes || '', timestamp: today.getTime() });
        resolve(hitokotoRes || '');
      } catch (error) {
        resolve('');
      }
    });
  }

  /** 格式化请假信息 */
  function formVacations(userVacation = {}, userVacationUsable = {}) {
    if (!userVacation || !Object.keys(userVacation).length) {
      return {};
    }
    const vacations = [];

    const annual = formVacationByKey(userVacation, '1');
    if (userVacationUsable && Object.keys(userVacationUsable).length && userVacationUsable['1'] && userVacationUsable['1'].length && userVacationUsable['1'] !== '-') {
      // 年假结余会失效，需从可休天数中获取
      const annualUsableTotal = eval((userVacationUsable['1'] || '').trim());
      vacations.push({ key: 'annual', name: '剩余年假', value: annualUsableTotal });
    } else {
      vacations.push({ key: 'annual', name: '剩余年假', value: annual.total - annual.used });
    }
    vacations.push({ key: 'annual-used', name: '已休年假', value: annual.used });
    vacations.push({ key: 'personal-used', name: '已请事假', value: formVacationByKey(userVacation, '4').used });
    vacations.push({ key: 'sick-used', name: '已请病假', value: formVacationByKey(userVacation, '3').used });
    vacations.push({ key: 'marriage-used', name: '已请婚假', value: formVacationByKey(userVacation, '6').used });
    vacations.push({ key: 'maternity-used', name: '已请产假', value: formVacationByKey(userVacation, '8').used });
    vacations.push({ key: 'paternity-used', name: '已请陪产假', value: formVacationByKey(userVacation, '7').used });
    vacations.push({ key: 'funeral-used', name: '已请丧假', value: formVacationByKey(userVacation, '9').used });
    vacations.push({ key: 'breastfeeding-used', name: '已请哺乳假', value: formVacationByKey(userVacation, '2').used });
    vacations.push({ key: 'injury-used', name: '已请工伤假', value: formVacationByKey(userVacation, '14').used });

    return vacations;
  }

  /** 格式化请假信息，例如：1/2 表示已休1天，共2天 */
  function formVacationByKey(userVacation = {}, key = '') {
    if (!userVacation || !Object.keys(userVacation).length || !key || !key.length) {
      return { total: 0, used: 0, key: key };
    }
    const vacation = (userVacation[key] || '').trim();
    const [used, total] = vacation.split('/') || [];
    return { total: eval(!total || total === '-' ? 0 : total), used: eval(!used || used === '-' ? 0 : used), key: key };
  }

  /** 计算日期相差天数 */
  function calculateDateDiff(date1, date2 = new Date()) {
    const diffTime = Math.abs(date2 - date1);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays;
  }

  /** 添加统计信息 */
  async function addStatisticsInfo() {
    return new Promise(async (resolve) => {
      // 不显示统计信息时，直接返回
      if (!settings.showStatisticsInfo) return resolve(true);

      var { navBar, exists } = hasNavBarItem(nav_statistics_info_id);
      if (!navBar || exists) return resolve(exists);

      await updateStatisticsInfo();
      await updateHitokoto();

      var liItem = document.createElement('li');
      liItem.id = nav_statistics_info_id;
      liItem.className = 'ng-star-inserted';
      liItem.style = 'display: inline-block; vertical-align: middle;';

      let statisticsInfo = '';
      statisticsInfo += `距离发工资：${statistics.diffDaysToPay} 天`;
      statisticsInfo += `\n距离周末：${statistics.diffDaysToWeekend} 天`;

      if (statistics.holidays && statistics.holidays.length) {
        statisticsInfo += `\n距离${statistics.holidays[0].name}：${statistics.holidays[0].diffDays} 天`;
      }

      const annual = (statistics.vacations || []).find((item) => item.key === 'annual').value || 0;
      statisticsInfo += `\n剩余年假：${annual} 天`;

      var statisticsButton = document.createElement('div');
      statisticsButton.textContent = statisticsInfo;
      statisticsButton.style.backgroundColor = 'transparent';
      statisticsButton.style.color = 'white';
      statisticsButton.style.border = 'none';
      statisticsButton.style.textAlign = 'left';
      statisticsButton.style.textDecoration = 'none';
      statisticsButton.style.fontSize = '10px';
      statisticsButton.style.cursor = 'pointer';
      statisticsButton.style.whiteSpace = 'pre-wrap';
      statisticsButton.style.lineHeight = 'normal';
      statisticsButton.style.verticalAlign = 'center';
      statisticsButton.style.padding = '0 8px';
      statisticsButton.style.position = 'relative';

      // 添加 hover
      statisticsButton.style.transition = 'background-color 0.3s ease';
      statisticsButton.addEventListener('mouseenter', () => {
        statisticsButton.style.backgroundColor = 'rgba(255, 255, 255, 0.1)';
        showStatisticsDetailInfo(statisticsButton);
      });
      statisticsButton.addEventListener('mouseleave', () => {
        statisticsButton.style.backgroundColor = 'transparent';
        hideStatisticsDetailInfo(statisticsButton);
      });

      // 添加到导航栏
      liItem.appendChild(statisticsButton);
      navBar.appendChild(liItem);

      resolve(true);
    });
  }

  /** 显示统计信息详情 */
  function showStatisticsDetailInfo(statisticsButton) {
    let detailInfo = `${statistics.todayDate}`;

    detailInfo += `\n\n📅`;
    detailInfo += `\n距离发工资：${statistics.diffDaysToPay} 天`;
    detailInfo += `\n距离周末：${statistics.diffDaysToWeekend} 天`;

    // 法定节假日
    if (statistics.holidays && statistics.holidays.length) {
      statistics.holidays.forEach((item) => {
        detailInfo += `\n距离${item.name}：${item.diffDays} 天`;
      });
    }

    // 休假
    if (statistics.vacations && statistics.vacations.length) {
      detailInfo += `\n\n♨️`;
      statistics.vacations.forEach((item) => {
        detailInfo += `\n${item.name}：${item.value} 天`;
      });
    }

    // 每日一言
    if (statistics.hitokoto && statistics.hitokoto.length) {
      const _hitokoto = statistics.hitokoto.replace(/.{1,10}/g, '$&\n');
      detailInfo += `\n\n🤔`;
      detailInfo += `\n${_hitokoto}`;
    }

    // 显示统计信息详情
    const statisticsDetailInfo = document.createElement('div');
    statisticsDetailInfo.id = 'statistics-detail-info';
    statisticsDetailInfo.textContent = detailInfo;
    statisticsDetailInfo.style.display = 'inline-block';
    statisticsDetailInfo.style.position = 'absolute';
    statisticsDetailInfo.style.top = '110%';
    statisticsDetailInfo.style.right = '0';
    statisticsDetailInfo.style.backgroundColor = 'rgba(0, 0, 0, 0.8)';
    statisticsDetailInfo.style.color = 'white';
    statisticsDetailInfo.style.padding = '8px';
    statisticsDetailInfo.style.borderRadius = '4px';
    statisticsDetailInfo.style.fontSize = '12px';
    statisticsDetailInfo.style.whiteSpace = 'pre-wrap';
    statisticsDetailInfo.style.lineHeight = 'normal';
    statisticsDetailInfo.style.verticalAlign = 'center';
    statisticsDetailInfo.style.zIndex = '1000';
    statisticsDetailInfo.style.width = 'max-content';

    statisticsButton.appendChild(statisticsDetailInfo);
  }

  /** 隐藏统计信息详情 */
  function hideStatisticsDetailInfo(statisticsButton) {
    const statisticsDetailInfo = statisticsButton.querySelector('#statistics-detail-info');
    if (statisticsDetailInfo) {
      statisticsButton.removeChild(statisticsDetailInfo);
    }
  }

  /** =================================== 节假日烟花 ============================================ */

  /** 添加放烟花按钮 */
  function addFireworksBtn() {
    return new Promise((resolve) => {
      if (!settings.showFireworksBtn) return resolve(true);

      var { navBar, exists } = hasNavBarItem(nav_fireworks_btn_id);
      if (!navBar || exists) return resolve(exists);

      var liItem = document.createElement('li');
      liItem.id = nav_fireworks_btn_id;
      liItem.className = 'ng-star-inserted';
      liItem.style = 'display: inline-block; vertical-align: middle;';

      var fireworksButton = document.createElement('button');
      fireworksButton.textContent = '放烟花';
      fireworksButton.style.backgroundColor = 'transparent';
      fireworksButton.style.color = 'white';
      fireworksButton.style.border = 'none';
      fireworksButton.style.textAlign = 'center';
      fireworksButton.style.textDecoration = 'none';
      fireworksButton.style.display = 'inline-block';
      fireworksButton.style.fontSize = '14px';
      fireworksButton.style.cursor = 'pointer';

      fireworksButton.onmouseover = function () {
        fireworksButton.style.backgroundColor = 'hsla(0, 0%, 100%, .2)';
      };
      fireworksButton.onmouseout = function () {
        fireworksButton.style.backgroundColor = 'transparent';
      };

      fireworksButton.onclick = function () {
        startFireworks(settings.fireworksText || defaultSettings.fireworksText);
      };

      fireworksButton.oncontextmenu = function (e) {
        e.preventDefault();
        showFireworksTextDialog();
      };

      liItem.appendChild(fireworksButton);
      navBar.insertBefore(liItem, navBar.firstChild);

      resolve(true);
    });
  }

  /** 设置烟花文案弹窗 */
  function showFireworksTextDialog() {
    const overlay = document.createElement('div');
    overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.45);z-index:9998;display:flex;align-items:center;justify-content:center;';

    const dialog = document.createElement('div');
    dialog.style.cssText = 'background:#fff;border-radius:8px;box-shadow:0 6px 30px rgba(0,0,0,.15);padding:24px;width:400px;';

    const title = document.createElement('div');
    title.style.cssText = 'font-size:16px;font-weight:600;color:#222;margin-bottom:16px;';
    title.textContent = '设置烟花文案';
    dialog.appendChild(title);

    const { inputElement, input } = createInputElement({ placeholder: '请输入烟花文案', value: settings.fireworksText });
    dialog.appendChild(inputElement);

    const btnContainer = document.createElement('div');
    btnContainer.style.cssText = 'text-align:right;margin-top:16px;padding-top:12px;border-top:1px solid #f0f0f0;';

    const cancelBtn = createButtonElement({
      title: '取消',
      onClick: () => document.body.removeChild(overlay),
    });

    const confirmBtn = createButtonElement({
      title: '确认',
      type: 'primary',
      onClick: () => {
        const text = input.value.trim();
        if (!text) {
          toast('文案不能为空');
          return;
        }
        const _settings = { ...settings };
        _settings.fireworksText = text;
        settings = _settings;
        setConfig('settings', _settings);
        document.body.removeChild(overlay);
        toast('保存成功');
      },
    });

    btnContainer.appendChild(cancelBtn);
    btnContainer.appendChild(confirmBtn);
    dialog.appendChild(btnContainer);

    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    input.focus();
    input.select();
  }

  /** 检查并播放烟花 */
  async function checkAndShowFireworks() {
    if (!settings.showFireworks) return;

    const currentUrl = getCurrentUrl();
    if (!currentUrl || (!currentUrl.includes('/portal/index') && !currentUrl.includes('/auth-callback'))) return;

    const today = new Date();
    const todayStr = formatDate(today);
    const dayOfWeek = today.getDay();

    let greeting = '';
    let holidayData = null;

    // 获取节假日数据
    try {
      const res = await requestGM({ url: holidayUrl, method: 'GET' });
      holidayData = JSON.parse(res || '{}');
    } catch (e) {
      log('【烟花】', '获取节假日数据失败', e);
    }

    if (holidayData) {
      const holidays = holidayData.holidays || {};
      const workdays = holidayData.workdays || {};

      // 节假日当天
      if (holidays[todayStr]) {
        const holidayType = holidays[todayStr].split(',') || [];
        const holidayName = holidayType[1] || holidayType[0] || '';
        if (holidayName) {
          greeting = `${holidayName}快乐`;
        }
      }

      // 周末（排除调休上班日）
      if (!greeting && (dayOfWeek === 0 || dayOfWeek === 6) && !workdays[todayStr]) {
        greeting = '周末快乐';
      }

      // 节假日放假前的最后一个工作日
      if (!greeting) {
        const tomorrow = new Date(today);
        tomorrow.setDate(tomorrow.getDate() + 1);
        const tomorrowStr = formatDate(tomorrow);
        const tomorrowDayOfWeek = tomorrow.getDay();
        if (holidays[tomorrowStr]) {
          // 明天是法定节假日
          const holidayType = holidays[tomorrowStr].split(',') || [];
          const holidayName = holidayType[1] || holidayType[0] || '';
          if (holidayName) {
            greeting = `${holidayName}快乐`;
          }
        } else if ((tomorrowDayOfWeek === 0 || tomorrowDayOfWeek === 6) && !workdays[tomorrowStr]) {
          // 明天是周末（非调休上班日）
          greeting = '周末快乐';
        }
      }
    } else if (dayOfWeek === 0 || dayOfWeek === 6) {
      // 获取节假日数据失败时，仅检测周末
      greeting = '周末快乐';
    }

    if (greeting) {
      setTimeout(() => startFireworks(greeting), 500);
    }
  }

  /** 播放烟花动画 */
  function startFireworks(text) {
    // 创建全屏容器
    const container = document.createElement('div');
    container.id = 'fireworks-container';
    container.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;z-index:99999;pointer-events:none;';
    container.oncontextmenu = function (e) {
      e.preventDefault();
    };

    const canvas = document.createElement('canvas');
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;
    canvas.style.cssText = 'display:block;';
    container.appendChild(canvas);
    document.body.appendChild(container);

    const ctx = canvas.getContext('2d');
    const W = canvas.width;
    const H = canvas.height;
    const startTime = Date.now();
    const duration = 14000;
    const colors = ['#ff0000', '#ff7700', '#ffff00', '#00ff00', '#00bfff', '#ff00ff', '#ffd700', '#ff1493', '#ff4500'];

    // ========== 文字粒子 ==========

    /** 获取文字像素位置 */
    function getTextPixels(text, maxW, maxH) {
      const offCanvas = document.createElement('canvas');
      offCanvas.width = Math.ceil(maxW);
      offCanvas.height = Math.ceil(maxH);
      const offCtx = offCanvas.getContext('2d');

      let fontSize = 110;
      offCtx.textAlign = 'center';
      offCtx.textBaseline = 'middle';
      offCtx.font = `bold ${fontSize}px "Microsoft YaHei","PingFang SC",sans-serif`;

      while (offCtx.measureText(text).width > maxW * 0.9 && fontSize > 20) {
        fontSize -= 2;
        offCtx.font = `bold ${fontSize}px "Microsoft YaHei","PingFang SC",sans-serif`;
      }

      offCtx.fillStyle = '#fff';
      offCtx.fillText(text, offCanvas.width / 2, offCanvas.height / 2);

      const imageData = offCtx.getImageData(0, 0, offCanvas.width, offCanvas.height);
      const pixels = [];
      const step = 3;
      const cx = offCanvas.width / 2;
      const cy = offCanvas.height / 2;

      for (let y = 0; y < offCanvas.height; y += step) {
        for (let x = 0; x < offCanvas.width; x += step) {
          if (imageData.data[(y * offCanvas.width + x) * 4 + 3] > 128) {
            pixels.push({ x: x - cx, y: y - cy });
          }
        }
      }
      return pixels;
    }

    const textPixels = getTextPixels(text, W * 0.85, 160);
    const textParticles = textPixels.map((p) => ({
      homeX: W / 2 + p.x,
      homeY: H / 2 + p.y,
      x: Math.random() * W,
      y: Math.random() * H,
      vx: 0,
      vy: 0,
      alpha: 1,
      size: 1.5 + Math.random() * 1.5,
      hue: Math.floor(Math.random() * 360), // 彩色
    }));

    // ========== 背景烟花粒子 ==========

    const bgParticles = [];
    const rockets = [];
    const confetti = [];

    function launchRocket() {
      const x = Math.random() * W * 0.8 + W * 0.1;
      const targetY = Math.random() * H * 0.35 + H * 0.1;
      const color = colors[Math.floor(Math.random() * colors.length)];
      rockets.push({ x, y: H, targetY, speed: 5 + Math.random() * 5, color, trail: [] });
    }

    function explodeRocket(x, y, color) {
      const count = 50 + Math.floor(Math.random() * 50);
      for (let i = 0; i < count; i++) {
        const angle = Math.random() * Math.PI * 2;
        const speed = Math.random() * 5 + 2;
        bgParticles.push({
          x,
          y,
          vx: Math.cos(angle) * speed,
          vy: Math.sin(angle) * speed,
          color: Math.random() > 0.3 ? color : colors[Math.floor(Math.random() * colors.length)],
          alpha: 1,
          size: 1.5 + Math.random() * 2.5,
          decay: 0.008 + Math.random() * 0.015,
          gravity: 0.04,
        });
      }
    }

    function hexToRgb(hex) {
      return `${parseInt(hex.slice(1, 3), 16)},${parseInt(hex.slice(3, 5), 16)},${parseInt(hex.slice(5, 7), 16)}`;
    }

    /** 生成彩带 */
    function spawnConfetti(count = 10) {
      const confettiColors = ['#ff0000', '#ff7700', '#ffff00', '#00ff00', '#00bfff', '#ff00ff', '#ffd700', '#ff1493', '#ff4500', '#7cfc00', '#00ffff'];
      for (let i = 0; i < count; i++) {
        confetti.push({
          x: Math.random() * W,
          y: -10 - Math.random() * 20,
          w: 4 + Math.random() * 8,
          h: 2 + Math.random() * 4,
          color: confettiColors[Math.floor(Math.random() * confettiColors.length)],
          vx: (Math.random() - 0.5) * 1.5,
          vy: 1 + Math.random() * 2.5,
          rot: Math.random() * Math.PI * 2,
          rotSpeed: (Math.random() - 0.5) * 0.1,
          swing: (Math.random() - 0.5) * 0.5,
          phase: Math.random() * Math.PI * 2,
          alpha: 0.7 + Math.random() * 0.3,
        });
      }
    }

    // ========== 文字粒子相位控制 ==========

    let phase = 'form'; // form → hold → explode → form → hold → explode → fade
    let phaseStart = Date.now();
    const phaseTimes = { form: 1800, hold: 2000, explode: 1600 };
    let cycleCount = 0;
    const maxCycles = 2;
    let lastLaunch = 0;
    let fading = false;

    function resetParticlesForReform() {
      textParticles.forEach((p) => {
        p.x = Math.random() * W;
        p.y = Math.random() * H;
        p.vx = 0;
        p.vy = 0;
        p.alpha = 1;
      });
    }

    function animate() {
      const elapsed = Date.now() - startTime;

      // 超时淡出
      if (elapsed > duration && !fading) {
        fading = true;
        container.style.opacity = '0';
        container.style.transition = 'opacity 1.5s ease';
        setTimeout(() => {
          if (document.body.contains(container)) document.body.removeChild(container);
        }, 1500);
      }

      // 相位切换
      const phaseElapsed = Date.now() - phaseStart;
      if (!fading && phaseElapsed > phaseTimes[phase]) {
        if (phase === 'form') {
          phase = 'hold';
          phaseStart = Date.now();
        } else if (phase === 'hold') {
          phase = 'explode';
          phaseStart = Date.now();
          textParticles.forEach((p) => {
            const angle = Math.atan2(p.homeY - H / 2, p.homeX - W / 2) + (Math.random() - 0.5) * 0.5;
            const speed = 6 + Math.random() * 10;
            p.vx = Math.cos(angle) * speed;
            p.vy = Math.sin(angle) * speed;
          });
        } else if (phase === 'explode') {
          cycleCount++;
          if (cycleCount >= maxCycles) {
            fading = true;
            container.style.opacity = '0';
            container.style.transition = 'opacity 1.5s ease';
            setTimeout(() => {
              if (document.body.contains(container)) document.body.removeChild(container);
            }, 1500);
          } else {
            phase = 'form';
            phaseStart = Date.now();
            resetParticlesForReform();
          }
        }
      }

      // ----- 更新文字粒子 -----
      if (phase === 'form') {
        textParticles.forEach((p) => {
          const dx = p.homeX - p.x;
          const dy = p.homeY - p.y;
          const dist = Math.sqrt(dx * dx + dy * dy);
          if (dist > 1) {
            const force = Math.min(0.08, (1 / (dist * 0.01 + 1)) * 0.06);
            p.vx += dx * force;
            p.vy += dy * force;
          }
          p.vx *= 0.88;
          p.vy *= 0.88;
          p.x += p.vx;
          p.y += p.vy;
          p.alpha = Math.min(1, p.alpha + 0.02);
        });
      } else if (phase === 'hold') {
        textParticles.forEach((p) => {
          p.x = p.homeX + (Math.random() - 0.5) * 1.5;
          p.y = p.homeY + (Math.random() - 0.5) * 1.5;
          p.alpha = 0.85 + Math.random() * 0.15;
        });
      } else if (phase === 'explode') {
        textParticles.forEach((p) => {
          p.vx *= 0.97;
          p.vy *= 0.97;
          p.x += p.vx;
          p.y += p.vy;
          p.vy += 0.05;
          p.alpha = Math.max(0, 1 - (Date.now() - phaseStart) / phaseTimes.explode);
        });
      }

      // ----- 更新背景烟花 & 彩带 -----
      if (!fading && elapsed - lastLaunch > 300 + Math.random() * 500) {
        launchRocket();
        lastLaunch = elapsed;
        if (Math.random() > 0.5) setTimeout(launchRocket, 100 + Math.random() * 200);
      }

      // 持续生成彩带
      if (!fading && Math.random() > 0.85) {
        spawnConfetti(3 + Math.floor(Math.random() * 5));
      }

      for (let i = rockets.length - 1; i >= 0; i--) {
        const r = rockets[i];
        r.y -= r.speed;
        r.trail.push({ x: r.x, y: r.y });
        if (r.trail.length > 12) r.trail.shift();
        if (r.y <= r.targetY) {
          explodeRocket(r.x, r.y, r.color);
          rockets.splice(i, 1);
        }
      }

      for (let i = bgParticles.length - 1; i >= 0; i--) {
        const p = bgParticles[i];
        p.x += p.vx;
        p.y += p.vy;
        p.vy += p.gravity;
        p.vx *= 0.98;
        p.alpha -= p.decay;
        if (p.alpha <= 0) bgParticles.splice(i, 1);
      }

      // 更新彩带
      for (let i = confetti.length - 1; i >= 0; i--) {
        const c = confetti[i];
        c.x += c.vx + Math.sin(c.phase) * c.swing;
        c.y += c.vy;
        c.rot += c.rotSpeed;
        c.phase += 0.02;
        if (c.y > H + 20) confetti.splice(i, 1);
      }

      // ----- 绘制 -----
      ctx.clearRect(0, 0, W, H);

      // 火箭尾迹
      rockets.forEach((r) => {
        for (let i = 0; i < r.trail.length; i++) {
          const a = (i / r.trail.length) * 0.8;
          ctx.beginPath();
          ctx.arc(r.trail[i].x, r.trail[i].y, 2 * (i / r.trail.length), 0, Math.PI * 2);
          ctx.fillStyle = `rgba(255,255,255,${a})`;
          ctx.fill();
        }
      });

      // 背景烟花粒子
      bgParticles.forEach((p) => {
        ctx.beginPath();
        ctx.arc(p.x, p.y, p.size * p.alpha, 0, Math.PI * 2);
        ctx.fillStyle = `rgba(${hexToRgb(p.color)},${Math.max(0, p.alpha * 0.6)})`;
        ctx.fill();
      });

      // 绘制彩带
      confetti.forEach((c) => {
        ctx.save();
        ctx.translate(c.x, c.y);
        ctx.rotate(c.rot);
        ctx.globalAlpha = c.alpha;
        ctx.fillStyle = c.color;
        ctx.fillRect(-c.w / 2, -c.h / 2, c.w, c.h);
        ctx.restore();
      });

      // 文字粒子（带发光效果）
      textParticles.forEach((p) => {
        if (p.alpha <= 0.01) return;
        ctx.beginPath();
        ctx.arc(p.x, p.y, p.size * Math.max(0.3, p.alpha), 0, Math.PI * 2);
        ctx.fillStyle = `hsla(${p.hue},100%,60%,${Math.max(0, p.alpha * 0.9)})`;
        ctx.fill();
        if (p.alpha > 0.5) {
          ctx.beginPath();
          ctx.arc(p.x, p.y, p.size * 2.5 * p.alpha, 0, Math.PI * 2);
          ctx.fillStyle = `hsla(${p.hue},100%,70%,${Math.max(0, p.alpha * 0.15)})`;
          ctx.fill();
        }
      });

      requestAnimationFrame(animate);
    }

    animate();

    // 点击关闭
    container.style.pointerEvents = 'auto';
    container.addEventListener('click', () => {
      if (!document.body.contains(container)) return;
      container.style.opacity = '0';
      container.style.transition = 'opacity 0.5s ease';
      setTimeout(() => {
        if (document.body.contains(container)) document.body.removeChild(container);
      }, 500);
    });
  }

  /** =================================== 每日一签 ============================================ */

  /** 添加每日一签按钮 */
  function addLotBtn() {
    return new Promise((resolve) => {
      if (!settings.showLotBtn) return resolve(true);

      var { navBar, exists } = hasNavBarItem(nav_lot_btn_id);
      if (!navBar || exists) return resolve(exists);

      var liItem = document.createElement('li');
      liItem.id = nav_lot_btn_id;
      liItem.className = 'ng-star-inserted';
      liItem.style = 'display: inline-block; vertical-align: middle;line-height: 100%; min-width: 50px; padding: 8px 2px; text-align: center; border-radius: 2px; cursor: pointer;';

      var lotButton = document.createElement('button');
      lotButton.style.backgroundColor = 'transparent';
      lotButton.style.color = 'white';
      lotButton.style.border = 'none';
      lotButton.style.textAlign = 'center';
      lotButton.style.textDecoration = 'none';
      lotButton.style.display = 'inline-flex';
      lotButton.style.flexDirection = 'column';
      lotButton.style.alignItems = 'center';
      lotButton.style.justifyContent = 'center';
      lotButton.style.fontSize = '14px';
      lotButton.style.cursor = 'pointer';
      lotButton.style.padding = '2px 6px';
      lotButton.style.lineHeight = '1';

      // 红色签牌 SVG 图标
      var lotBtnIcon =
        '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="22" viewBox="0 0 24 26" fill="none" style="display:block;">' +
        '<rect x="4" y="5" width="16" height="18" rx="2" fill="#d32f2f" stroke="#b71c1c" stroke-width="0.8"/>' +
        '<rect x="6" y="4" width="12" height="2.5" rx="1.25" fill="#c62828"/>' +
        '<path d="M12 4v-2" stroke="#d32f2f" stroke-width="1.2" stroke-linecap="round"/>' +
        '<circle cx="12" cy="10.5" r="2" fill="none" stroke="#f5c542" stroke-width="0.8"/>' +
        '<line x1="7" y1="15.5" x2="17" y2="15.5" stroke="#f5c542" stroke-width="0.8" stroke-linecap="round" opacity="0.6"/>' +
        '<line x1="7" y1="18" x2="13" y2="18" stroke="#f5c542" stroke-width="0.8" stroke-linecap="round" opacity="0.6"/>' +
        '</svg>';

      var lotBtnResult = document.createElement('span');
      lotBtnResult.id = 'lotBtnResult';
      lotBtnResult.style.cssText = 'display:none;font-size:11px;color:#FFD700;line-height:1.3;margin-top:2px;white-space:nowrap;';

      // 检查当天是否已抽签
      var savedDailyLot = getConfig('dailyLot');
      if (savedDailyLot && savedDailyLot.date === new Date().toISOString().slice(0, 10)) {
        lotBtnResult.textContent = savedDailyLot.lot.level;
        lotBtnResult.style.display = 'block';
      } else {
        lotBtnResult.textContent = '抽签';
        lotBtnResult.style.display = 'block';
      }

      lotButton.innerHTML = lotBtnIcon;
      lotButton.appendChild(lotBtnResult);

      // 悬浮提示工具条
      var tooltip = document.createElement('div');
      tooltip.id = 'lotTooltip';
      tooltip.style.cssText =
        'display:none;position:fixed;z-index:10000;background:linear-gradient(135deg,#8B0000 0%,#CD853F 50%,#8B0000 100%);border:2px solid #FFD700;border-radius:12px;padding:16px 18px;' +
        'box-shadow:0 8px 30px rgba(0,0,0,0.5);text-align:center;min-width:220px;max-width:280px;pointer-events:none;';
      var tooltipLevel = document.createElement('div');
      tooltipLevel.style.cssText = 'font-size:22px;font-weight:bold;margin-bottom:4px;';
      var tooltipType = document.createElement('div');
      tooltipType.style.cssText = 'font-size:13px;color:#FFF8DC;margin-bottom:8px;';
      var tooltipContent = document.createElement('div');
      tooltipContent.style.cssText = 'font-size:14px;color:#fff;line-height:1.5;padding:8px;background:rgba(0,0,0,0.2);border-radius:8px;border:1px solid rgba(255,215,0,0.5);';
      var tooltipExplain = document.createElement('div');
      tooltipExplain.style.cssText = 'font-size:12px;color:#FFD700;line-height:1.5;margin-top:8px;';
      tooltip.appendChild(tooltipLevel);
      tooltip.appendChild(tooltipType);
      tooltip.appendChild(tooltipContent);
      tooltip.appendChild(tooltipExplain);
      document.body.appendChild(tooltip);

      function updateTooltip() {
        var saved = getConfig('dailyLot');
        if (saved && saved.date === new Date().toISOString().slice(0, 10)) {
          var lot = saved.lot;
          tooltipLevel.textContent = lot.level;
          tooltipLevel.style.color = '#FFD700';
          if (lot.level === '上上签') tooltipLevel.style.textShadow = '0 0 20px rgba(255,215,0,0.8)';
          else if (lot.level === '上吉签') {
            tooltipLevel.style.color = '#FFA500';
            tooltipLevel.style.textShadow = '0 0 15px rgba(255,165,0,0.6)';
          } else if (lot.level === '中吉签') {
            tooltipLevel.style.color = '#FFD700';
            tooltipLevel.style.opacity = '0.9';
            tooltipLevel.style.textShadow = 'none';
          } else {
            tooltipLevel.style.color = '#CD853F';
            tooltipLevel.style.opacity = '0.8';
            tooltipLevel.style.textShadow = 'none';
          }
          tooltipType.textContent = lot.type + '签';
          tooltipContent.textContent = lot.content;
          tooltipExplain.textContent = lot.explain;
          return true;
        }
        return false;
      }

      lotButton.onmouseover = function () {
        lotButton.style.backgroundColor = 'hsla(0, 0%, 100%, .2)';
        lotButton.style.borderRadius = '4px';
        if (updateTooltip()) {
          var rect = lotButton.getBoundingClientRect();
          var tipWidth = 260;
          var left = rect.left + rect.width / 2 - tipWidth / 2;
          if (left + tipWidth > window.innerWidth - 8) left = window.innerWidth - tipWidth - 8;
          if (left < 8) left = 8;
          tooltip.style.display = 'block';
          tooltip.style.left = left + 'px';
          tooltip.style.top = rect.bottom + 8 + 'px';
          tooltip.style.transform = 'none';
        }
      };
      lotButton.onmouseout = function () {
        lotButton.style.backgroundColor = 'transparent';
        tooltip.style.display = 'none';
      };

      lotButton.onclick = function () {
        showLottery();
      };

      liItem.appendChild(lotButton);
      navBar.insertBefore(liItem, settings.showStatisticsInfo ? navBar.lastChild : null);

      resolve(true);
    });
  }

  var LOTTERIES = (function () {
    var _d = atob;
    var _u = function (s) {
      return decodeURIComponent(
        Array.prototype.map
          .call(_d(s), function (c) {
            return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
          })
          .join(''),
      );
    };
    return JSON.parse(
      _u(
        'W3sibGV2ZWwiOiLkuIrkuIrnrb4iLCJ0eXBlIjoi6LSi6L+QIiwiY29udGVudCI6Iui0oua6kOa7mua7mu+8jOWFq+aWueadpei0ou+8jOWygeWygeW5s+WuiSIsImV4cGxhaW4iOiLku4rml6XotKLov5DkuqjpgJrvvIzml6DorrrmraPotKLov5jmmK/lgY/otKLpg73mnInlpb3mtojmga/jgILmipXotYTnkIbotKLlj6/lvpfkuLDljprlm57miqXvvIzlu7rorq7miormj6Hml7bmnLrvvIzotKLmupDlub/ov5vjgIIifSx7ImxldmVsIjoi5LiK5LiK562+IiwidHlwZSI6IuW5s+WuiSIsImNvbnRlbnQiOiLlubPlronlkInnpaXvvIzkuIfkuovlpoLmhI/vvIzlv4Pmg7PkuovmiJAiLCJleHBsYWluIjoi5LuK5pel5bmz5a6J6aG66YGC77yM5peg54G+5peg6Zq+44CC5pen5Y6E5bey5Y6777yM5paw5oSB5LiN5L6177yM6ZiW5a625a6J5bq377yM56aP5rO957u16ZW/44CCIn0seyJsZXZlbCI6IuS4iuS4iuetviIsInR5cGUiOiLkuovkuJoiLCJjb250ZW50Ijoi5LqL5Lia6IW+6aOe77yM5q2l5q2l6auY5Y2H77yM5YmN56iL5Ly86ZSmIiwiZXhwbGFpbiI6IuS6i+S4muS4iuWwhui/juadpemHjeWkp+acuumBh++8jOaciei0teS6uuebuOWKqeOAguaCqOeahOaJjeiDveWwhuW+l+WIsOWFheWIhuWPkeaMpe+8jOWNh+iBjOWKoOiWquWcqOacm+OAgiJ9LHsibGV2ZWwiOiLkuIrkuIrnrb4iLCJ0eXBlIjoi5ae757yYIiwiY29udGVudCI6IuWkqei1kOiJr+e8mO+8jOWWnOe7k+i/nueQhu+8jOW5uOemj+e+jua7oSIsImV4cGxhaW4iOiLljZXouqvogIXmnInmnJvpgYfliLDlkb3kuK3ms6jlrprnmoTkvLTkvqPvvIzmnInkvLTkvqPogIXmhJ/mg4XlsIbmm7TliqDnlJzonJzjgILlrrbluq3lkoznnabvvIzlubjnpo/nvo7mu6HjgIIifSx7ImxldmVsIjoi5LiK5LiK562+IiwidHlwZSI6IuWBpeW6tyIsImNvbnRlbnQiOiLouqvlvLrlipvlo67vvIznsr7npZ7mipbmk57vvIznmb7nl4XkuI3kvrUiLCJleHBsYWluIjoi6Lqr5L2T54q25Ya15p6B5L2z77yM57K+5Yqb5YWF5rKb44CC5pen55a+5bC95Y6777yM5paw5oGZ5LiN55Sf77yM5q2j5piv6L+Q5Yqo5YW755Sf55qE5aW95pe25py644CCIn0seyJsZXZlbCI6IuS4iuS4iuetviIsInR5cGUiOiLlrabkuJoiLCJjb250ZW50Ijoi6YeR5qac6aKY5ZCN77yM5a2m5Lia5pyJ5oiQ77yM5YmN56iL5LiH6YeMIiwiZXhwbGFpbiI6IuWtpuS4mui/kOaegeS9s++8jOiAg+ivlei/kOaXuuOAguWLpOWli+WIu+iLpue7iOacieWbnuaKpe+8jOaZuuaFp+S5i+mXqOS4uuS9oOaVnuW8gOOAgiJ9LHsibGV2ZWwiOiLkuIrkuIrnrb4iLCJ0eXBlIjoi5Lq66ZmFIiwiY29udGVudCI6Iui0teS6uuebiOmXqO+8jOW3puWPs+mAoua6kO+8jOS8l+acm+aJgOW9kiIsImV4cGxhaW4iOiLkurrpmYXlhbPns7vpobrpgYLvvIzotLXkurrov5DlvLrlirLjgILlnKjnpL7kuqTlnLrlkIjkuK3lsIblpKfmlL7lvILlvanvvIzojrflvpfku5bkurrorqTlj6/kuI7mlK/mjIHjgIIifSx7ImxldmVsIjoi5LiK5ZCJ562+IiwidHlwZSI6Iui0oui/kCIsImNvbnRlbnQiOiLlpKflkInlpKfliKnvvIzotKLluJvkuLDnm4jvvIzml6Xnp6/mnIjlr4wiLCJleHBsYWluIjoi6LSi6L+Q5LiN6ZSZ77yM6Jm95peg5pq05a+M5LmL6LGh77yM5L2G57uG5rC06ZW/5rWB44CC5Yuk5L+t5oyB5a6277yM55CG6LSi5pyJ6YGT77yM6LSi5bib6Ieq54S25Liw55uI44CCIn0seyJsZXZlbCI6IuS4iuWQieetviIsInR5cGUiOiLlubPlrokiLCJjb250ZW50Ijoi5bmz5a6J5peg5LqL77yM56aP5rCU5ruh5ruh77yM5Zac5rCU55uI6ZeoIiwiZXhwbGFpbiI6IuW5s+WuieaXpeWtkOeahOWlveaXtuWFie+8jOaXoOWkp+eBvuWkp+mavuOAguemj+awlOa7oea7oe+8jOWutuW6reWWnOawlOebiOmXqO+8jOePjeaDnOW9k+S4i+e+juWlveOAgiJ9LHsibGV2ZWwiOiLkuIrlkInnrb4iLCJ0eXBlIjoi5LqL5LiaIiwiY29udGVudCI6IuS6i+S4mumhuumBgu+8jOW+l+W/g+W6lOaJi++8jOWJjeeoi+S8vOmUpiIsImV4cGxhaW4iOiLlt6XkvZzov5vlsZXpobrliKnvvIzog73lipvlvpfliLDorqTlj6/jgILomb3mnInlsI/ms6LmipjvvIzkvYblnYfog73ljJbpmankuLrlpLfvvIznqLPmraXlkJHliY3jgIIifSx7ImxldmVsIjoi5LiK5ZCJ562+IiwidHlwZSI6IuWnu+e8mCIsImNvbnRlbnQiOiLoirHlpb3mnIjlnIbvvIzmg4XmipXmhI/lkIjvvIznlJzonJzmuKnppqgiLCJleHBsYWluIjoi5oSf5oOF6L+Q5Yq/6Imv5aW977yM5Y2V6Lqr6ICF5pyJ57yY6YGH6KeB5b+D5Luq5LmL5Lq677yM5pyJ5Ly05L6j6ICF5oSf5oOF55Sc6Jyc44CC5LqS5pWs5LqS54ix77yM5bm456aP6ZW/5LmF44CCIn0seyJsZXZlbCI6IuS4iuWQieetviIsInR5cGUiOiLlgaXlurciLCJjb250ZW50Ijoi6Lqr5by65L2T5YGl77yM57K+56We55+N6ZOE77yM5peg55eF5peg54G+IiwiZXhwbGFpbiI6Iui6q+S9k+eKtuaAgeiJr+Wlve+8jOeyvuelnumlsea7oeOAguazqOaEj+WKs+mAuOe7k+WQiO+8jOmAguW9k+mUu+eCvO+8jOWBpeW6t+eKtuaAgeWwhui2iuadpei2iuWlveOAgiJ9LHsibGV2ZWwiOiLkuIrlkInnrb4iLCJ0eXBlIjoi5Ye66KGMIiwiY29udGVudCI6IuWHuuihjOmhuuWIqe+8jOS4gOi3r+mhuumjju+8jOW5s+WuieW+gOi/lCIsImV4cGxhaW4iOiLlh7rooYzov5Dlir/oia/lpb3vvIzml4XpgJTpobrliKnjgILml6Dorrrlh7rlt67ov5jmmK/ml4XooYzvvIzpg73lsIblubPlronlvoDov5TvvIzkuIDot6/pobrpo47jgIIifSx7ImxldmVsIjoi5LiK5ZCJ562+IiwidHlwZSI6IuWutuWuhSIsImNvbnRlbnQiOiLlrrblroXlhbTml7rvvIzlkoznnabmuKnppqjvvIznpo/ms73nu7Xplb8iLCJleHBsYWluIjoi5a625a6F6L+Q5Yq/5YW05pe677yM5a625bqt5oiQ5ZGY5ZKM552m55u45aSE44CC5rip6aao5ZKM6LCQ55qE5a625bqt5rCb5Zu05bCG5bim5p2l57u157u156aP5rO944CCIn0seyJsZXZlbCI6IuS4reWQieetviIsInR5cGUiOiLotKLov5AiLCJjb250ZW50Ijoi6LSi6L+Q5bmz56iz77yM6YeP5YWl5Li65Ye677yM56ev5bCR5oiQ5aSaIiwiZXhwbGFpbiI6Iui0oui/kOW5s+eos+aZrumAmu+8jOS4jeWunOaKleacuuWGkumZqeOAgumHj+WFpeS4uuWHuu+8jOiKguS/reaMgeWutu+8jOenr+WwkeaIkOWkmuS5n+aYr+S4gOeslOi0ouWvjOOAgiJ9LHsibGV2ZWwiOiLkuK3lkInnrb4iLCJ0eXBlIjoi5bmz5a6JIiwiY29udGVudCI6IuW5s+W5s+WuieWuie+8jOmhuumhuuWIqeWIqe+8jOWugemdmeiHtOi/nCIsImV4cGxhaW4iOiLlubPlronov5DlsJrlj6/vvIzml6DlpKfngb7lpKfpmr7jgILkvYbpnIDms6jmhI/ml6XluLjlronlhajvvIzkuI3lj6/lpKfmhI/jgILlv4PpnZnoh6rnhLblh4nvvIzlubPlronmmK/npo/jgIIifSx7ImxldmVsIjoi5Lit5ZCJ562+IiwidHlwZSI6IuS6i+S4miIsImNvbnRlbnQiOiLli6Tli4nmlazkuJrvvIzohJrouI/lrp7lnLDvvIznqLPmraXliY3ov5siLCJleHBsYWluIjoi5LqL5Lia6ZyA6ISa6LiP5a6e5Zyw77yM5Yuk5YuJ5Yqq5Yqb44CC6Jm95pyJ5Zuw6Zq+5oyR5oiY77yM5L2G5Y+q6KaB5Z2a5oyB5LiN5oeI77yM57uI5Lya5pyJ5omA56qB56C044CCIn0seyJsZXZlbCI6IuS4reWQieetviIsInR5cGUiOiLlp7vnvJgiLCJjb250ZW50Ijoi6aG65YW26Ieq54S277yM5rC05Yiw5rig5oiQ77yM6Z2Z5b6F6Iqx5byAIiwiZXhwbGFpbiI6IuWnu+e8mOiusuaxgue8mOWIhu+8jOS4jeWPr+W8uuaxguOAguiAkOW/g+etieW+he+8jOmhuuWFtuiHqueEtu+8jOe8mOWIhuS8muWmguacn+iAjOiHs+OAguW3suaBi+eIseiAheaEn+aDheW5s+eos+WPkeWxleOAgiJ9LHsibGV2ZWwiOiLkuK3lkInnrb4iLCJ0eXBlIjoi5YGl5bq3IiwiY29udGVudCI6IuWBpeW6t+W5s+W5s++8jOazqOmHjeWFu+eUn++8jOa4kOWFpeS9s+WigyIsImV4cGxhaW4iOiLlgaXlurfnirblhrXkuIDoiKzvvIzpnIDopoHlpJrliqDms6jmhI/jgILlkIjnkIbppa7po5/vvIzpgILlvZPov5DliqjvvIzkv53mjIHoia/lpb3nmoTkvZzmga/kuaDmg6/jgIIifSx7ImxldmVsIjoi5Lit5ZCJ562+IiwidHlwZSI6IuWtpuS4miIsImNvbnRlbnQiOiLli6Tog73ooaXmi5nvvIzmjIHkuYvku6XmgZLvvIznu4jmnInmiYDmiJAiLCJleHBsYWluIjoi5a2m5Lia6L+Q5Yq/5Lit562J77yM6ZyA5LuY5Ye65pu05aSa5Yqq5Yqb44CC5L2G5pyJ5LuY5Ye65b+F5pyJ5Zue5oql77yM5Z2a5oyB5bCx5piv6IOc5Yip44CC5oiS6aqE5oiS6LqB44CCIn0seyJsZXZlbCI6IuS4reWQieetviIsInR5cGUiOiLkurrpmYUiLCJjb250ZW50Ijoi5Lul5ZKM5Li66LS177yM5a695Lul5b6F5Lq677yM5reh5a6a5LuO5a65IiwiZXhwbGFpbiI6IuS6uumZheWFs+ezu+WwmuWPr++8jOS9humcgOazqOaEj+iogOihjOOAguS7peWSjOS4uui0te+8jOWuveS7peW+heS6uu+8jOmBh+WIsOefm+ebvuWGt+mdmeWkhOeQhuOAgiJ9LHsibGV2ZWwiOiLlkInnrb4iLCJ0eXBlIjoi6LSi6L+QIiwiY29udGVudCI6IuWwj+i0oui/m+iii++8jOenr+WwkeaIkOWkmu+8jOe7huawtOmVv+a1gSIsImV4cGxhaW4iOiLotKLov5DlubPmt6HvvIzml6DlpKfnmoTov5votKbjgILkvYblsI/otKLov5DkuI3mlq3vvIznp6/lsJHmiJDlpJrmnInml6Dor7Tlpb3jgILohJrouI/lrp7lnLDnmoTmrLrotKLmmK/lr7zlpITjgIIifSx7ImxldmVsIjoi5ZCJ562+IiwidHlwZSI6IuW5s+WuiSIsImNvbnRlbnQiOiLlronlsYXkuZDkuJrvvIzlsoHmnIjpnZnlpb3vvIznjrDkuJblronnqLMiLCJleHBsYWluIjoi5bmz5a6J6L+Q5pmu6YCa77yM5L2G5rGC56iz5Li65LiK44CC5peg6aOO5peg5rWq5L6/5piv5aW95pel5a2Q77yM5a6J5bGF5LmQ5Lia77yM5Lqr5Y+X5bmz5reh55Sf5rS744CCIn0seyJsZXZlbCI6IuWQieetviIsInR5cGUiOiLkuovkuJoiLCJjb250ZW50Ijoi5pys5YiG5YGa5LqL77yM5oGq5bC96IGM5a6I77yM5peg5oSn5LqO5b+DIiwiZXhwbGFpbiI6IuS6i+S4mui/kOW5s+a3oe+8jOS4jeWunOWPmOWKqOOAguWBmuWlveacrOiBjOW3peS9nO+8jOaBquWwveiBjOWuiO+8jOeUqOW5s+W4uOW/g+WvueW+heW3peS9nOS4reeahOW+l+WkseOAgiJ9LHsibGV2ZWwiOiLlkInnrb4iLCJ0eXBlIjoi5ae757yYIiwiY29udGVudCI6IuW5s+W5s+a3oea3oeaJjeaYr+ecn++8jOe7huawtOmVv+a1geaDheaEj+a3sSIsImV4cGxhaW4iOiLlp7vnvJjov5DlubPmt6HkvYbnnJ/lrp7jgILlubPmt6HkuK3nmoTnnJ/mg4XmnIDkuLrlj6/otLXvvIznu4bmsLTplb/mtYHnmoTmhJ/mg4XmnIDmmK/plb/kuYXjgILnj43mg5znnLzliY3kurrjgIIifSx7ImxldmVsIjoi5ZCJ562+IiwidHlwZSI6IuWBpeW6tyIsImNvbnRlbnQiOiLml6Dnl4Xml6Dngb7lsLHmmK/npo/vvIzlubPlronlgaXlurfmnIDnj43otLUiLCJleHBsYWluIjoi5YGl5bq36L+Q5Yq/5LiA6Iis77yM5L2G5peg5aSn56KN44CC5rOo5oSP5a2j6IqC5Y+Y5YyW77yM5Y+K5pe25aKe5YeP6KGj54mp44CC6Lqr5L2T5YGl5bq35bCx5piv5pyA5aSn55qE56aP5rCU44CCIn0seyJsZXZlbCI6IuWQieetviIsInR5cGUiOiLlh7rooYwiLCJjb250ZW50Ijoi5a6J5q2l5b2T6L2m77yM56iz5Lit5rGC6L+b77yM5bmz5a6J6aG66YGCIiwiZXhwbGFpbiI6IuWHuuihjOi/kOWKv+aZrumAmu+8jOS4jeWunOi/nOihjOOAguWmguaenOW/hemhu+WHuuihjO+8jOW7uuiuruaPkOWJjeWBmuWlveWHhuWkh++8jOazqOaEj+S6pOmAmuWuieWFqOOAgiJ9LHsibGV2ZWwiOiLlkInnrb4iLCJ0eXBlIjoi5a625a6FIiwiY29udGVudCI6Iuefpei2s+W4uOS5kO+8jOWutuWSjOS4h+S6i+WFtO+8jOW5s+WuieaYr+emjyIsImV4cGxhaW4iOiLlrrblroXov5Dlir/lubPnqLPjgILnn6XotrPluLjkuZDvvIzlrrbluq3lkoznnabmnIDkuLrph43opoHjgILlubPlubPmt6Hmt6HmiY3mmK/nnJ/vvIzlrrblroXljLnkuZDjgIIifV0=',
      ),
    );
  })();

  /** 显示抽签弹窗 */
  function showLottery() {
    var todayStr = new Date().toISOString().slice(0, 10);
    var savedLot = getConfig('dailyLot');
    var hasDrawnToday = savedLot && savedLot.date === todayStr;

    // 注入样式
    if (!document.getElementById('lotStyle')) {
      var s = document.createElement('style');
      s.id = 'lotStyle';
      s.textContent =
        '@keyframes lotFadeIn{from{opacity:0;transform:scale(0.8)}to{opacity:1;transform:scale(1)}}' +
        '@keyframes lotShake{0%,100%{transform:translateX(0) rotate(0)}10%{transform:translateX(-14px) rotate(-5deg)}20%{transform:translateX(14px) rotate(5deg)}30%{transform:translateX(-12px) rotate(-4deg)}40%{transform:translateX(12px) rotate(4deg)}50%{transform:translateX(-10px) rotate(-3deg)}60%{transform:translateX(10px) rotate(3deg)}70%{transform:translateX(-8px) rotate(-2deg)}80%{transform:translateX(8px) rotate(2deg)}90%{transform:translateX(-4px) rotate(-1deg)}}' +
        '@keyframes lotSparkle{0%,100%{opacity:0.15}50%{opacity:0.3}}' +
        '@keyframes stickUp{from{bottom:20px}to{bottom:120px}}' +
        '@keyframes slideUp{from{transform:translateY(40px);opacity:0}to{transform:translateY(0);opacity:1}}';
      document.head.appendChild(s);
    }

    var overlay = document.createElement('div');
    overlay.style.cssText =
      'position:fixed;top:0;left:0;width:100%;height:100%;z-index:9998;display:flex;flex-direction:column;align-items:center;justify-content:center;font-family:"STKaiti","KaiTi","Microsoft YaHei",serif;' +
      'background:rgba(0,0,0,0.55);';

    // 主容器
    var container = document.createElement('div');
    container.style.cssText = 'position:relative;text-align:center;z-index:1;';

    // 标题
    var title = document.createElement('div');
    title.style.cssText = 'font-size:36px;color:#fff;text-shadow:0 0 20px rgba(255,215,0,0.6),0 0 40px rgba(255,215,0,0.3);margin-bottom:24px;letter-spacing:4px;font-weight:bold;';
    title.textContent = '每日一签';
    container.appendChild(title);

    // 抽签筒容器
    var cylContainer = document.createElement('div');
    cylContainer.style.cssText = 'position:relative;width:200px;height:280px;margin:10px auto;cursor:pointer;transition:transform 0.3s ease;';
    cylContainer.onmouseenter = function () {
      cylContainer.style.transform = 'scale(1.03)';
    };
    cylContainer.onmouseleave = function () {
      cylContainer.style.transform = 'scale(1)';
    };

    // 签条
    var stick = document.createElement('div');
    stick.id = 'lotStick';
    stick.style.cssText =
      'position:absolute;top:85px;left:50%;transform:translateX(-50%);width:35px;height:160px;' +
      'background:linear-gradient(90deg,#FFD700 0%,#FFF8DC 50%,#FFD700 100%);border-radius:3px;' +
      'box-shadow:0 2px 10px rgba(0,0,0,0.3),inset 0 0 20px rgba(255,215,0,0.3);z-index:20;' +
      'transition:top 0.7s cubic-bezier(0.34,1.56,0.64,1);';
    stick.textContent = '签';
    stick.style.fontSize = '16px';
    stick.style.color = '#8B0000';
    stick.style.fontWeight = 'bold';
    stick.style.lineHeight = '30px';
    stick.style.textAlign = 'center';

    // 签筒顶部（筒口）
    var cylTop = document.createElement('div');
    cylTop.style.cssText =
      'position:absolute;top:50px;left:50%;transform:translateX(-50%);width:140px;height:30px;' +
      'background:linear-gradient(90deg,#8B4513 0%,#CD853F 30%,#FFD700 50%,#CD853F 70%,#8B4513 100%);' +
      'border-radius:50% 50% 0 0;box-shadow:0 5px 15px rgba(0,0,0,0.3);z-index:12;';

    // 签筒主体
    var cylBody = document.createElement('div');
    cylBody.style.cssText =
      'position:absolute;top:75px;left:50%;transform:translateX(-50%);width:120px;height:180px;' +
      'background:linear-gradient(90deg,#8B4513 0%,#CD853F 15%,#FFD700 25%,#FFD700 75%,#CD853F 85%,#8B4513 100%);' +
      'border-radius:0 0 5px 5px;' +
      'box-shadow:inset 0 0 30px rgba(0,0,0,0.3),0 10px 30px rgba(0,0,0,0.5),0 0 50px rgba(255,215,0,0.2);z-index:11;';

    // 筒内阴影
    var cylInner = document.createElement('div');
    cylInner.style.cssText =
      'position:absolute;top:0;left:50%;transform:translateX(-50%);width:100px;height:170px;' +
      'background:linear-gradient(180deg,rgba(139,69,19,0.8) 0%,rgba(0,0,0,0.6) 100%);border-radius:0 0 3px 3px;';

    // 组装签筒
    cylBody.appendChild(cylInner);
    cylContainer.appendChild(stick);
    cylContainer.appendChild(cylTop);
    cylContainer.appendChild(cylBody);

    container.appendChild(cylContainer);

    // 提示文字
    var hint = document.createElement('div');
    hint.style.cssText = 'color:#FFE4B5;font-size:18px;margin-top:24px;text-shadow:0 0 10px rgba(0,0,0,0.8);letter-spacing:2px;';
    hint.textContent = '点击签筒，抽取运势签';
    container.appendChild(hint);

    overlay.appendChild(container);

    // 结果弹窗
    var resultOverlay = document.createElement('div');
    resultOverlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.8);z-index:9999;display:none;align-items:center;justify-content:center;';

    var resultBox = document.createElement('div');
    resultBox.style.cssText =
      'position:relative;background:linear-gradient(135deg,#8B0000 0%,#CD853F 50%,#8B0000 100%);padding:36px 32px;border-radius:20px;text-align:center;max-width:90%;width:400px;' +
      'box-shadow:0 0 50px rgba(255,215,0,0.5),inset 0 0 30px rgba(0,0,0,0.3);border:3px solid #FFD700;animation:slideUp 0.5s ease;';

    var closeResultBtn = document.createElement('div');
    closeResultBtn.textContent = '✕';
    closeResultBtn.style.cssText =
      'position:absolute;top:10px;right:14px;width:28px;height:28px;line-height:28px;text-align:center;font-size:16px;color:#FFD700;cursor:pointer;border-radius:50%;transition:all 0.2s;';
    closeResultBtn.onmouseenter = function () {
      closeResultBtn.style.background = 'rgba(255,255,255,0.15)';
    };
    closeResultBtn.onmouseleave = function () {
      closeResultBtn.style.background = 'transparent';
    };
    closeResultBtn.onclick = function () {
      if (document.body.contains(overlay)) document.body.removeChild(overlay);
    };
    resultBox.appendChild(closeResultBtn);
    var resultLevel = document.createElement('div');
    resultLevel.id = 'lotResultLevel';
    resultLevel.style.cssText = 'font-size:32px;margin-bottom:16px;';

    var resultType = document.createElement('div');
    resultType.id = 'lotResultType';
    resultType.style.cssText = 'font-size:18px;color:#FFF8DC;margin-bottom:12px;';

    var resultContent = document.createElement('div');
    resultContent.id = 'lotResultContent';
    resultContent.style.cssText = 'font-size:20px;color:#fff;line-height:1.6;margin-bottom:16px;padding:16px;background:rgba(0,0,0,0.2);border-radius:10px;border:2px solid #FFD700;';

    var resultExplain = document.createElement('div');
    resultExplain.id = 'lotResultExplain';
    resultExplain.style.cssText = 'font-size:15px;color:#FFD700;line-height:1.6;margin-bottom:24px;';

    var drawAgainBtn = document.createElement('button');
    drawAgainBtn.textContent = '再抽一次';
    drawAgainBtn.style.cssText =
      'padding:12px 40px;font-size:18px;background:linear-gradient(135deg,#FFD700 0%,#FFA500 100%);border:none;border-radius:30px;color:#8B0000;cursor:pointer;font-family:inherit;font-weight:bold;' +
      'box-shadow:0 5px 20px rgba(0,0,0,0.3);transition:all 0.3s ease;';
    drawAgainBtn.onmouseenter = function () {
      drawAgainBtn.style.transform = 'scale(1.05)';
      drawAgainBtn.style.boxShadow = '0 8px 30px rgba(255,215,0,0.5)';
    };
    drawAgainBtn.onmouseleave = function () {
      drawAgainBtn.style.transform = 'scale(1)';
      drawAgainBtn.style.boxShadow = '0 5px 20px rgba(0,0,0,0.3)';
    };

    resultBox.appendChild(resultLevel);
    resultBox.appendChild(resultType);
    resultBox.appendChild(resultContent);
    resultBox.appendChild(resultExplain);
    resultBox.appendChild(drawAgainBtn);
    resultOverlay.appendChild(resultBox);

    overlay.appendChild(resultOverlay);
    document.body.appendChild(overlay);

    /** 填充签文到结果弹窗 */
    function showLotResult(lot) {
      // 更新导航栏按钮结果
      var btnResult = document.getElementById('lotBtnResult');
      if (btnResult) {
        btnResult.textContent = lot.level;
        btnResult.style.display = 'block';
      }

      resultLevel.textContent = lot.level;
      resultLevel.style.color = '#FFD700';
      if (lot.level === '上上签') resultLevel.style.textShadow = '0 0 20px rgba(255,215,0,0.8)';
      else if (lot.level === '上吉签') {
        resultLevel.style.color = '#FFA500';
        resultLevel.style.textShadow = '0 0 15px rgba(255,165,0,0.6)';
      } else if (lot.level === '中吉签') {
        resultLevel.style.color = '#FFD700';
        resultLevel.style.opacity = '0.9';
        resultLevel.style.textShadow = 'none';
      } else {
        resultLevel.style.color = '#CD853F';
        resultLevel.style.opacity = '0.8';
        resultLevel.style.textShadow = 'none';
      }
      resultType.textContent = lot.type + '签';
      resultContent.textContent = lot.content;
      resultExplain.textContent = lot.explain;
      resultOverlay.style.display = 'flex';
    }

    // 如果今天已抽签，直接显示结果
    if (hasDrawnToday) {
      showLotResult(savedLot.lot);
      stick.style.top = '-90px';
    }

    var isDrawing = false;

    /** 抽签 */
    function drawLot() {
      if (isDrawing) return;
      isDrawing = true;

      hint.textContent = '摇签中...';
      cylContainer.style.animation = 'lotShake 0.7s ease-in-out';

      // 400ms 后签条弹出
      setTimeout(function () {
        stick.style.top = '-90px';
        stick.style.zIndex = '20';
      }, 400);

      // 动画结束后显示签文
      setTimeout(function () {
        var idx = Math.floor(Math.random() * LOTTERIES.length);
        var lot = LOTTERIES[idx];

        setConfig('dailyLot', { date: todayStr, lot: lot });
        showLotResult(lot);
        hint.textContent = '点击签筒，抽取运势签';

        // 重置动画
        setTimeout(function () {
          cylContainer.style.animation = '';
        }, 100);
      }, 1400);
    }

    /** 再抽一次 */
    function drawAgain() {
      resultOverlay.style.display = 'none';
      stick.style.top = '85px';
      stick.style.zIndex = '20';
      isDrawing = false;
    }

    // 事件绑定
    cylContainer.addEventListener('click', drawLot);
    drawAgainBtn.addEventListener('click', drawAgain);

    // 点击遮罩关闭
    overlay.addEventListener('click', function (e) {
      if (e.target === overlay && document.body.contains(overlay)) document.body.removeChild(overlay);
      else if (e.target === resultOverlay && document.body.contains(overlay)) {
        document.body.removeChild(overlay);
      }
    });
  }

  /** =================================== 表单工具 ============================================ */

  /** 检查页面是否为表单编辑页面 */
  function isFormPage(templateId) {
    const currentUrl = getCurrentUrl();
    if (!currentUrl || !currentUrl.length) return false;

    if (templateId && templateId.length) {
      return currentUrl.includes(`templateId=${templateId}`) || currentUrl.includes(`versionId=${templateId}`);
    }

    return currentUrl.includes('templateId=');
  }

  /** 获取页面表单 */
  function getPageForm() {
    // 是否为编辑模式
    const operatorBar = document.querySelector('#start-operator');
    if (!operatorBar || !operatorBar.outerHTML.includes('提交申请')) {
      log('【检查页面】', '非表单编辑页面');
      return null;
    }

    // 获取表单
    const formcontent = queryIFrameFormContent(document);
    if (!formcontent) {
      log('【检查页面】', '未找到表单');
      return null;
    }

    return formcontent;
  }

  /** 获取页面表单中 输入框 */
  function getPageFormTextarea(fields = []) {
    // 获取表单
    const formcontent = getPageForm();
    if (!formcontent) {
      return null;
    }

    // 获取输入框
    const textarea = queryFormTextarea(formcontent, fields);
    if (!textarea) {
      log('【检查页面】', `未找到输入框： ${fields.join('|')}`);
      return null;
    }

    return textarea;
  }

  /** 获取 iFrame 表单 */
  function queryIFrameFormContent(document) {
    if (!document) return;
    const formIframe = document.querySelectorAll('#iFrameResizer0');
    if (!formIframe || !formIframe.length) return;
    const iframeDocument = formIframe[0].contentDocument || formIframe[0].contentWindow.document;
    if (!iframeDocument) return;
    const formcontent = iframeDocument.querySelector('#formcontent');
    return formcontent;
  }

  /** 获取获取表单中 输入框 */
  function queryFormTextarea(formContent, fields = []) {
    if (!formContent || !fields.length) return;
    let formTextareas = formContent.querySelectorAll('textarea');
    if (!formTextareas || !formTextareas.length) return;
    let formTextarea = null;
    formTextareas.forEach((item) => {
      const fsref = item.getAttribute('fsref');
      if (!fsref) return;
      const textareas = fields.filter((field) => fsref.includes(field));
      if (textareas.length) {
        formTextarea = item;
        return;
      }
    });
    return formTextarea;
  }

  /** =================================== 通用工具 ============================================ */

  /** 检查是否是管理页面，即非登录、注册等页面 */
  function isManagePage() {
    const pageUrl = window.location.href;
    return !pageUrl.includes('/account/login') && !pageUrl.includes('/account/resetpwd') && !pageUrl.includes('/account/register') && !pageUrl.includes('/account/logout');
  }

  /** 获取导航栏 */
  function getNavBar() {
    const navBar = document.querySelector('#top-global');
    return navBar && navBar.children ? navBar : null;
  }

  /** 导航栏是否存在元素 */
  function hasNavBarItem(elementId) {
    const navBar = getNavBar();
    if (navBar && navBar.children && Array.from(navBar.children).some((element) => element.id === elementId)) {
      return { id: elementId, exists: true, navBar };
    } else {
      return { id: elementId, exists: false, navBar: navBar };
    }
  }

  /** 获取表单底部操作栏 */
  function getFormFooterBar() {
    const footer = document.querySelector('#workflow-footer');
    if (!footer) return null;
    const formFooter = footer.querySelector('.form-footer');
    if (!formFooter) return null;
    return formFooter.querySelector('.content');
  }

  /** 获取配置 */
  function getConfig(key) {
    const config = GM_getValue('gdytw') || {};
    if (key && key.length) return config[key];
    return config;
  }

  /** 设置配置 */
  function setConfig(key, value) {
    const config = GM_getValue('gdytw') || {};
    if (key && key.length) {
      config[key] = value;
    }
    GM_setValue('gdytw', config);
  }

  /** 获取授权 */
  function getAuthorization() {
    return new Promise((resolve) => {
      authorization = JSON.parse(sessionStorage.getItem('UniWork.user:http://oa.gdytw.net/identity:appjs') || '{}').access_token || '';
      authorization.length ? resolve(authorization) : resolve('');
    });
  }

  /** 获取当前页面 URL */
  function getCurrentUrl() {
    const url = window.location.href;
    if (!url || !url.length) {
      log('【获取当前页面 URL】', url);
      return url;
    }

    const Q_PARAM = '?q=';
    const qIndex = url.indexOf(Q_PARAM);
    if (qIndex === -1) {
      log('【获取当前页面 URL】', url);
      return url;
    }

    const key = getURLPassword();
    if (!key || !key.length) {
      log('【获取当前页面 URL】', url);
      return url;
    }

    try {
      // 提取base64编码的加密数据（q=后面的部分）
      const encryptedBase64 = url.substring(qIndex + 3);
      const decodedData = atob(encryptedBase64);

      const ivHex = decodedData.substring(0, 32);
      const cipherHex = decodedData.substring(32);

      if (!CryptoJS) {
        throw new Error('CryptoJS not available. Please include crypto-js library.');
      }

      // 解析IV和密钥
      const iv = CryptoJS.enc.Hex.parse(ivHex);
      const parsedKey = CryptoJS.enc.Utf8.parse(key);

      // AES解密
      const decrypted = CryptoJS.AES.decrypt(cipherHex, parsedKey, { iv: iv, padding: CryptoJS.pad.Pkcs7 });
      const decryptedUrl = decrypted.toString(CryptoJS.enc.Utf8);

      if (!decryptedUrl) return url;

      // 如果是完整URL，保留q参数之前的部分
      if (url.startsWith('http')) {
        const fullUrl = url.substring(0, qIndex).trim().replace(/\/$/, '') + '/' + decryptedUrl.trim().replace(/^\/+/, '');
        log('【获取当前页面 URL】', fullUrl);
        return fullUrl;
      }

      return decryptedUrl;
    } catch (error) {
      log('【获取当前页面 URL】', `解密失败： ${error.message}, 原始URL: ${url}`);
      return url;
    }
  }

  /** 获取 URL 解析密码 */
  function getURLPassword() {
    const cookies = getCookies() || {};
    const vid = cookies['.vid'] || '';

    let password = '';
    if (vid.length) {
      password = atob(decodeURIComponent(vid));
      password = CryptoJS.enc.Hex.parse(password).toString(CryptoJS.enc.Utf8);
    }

    return password;
  }

  /** 获取所有 Cookie */
  function getCookies() {
    const cookies = {};
    document.cookie.split(';').forEach((item) => {
      const [key, value] = item.trim().split('=');
      if (key) cookies[key] = decodeURIComponent(value);
    });
    return cookies;
  }

  /** 拦截请求 */
  function hookRequest({ url, fun }) {
    if (!url || !url.length || !fun || Object.prototype.toString.call(fun) !== '[object Function]') {
      return;
    }
    const originOpen = XMLHttpRequest.prototype.open;
    XMLHttpRequest.prototype.open = function (_, u) {
      if (url === u) {
        this.addEventListener('readystatechange', function () {
          if (this.readyState === 4) {
            Object.defineProperty(this, 'response', { writable: true });
            this.response = JSON.stringify(fun(JSON.parse(this.responseText)));
          }
        });
      }
      originOpen.apply(this, arguments);
    };
  }

  /** 主动请求 */
  function request({ url, data, method = 'POST', headers }) {
    return new Promise(async (resolve, reject) => {
      if (method !== 'POST' && method !== 'GET') {
        return reject('请求方法错误');
      }

      // 获取授权
      if (!authorization || !authorization.length) {
        const auth = await getAuthorization();
        if (!auth || !auth.length) {
          toast('获取授权失败');
          reject('获取授权失败');
        }
      }

      // 发起请求
      const xhr = new XMLHttpRequest();
      if (url && !url.startsWith('http')) {
        url = baseUrl + url;
      }

      if (method === 'POST') {
        // POST
        xhr.open('POST', url, true);
        // header
        if (!headers) {
          headers = { 'Content-Type': 'application/json', Authorization: 'Bearer ' + authorization };
        }
        for (const key in headers) {
          xhr.setRequestHeader(key, headers[key]);
        }
        xhr.onreadystatechange = function () {
          if (xhr.readyState === 4 && xhr.status === 200) {
            const res = JSON.parse(xhr.responseText);
            if (res.Status !== 200) return reject(xhr.responseText);
            resolve(res);
          } else if (xhr.readyState === 4) {
            reject(xhr.responseText);
          }
        };
        xhr.send(JSON.stringify(data || {}));
      } else {
        // GET
        data = new URLSearchParams({ ...data, access_token: authorization });
        xhr.open('GET', url + '?' + data, true);
        xhr.onreadystatechange = function () {
          if (xhr.readyState === 4 && xhr.status === 200) {
            resolve(xhr.responseText);
          } else if (xhr.readyState === 4) {
            reject(xhr.responseText);
          }
        };
        xhr.send();
      }
    });
  }

  /** 主动请求(跨域) */
  function requestGM({ url, data, method = 'POST', headers }) {
    return new Promise(async (resolve, reject) => {
      if (method !== 'POST' && method !== 'GET') {
        return reject('请求方法错误');
      }

      // 获取授权
      if (!authorization || !authorization.length) {
        const auth = await getAuthorization();
        if (!auth || !auth.length) {
          toast('获取授权失败');
          reject('获取授权失败');
        }
      }

      // 发起请求
      GM_xmlhttpRequest({
        method,
        url,
        headers,
        data: JSON.stringify(data),
        onload: function (response) {
          if (response.status === 200) {
            resolve(response.responseText);
          } else {
            reject(response.responseText);
          }
        },
        onerror: function (error) {
          reject(error);
        },
      });
    });
  }

  /** OpenAI 对话 */
  function openAIChat(message) {
    return new Promise(async (resolve, reject) => {
      if (!message || !message.length) {
        return reject('消息不能为空');
      }
      try {
        const apiKey = settings.openAIAPIKey;
        if (!apiKey || !apiKey.length) {
          return reject('API 密钥未配置');
        }

        let baseUrl = settings.openAIBaseURL;
        if (!baseUrl || !baseUrl.length || !baseUrl.startsWith('http')) {
          baseUrl = defaultOpenAIBaseURL;
        }
        baseUrl = baseUrl.replace(/\/$/, '');

        let model = settings.openAIModel;
        if (!model || !model.length) {
          model = defaultOpenAIModel;
        }

        let systemPrompt = settings.logSystemPrompt;
        if (!systemPrompt || !systemPrompt.length) {
          systemPrompt = defaultLogSystemPrompt;
        }

        log('【OpenAI对话配置】', { baseUrl, model, apiKey, systemPrompt });

        const chatRes = await requestGM({
          url: `${baseUrl}/v1/chat/completions`,
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Authorization: 'Bearer ' + apiKey,
          },
          data: {
            model,
            messages: [
              { role: 'system', content: systemPrompt },
              { role: 'user', content: message },
            ],
          },
        });

        log('【OpenAI对话响应】', chatRes);
        const chatData = JSON.parse(chatRes);

        if (chatData.choices && chatData.choices.length) {
          const content = chatData.choices[0].message.content || '';
          resolve(content);
        } else {
          reject('对话返回结果异常');
        }
      } catch (error) {
        log('【OpenAI发起对话失败】', error);
        reject(error);
      }
    });
  }

  /** Toast */
  function toast(msg) {
    const toastContainer = document.createElement('div');
    toastContainer.textContent = msg;
    toastContainer.style.position = 'fixed';
    toastContainer.style.bottom = '20px';
    toastContainer.style.left = '50%';
    toastContainer.style.transform = 'translateX(-50%)';
    toastContainer.style.backgroundColor = 'rgba(0,0,0,0.7)';
    toastContainer.style.color = 'white';
    toastContainer.style.padding = '10px 20px';
    toastContainer.style.borderRadius = '5px';
    toastContainer.style.zIndex = '1000';
    document.body.appendChild(toastContainer);

    setTimeout(() => {
      document.body.removeChild(toastContainer);
    }, 3000); // 消息将在3秒后消失
  }

  /** 下载 markdown 文件 */
  function downloadMarkdown(markdownContent) {
    // 创建 Blob 对象
    const blob = new Blob([markdownContent], {
      type: 'text/markdown;charset=utf-8',
    });

    let year = settings.weekDailyLogYear || '';
    if (!year.length || !isYearValid(year)) {
      year = new Date().getFullYear().toString();
    }

    // 创建下载链接
    const downloadLink = document.createElement('a');
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = `${userName}_周报_${year}_${new Date().toLocaleDateString()}.md`;

    // 触发下载
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);

    // 释放 URL 对象
    URL.revokeObjectURL(downloadLink.href);
  }

  /** 日志 */
  function log(...args) {
    if (settings.debug) {
      console.log(...args);
    }
  }
})();
