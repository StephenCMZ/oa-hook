// ==UserScript==
// @name         OA 系统
// @namespace    https://github.com/StephenCMZ/oa-hook.git
// @version      0.2
// @description  OA 系统
// @author       StephenChen
// @match        http://oa.gdytw.net/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_xmlhttpRequest
// @downloadURL  https://github.com/StephenCMZ/oa-hook/blob/main/oa.js
// @updateURL    https://github.com/StephenCMZ/oa-hook/blob/main/oa.js
// ==/UserScript==

(function () {
  'use strict';

  const debug = false;
  const baseUrl = 'http://oa.gdytw.net';
  const getLinkPageUrl = '/api/Portal/Content/LinkGetPage';
  const logListUrl = '/api/Workflow/FlowList/GetRequireList';
  const logDetailUrl = '/api/Workflow/FlowMan/GetDetail';
  const logContentUrl = '/api/Form/ExternalDataSource/GetDataList';
  const workFlowDetailUrl = '/api/Workflow/FlowMan/GetPrint';

  // AI
  const cozeChatUrl = 'https://api.coze.cn/v3/chat';
  const cozeRetrieveUrl = 'https://api.coze.cn/v3/chat/retrieve';
  const cozeMessageUrl = 'https://api.coze.cn/v3/chat/message/list';
  let cozeAccessToken = getConfig('cozeAccessToken') || '';
  const cozeBotId = '7472312758722560039';

  const year = new Date().getFullYear().toString();
  const pageSize = 200;
  let authorization = '';
  let userName = '';

  const loadFormTimes = 1000;

  init();

  /** 初始化 */
  function init() {
    hookShortMenu();
    window.addEventListener('load', function () {
      setTimeout(addSettingBtn, 1000);
      setTimeout(addExportBtn, 1000);
      setTimeout(autoFillFormPlan, loadFormTimes);
      setTimeout(autoFillFormWeekLog, loadFormTimes);
    });
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
    // 是否为编辑模式
    const operatorBar = document.querySelector('#start-operator');
    if (!operatorBar || !operatorBar.outerHTML.includes('提交申请')) {
      log('【检查页面】', '非表单编辑页面');
      return;
    }

    // 获取表单
    const formcontent = queryIFrameFormContent(document);
    if (!formcontent) {
      log('【检查页面】', '未找到表单');
      return;
    }

    // 获取计划输入框
    const planTextarea = queryFormTextarea(formcontent, ['明日工作计划', '下周工作计划']);
    if (!planTextarea) {
      log('【检查页面】', '未找到计划输入框');
      return;
    }

    // 填入计划
    fillFormPlan(planTextarea);
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

  /** =================================== 自动填充本周工作总结 ============================================ */

  /** 自动填充表单本周工作总结 */
  function autoFillFormWeekLog() {
    // 是否为编辑模式
    const operatorBar = document.querySelector('#start-operator');
    if (!operatorBar || !operatorBar.outerHTML.includes('提交申请')) {
      log('【检查页面】', '非表单编辑页面');
      return;
    }

    // 获取表单
    const formcontent = queryIFrameFormContent(document);
    if (!formcontent) {
      log('【检查页面】', '未找到表单');
      return;
    }

    // 获取本周工作总结输入框
    const weekLogTextarea = queryFormTextarea(formcontent, ['本周工作总结']);
    if (!weekLogTextarea) {
      log('【检查页面】', '未找到本周工作总结输入框');
      return;
    }

    fillFormWeekLog(weekLogTextarea);
  }

  /** 填充表单本周工作总结内容 */
  async function fillFormWeekLog(weekLogTextarea) {
    if (!weekLogTextarea) return;

    try {
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

      // AI 整理内容
      if (cozeAccessToken && cozeAccessToken.length) {
        try {
          const aiLogDetails = await cozeChat(weekLogs);
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
    }
  }

  function getWeekDailyLogList() {
    let startDate = getConfig('weekDailyLogStartDate') || '';
    let endDate = getConfig('weekDailyLogEndDate') || '';
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

  /** =================================== 导出全年周志 ============================================ */

  /** 导航栏添加 下载周志 按钮 */
  function addExportBtn() {
    var navBar = document.querySelector('#top-global');
    if (!navBar || !navBar.children) return;
    if (Array.from(navBar.children).some((element) => element.id === 'export')) return;

    var liItem = document.createElement('li');
    liItem.id = 'export';
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
  }

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
    var navBar = document.querySelector('#top-global');
    if (!navBar || !navBar.children) return;
    if (Array.from(navBar.children).some((element) => element.id === 'setting')) return;

    var liItem = document.createElement('li');
    liItem.id = 'setting';
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

    settingButton.onclick = settings;

    liItem.appendChild(settingButton);
    navBar.insertBefore(liItem, navBar.firstChild);
  }

  /** 设置弹窗 */
  function settings() {
    // 创建弹窗容器
    const dialog = document.createElement('div');
    dialog.style.position = 'fixed';
    dialog.style.top = '50%';
    dialog.style.left = '50%';
    dialog.style.transform = 'translate(-50%, -50%)';
    dialog.style.backgroundColor = 'white';
    dialog.style.padding = '20px';
    dialog.style.borderRadius = '4px';
    dialog.style.boxShadow = '0 2px 8px rgba(0, 0, 0, 0.15)';
    dialog.style.zIndex = '9999';

    // 创建弹窗标题
    const title = document.createElement('h2');
    title.textContent = '设置';
    title.style.marginTop = '0';
    dialog.appendChild(title);

    // 创建输入框容器
    const inputContainer = document.createElement('div');
    inputContainer.style.display = 'flex';
    inputContainer.style.flexDirection = 'column';
    inputContainer.style.gap = '15px';
    inputContainer.style.marginBottom = '20px';

    // 创建 AI 密钥输入框
    const aiKeyInput = document.createElement('input');
    aiKeyInput.type = 'text';
    aiKeyInput.placeholder = '请输入 AI 密钥';
    aiKeyInput.value = cozeAccessToken;
    aiKeyInput.style.width = '400px';
    aiKeyInput.style.padding = '8px';
    aiKeyInput.style.border = '1px solid #d9d9d9';
    aiKeyInput.style.borderRadius = '4px';
    inputContainer.appendChild(aiKeyInput);

    // 创建周志开始时间输入框
    const weekDailyLogStartDateInput = document.createElement('input');
    weekDailyLogStartDateInput.type = 'text';
    weekDailyLogStartDateInput.placeholder = '周志开始时间格式为 YYYY-MM-DD, 不填默认本周一';
    weekDailyLogStartDateInput.value = getConfig('weekDailyLogStartDate') || '';
    weekDailyLogStartDateInput.style.width = '400px';
    weekDailyLogStartDateInput.style.padding = '8px';
    weekDailyLogStartDateInput.style.border = '1px solid #d9d9d9';
    weekDailyLogStartDateInput.style.borderRadius = '4px';
    inputContainer.appendChild(weekDailyLogStartDateInput);

    // 创建周志结束时间输入框
    const weekDailyLogEndDateInput = document.createElement('input');
    weekDailyLogEndDateInput.type = 'text';
    weekDailyLogEndDateInput.placeholder = '周志结束时间格式为 YYYY-MM-DD, 不填默认本周日';
    weekDailyLogEndDateInput.value = getConfig('weekDailyLogEndDate') || '';
    weekDailyLogEndDateInput.style.width = '400px';
    weekDailyLogEndDateInput.style.padding = '8px';
    weekDailyLogEndDateInput.style.border = '1px solid #d9d9d9';
    weekDailyLogEndDateInput.style.borderRadius = '4px';
    inputContainer.appendChild(weekDailyLogEndDateInput);

    dialog.appendChild(inputContainer);

    // 创建按钮容器
    const btnContainer = document.createElement('div');
    btnContainer.style.textAlign = 'right';

    // 创建取消按钮
    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = '取消';
    cancelBtn.style.marginRight = '8px';
    cancelBtn.style.padding = '4px 15px';
    cancelBtn.style.backgroundColor = '#f0f0f0';
    cancelBtn.style.border = 'none';
    cancelBtn.style.borderRadius = '4px';
    cancelBtn.style.cursor = 'pointer';
    cancelBtn.onclick = () => {
      document.body.removeChild(dialog);
    };

    // 创建确认按钮
    const confirmBtn = document.createElement('button');
    confirmBtn.textContent = '确认';
    confirmBtn.style.padding = '4px 15px';
    confirmBtn.style.backgroundColor = '#1890ff';
    confirmBtn.style.color = 'white';
    confirmBtn.style.border = 'none';
    confirmBtn.style.borderRadius = '4px';
    confirmBtn.style.cursor = 'pointer';
    confirmBtn.onclick = () => {
      // 保存 AI 密钥配置
      cozeAccessToken = aiKeyInput.value;
      setConfig('cozeAccessToken', aiKeyInput.value);

      // 保存周志开始时间
      const weekDailyLogStartDate = weekDailyLogStartDateInput.value || '';
      if (weekDailyLogStartDate.length && !isDateValid(weekDailyLogStartDate)) {
        toast('周志开始时间格式异常');
        return;
      }
      setConfig('weekDailyLogStartDate', weekDailyLogStartDate);

      // 保存周志结束时间
      const weekDailyLogEndDate = weekDailyLogEndDateInput.value || '';
      if (weekDailyLogEndDate.length && !isDateValid(weekDailyLogEndDate)) {
        toast('周志结束时间格式异常');
        return;
      }
      setConfig('weekDailyLogEndDate', weekDailyLogEndDate);

      // 关闭弹窗
      document.body.removeChild(dialog);
      toast('保存成功');
    };

    btnContainer.appendChild(cancelBtn);
    btnContainer.appendChild(confirmBtn);
    dialog.appendChild(btnContainer);

    document.body.appendChild(dialog);
  }

  /** =================================== 通用工具 ============================================ */

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
      authorization =
        JSON.parse(sessionStorage.getItem('UniWork.user:http://oa.gdytw.net/identity:appjs') || '{}').access_token ||
        '';
      authorization.length ? resolve(authorization) : resolve('');
    });
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

  /** 扣子对话 */
  function cozeChat(message) {
    return new Promise(async (resolve, reject) => {
      if (!message || !message.length) {
        return reject('消息不能为空');
      }

      try {
        // 获取 OA 授权, 作为 cozeUserId
        if (!authorization || !authorization.length) {
          const auth = await getAuthorization();
          if (!auth || !auth.length) {
            toast('获取授权失败');
            reject('获取授权失败');
          }
        }

        // 发起对话
        const enterMessage = {
          role: 'user',
          content: message,
          content_type: 'text',
        };
        let chatRes = await requestGM({
          url: cozeChatUrl,
          headers: {
            'Content-Type': 'application/json',
            Authorization: 'Bearer ' + cozeAccessToken,
          },
          data: {
            bot_id: cozeBotId,
            user_id: authorization,
            additional_messages: [enterMessage],
          },
        });
        log('【AI发起对话】', chatRes);
        chatRes = JSON.parse(chatRes);

        // 请求失败
        if (chatRes.code !== 0) {
          return reject(chatRes);
        }

        // 查询结果
        const chatResponse = await cozeChatResponse(chatRes.data);
        if (!chatResponse.data || !chatResponse.data.length) {
          return reject('对话失败');
        }

        // 解析结果
        const chatContent = chatResponse.data.find((item) => item.role === 'assistant' && item.type === 'answer');
        let content = (chatContent.content || '').replaceAll('### ', '');

        resolve(content);
      } catch (error) {
        log('【AI发起对话失败】', error);
        reject(error);
      }
    });
  }

  /** 扣子对话结果查询 */
  function cozeChatResponse(chat) {
    return new Promise(async (resolve, reject) => {
      if (!chat || !chat.id || !chat.conversation_id) {
        return reject('对话不能为空');
      }
      const { id, conversation_id } = chat;

      const poll = () => {
        return requestGM({
          url: `${cozeRetrieveUrl}?conversation_id=${conversation_id}&chat_id=${id}`,
          method: 'GET',
          headers: {
            'Content-Type': 'application/json',
            Authorization: 'Bearer ' + cozeAccessToken,
          },
        });
      };

      try {
        // 轮询对话状态
        let chatRes = await poll();
        log('【AI轮询对话状态】', chatRes);
        chatRes = JSON.parse(chatRes);
        while (chatRes.data.status !== 'completed') {
          await new Promise((resolve) => setTimeout(resolve, 1000)); // 每1秒轮询一次
          chatRes = await poll();
          log('【AI轮询对话状态】', chatRes);
          chatRes = JSON.parse(chatRes);
        }

        // 查询结果
        let chatMessageRes = await requestGM({
          url: `${cozeMessageUrl}?conversation_id=${conversation_id}&chat_id=${id}`,
          method: 'GET',
          headers: {
            'Content-Type': 'application/json',
            Authorization: 'Bearer ' + cozeAccessToken,
          },
        });
        log('【AI查询对话结果】', chatMessageRes);
        chatMessageRes = JSON.parse(chatMessageRes);
        if (chatMessageRes.code !== 0) {
          return reject(chatMessageRes);
        }

        resolve(chatMessageRes);
      } catch (error) {
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

    // 创建下载链接
    const downloadLink = document.createElement('a');
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = `${userName}_周报_${new Date().toLocaleDateString()}.md`;

    // 触发下载
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);

    // 释放 URL 对象
    URL.revokeObjectURL(downloadLink.href);
  }

  /** 日志 */
  function log(...args) {
    if (debug) {
      console.log(...args);
    }
  }
})();
