// ==UserScript==
// @name         OA 系统
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  OA 系统
// @author       StephenChen
// @match        http://oa.gdytw.net/*
// @grant        none
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
      setTimeout(addExportBtn, 1000);
      setTimeout(autoFormPlan, loadFormTimes);
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
  function autoFormPlan() {
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
    const planTextarea = queryFormPlanTextarea(formcontent);
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

  /** 获取获取表单中 明日/下周工作计划 输入框 */
  function queryFormPlanTextarea(formContent) {
    if (!formContent) return;
    let planTextarea = formContent.querySelectorAll('textarea');
    if (!planTextarea || !planTextarea.length) return;
    planTextarea.forEach((item) => {
      const fsref = item.getAttribute('fsref');
      if (!fsref) return;
      if (fsref.includes('明日工作计划') || fsref.includes('下周工作计划')) {
        planTextarea = item;
        return;
      }
    });
    return planTextarea;
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

  /** =================================== 导出全年周志 ============================================ */

  /** 导航栏添加 下载周志 按钮 */
  function addExportBtn() {
    var navBar = document.querySelector('#top-global');
    if (navBar && navBar.firstChild.id !== 'export') {
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

  /** =================================== 通用工具 ============================================ */

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
  function request({ url, data, method = 'POST' }) {
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
      if (method === 'POST') {
        xhr.open('POST', baseUrl + url, true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.setRequestHeader('Authorization', 'Bearer ' + authorization);
        data = JSON.stringify(data || {});
        xhr.onreadystatechange = function () {
          if (xhr.readyState === 4 && xhr.status === 200) {
            const res = JSON.parse(xhr.responseText);
            if (res.Status !== 200) return reject(xhr.responseText);
            resolve(res);
          } else if (xhr.readyState === 4) {
            reject(xhr.responseText);
          }
        };
        xhr.send(data);
      } else {
        data = new URLSearchParams({ ...data, access_token: authorization });
        xhr.open('GET', baseUrl + url + '?' + data, true);
        xhr.onreadystatechange = function () {
          if (xhr.readyState === 4 && xhr.status === 200) {
            resolve(xhr.responseText);
          } else if (xhr.readyState === 4) {
            reject(xhr.responseText);
          }
        };
        xhr.send(data);
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
