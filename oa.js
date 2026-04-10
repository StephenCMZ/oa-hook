// ==UserScript==
// @name         OA 系统
// @namespace    https://github.com/StephenCMZ/oa-hook.git
// @version      0.8
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

  const debug = false;
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

  // 设置
  const defaultSettings = {
    cozeAccessToken: '', // AI 密钥
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
  };
  let settings = { ...defaultSettings, ...getConfig('settings') };

  // 表单模板 ID
  const dailyTemplateId = '592233945022595072';
  const weekTemplateId = '592231167478988800';

  // 组件 ID
  const nav_setting_btn_id = 'setting_btn';
  const nav_export_btn_id = 'export_btn';
  const nav_statistics_info_id = 'statistics_info';

  // 统计信息
  let statistics = {};

  // AI
  const cozeChatUrl = 'https://api.coze.cn/v3/chat';
  const cozeRetrieveUrl = 'https://api.coze.cn/v3/chat/retrieve';
  const cozeMessageUrl = 'https://api.coze.cn/v3/chat/message/list';
  const cozeBotId = '7472312758722560039';

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
      guardAddElement(addStatisticsInfo); // 添加导航栏统计信息
      guardFillEditForm(autoFillFormPlan, [dailyTemplateId, weekTemplateId]); // 自动填充明日/下周工作计划
      guardFillEditForm(autoFillFormDailyLog, [dailyTemplateId]); // 自动填充日报记录
      guardFillEditForm(autoFillFormWeekLog, [weekTemplateId]); // 自动填充周报记录
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

      // AI 整理内容
      const cozeAccessToken = settings.cozeAccessToken;
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

  /** =================================== 自动选择日报/周报抄送人和点评人 ============================================ */

  /** 自动选择日报/周报抄送人和点评人 */
  function autoSelectReviewer() {
    if (!settings.autoSelectReviewer) return;
    hookRequest({
      url: workFlowGetPreSelUsersUrl,
      fun: function (res) {
        const users = res.Data.ForPreSelUsers;
        if (users && users.length === 2) {
          const toNodeName0 = (users[0] || {}).ToNodeName; // 抄送
          const toNodeName1 = (users[1] || {}).ToNodeName; // 点评人

          if (toNodeName0 === '抄送' && toNodeName1 === '点评人') {
            // 抄送，默认选择关信东
            const copyer = `{"Id":"589376478202929152","Name":"关信东","Code":"7007","DepId":"325475396013277184","DepName":"综合事务部","DepFullName":"综合事务部","DepPath":"0325475396013277184","SortCode":"9999","AllowLogin":true,"Actived":true,"Pinyin":"GuanXinDong","InitialPinyin":"GXD","type":0}`;
            users[0] = {
              ...users[0],
              SelUsers: [JSON.parse(copyer)],
            };
            // 点评人，默认选择孔文威
            const reviewer = `{"Name":"孔文威","Account":"2011","Code":"2011","Gender":true,"ContactVisibility":true,"Mobile":"18028196559","Cornet":null,"Telephone":null,"Email":"422963845@qq.com","WeChat":null,"SuperiorId":"0","Actived":true,"AllowLogin":true,"InitialPinyin":"KWW","Pinyin":"KongWenWei","IsOnline":false,"IsExternal":false,"ExtensionNumber":null,"AllowMobile":true,"InitialWubi":"BYD","AccountValidity":null,"Pin":null,"Creator":null,"CreatorId":"0","CreateTime":null,"UpdateUserName":"龙加鎏","UpdateUserId":"583962765559959552","UpdateTime":"2025-12-30T10:46:08.5518513","InactiveTime":null,"Department":{"Name":"软件信息部","Path":"0324796940434804736","FullName":"软件信息部","Actived":true,"SortCode":"0005","Categories":"","WeChatWorkDepId":59,"WeChatWorkCorpId":"ww41fa3fdd30318beb","Manning":0,"NoLimit":false,"ContactVisibility":true,"UserCount":0,"Id":"324796940434804736"},"Positions":[{"UserId":"583962756424765440","PositionId":"668027858240782336","Major":true,"Sequence":1,"SortCode":"999999","Position":{"JobId":"645577122982703104","OrganizeId":"324796940434804736","Name":"软件信息部/技术总监","Job":{"Name":"技术总监","SysBuildIn":false,"SortCode":999999,"Id":"645577122982703104"},"Id":"668027858240782336"},"Id":"583962757246849024"}],"Roles":[],"OrganizeRoles":[],"Password":null,"Superior":null,"RelationOrganizes":null,"UserCardExtra":null,"Id":"583962756424765440","Job":{"Name":"技术总监","SysBuildIn":false,"SortCode":999999,"Id":"645577122982703104"},"UserId":"583962756424765440","FormerName":null,"PostId":"5","Post":"技术岗","PostIdPath":["5"],"TitleId":null,"Title":null,"TitleConferingDate":null,"TitleGradeId":null,"TitleGrade":null,"Status":"在职","EmployeeCategory":null,"Birthday":"1986-02-18T00:00:00","IdentityCardNumber":"440682198602181036","PassportNumber":null,"Nationality":"中国","EthnicGroup":"汉族","NativePlace":"佛山","CurrentResidence":"广州市白云区京溪路云景花园新云桂苑16栋","RegistedResidence":null,"MaritalStatus":"未婚","PoliticalStatus":"群众","HighestDegree":"研究生","StartingDateOfFirstJob":"2009-03-01T00:00:00","EnterDate":"2017-02-06T00:00:00","ResignationDate":null,"RetirementDate":null,"Insurance":null,"InsuredDate":null,"AttendanceMachineId":null,"PostCategory":null,"BirthAddress":null,"JoinPartyDate":null,"ArchivesNumber":null,"StaffingCategory":null,"AccountUnitId":null,"AccountUnitName":null,"PersonnelModel":null,"OrgJobMid":null,"AvatarFileId":null,"IsLeader":null,"PersonnelContract":null,"DepartmentRecord":null,"ExtraFile":null,"EmployeeId":"583962756487680000","type":0}`;
            users[1] = {
              ...users[1],
              SelUsers: [JSON.parse(reviewer)],
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

  /** 设置弹窗 */
  function showSettings() {
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
    dialog.style.minWidth = '500px';
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

    // 创建设置项
    const settingItems = [
      { key: 'cozeAccessToken', type: 'text', placeholder: '请输入 AI 密钥' },
      { key: 'weekDailyLogStartDate', type: 'text', placeholder: '自动填充周志开始时间格式为 YYYY-MM-DD, 不填默认本周一' },
      { key: 'weekDailyLogEndDate', type: 'text', placeholder: '自动填充周志结束时间格式为 YYYY-MM-DD, 不填默认本周日' },
      { key: 'weekDailyLogYear', type: 'text', placeholder: '下载周志年份格式为 YYYY, 不填默认当前年份' },
      { key: 'autoFillDailyLog', type: 'checkbox', labelText: '自动填充日志' },
      { key: 'autoFillWeeklyLog', type: 'checkbox', labelText: '自动填充周志' },
      { key: 'aiFillWeeklyLog', type: 'checkbox', labelText: 'AI 整理周志' },
      { key: 'autoFillPlan', type: 'checkbox', labelText: '自动填充计划' },
      { key: 'autoSelectReviewer', type: 'checkbox', labelText: '自动选点评人' },
      { key: 'showDownloadWeekDailyLogBtn', type: 'checkbox', labelText: '显示下载周志' },
      { key: 'showStatisticsInfo', type: 'checkbox', labelText: '显示统计信息' },
      { key: 'showHitokoto', type: 'checkbox', labelText: '显示每日一言' },
    ];
    settingItems.forEach((item) => {
      if (item.type === 'text') {
        const { inputElement, input } = createInputElement({ placeholder: item.placeholder, value: settings[item.key] });
        inputContainer.appendChild(inputElement);
        item.formItem = input;
      } else if (item.type === 'checkbox') {
        const { checkboxElement, checkbox } = createCheckboxElement({ labelText: item.labelText, value: settings[item.key] });
        inputContainer.appendChild(checkboxElement);
        item.formItem = checkbox;
      }
    });
    dialog.appendChild(inputContainer);

    // 创建按钮容器
    const btnContainer = document.createElement('div');
    btnContainer.style.textAlign = 'right';

    // 创建取消按钮
    const cancelBtn = createButtonElement({
      title: '取消',
      onClick: () => document.body.removeChild(dialog),
    });

    // 创建确认按钮
    const confirmBtn = createButtonElement({
      title: '确认',
      type: 'primary',
      onClick: () => {
        // 校验周志开始时间
        const weekDailyLogStartDate = settingItems.find((item) => item.key === 'weekDailyLogStartDate')?.formItem?.value || '';
        if (weekDailyLogStartDate.length && !isDateValid(weekDailyLogStartDate)) {
          toast('周志开始时间格式异常');
          return;
        }
        // 校验周志结束时间
        const weekDailyLogEndDate = settingItems.find((item) => item.key === 'weekDailyLogEndDate')?.formItem?.value || '';
        if (weekDailyLogEndDate.length && !isDateValid(weekDailyLogEndDate)) {
          toast('周志结束时间格式异常');
          return;
        }
        // 校验下载周志年份
        const weekDailyLogYear = settingItems.find((item) => item.key === 'weekDailyLogYear')?.formItem?.value || '';
        if (weekDailyLogYear.length && !isYearValid(weekDailyLogYear)) {
          toast('周志年份格式异常');
          return;
        }

        // 获取所有设置项的值
        const _settings = { ...settings };
        settingItems.forEach((item) => {
          if (item.type === 'text') {
            _settings[item.key] = item.formItem.value;
          } else if (item.type === 'checkbox') {
            _settings[item.key] = item.formItem.checked;
          }
        });

        // 保存设置
        settings = _settings;
        setConfig('settings', _settings);
        console.log('保存设置:', _settings);

        // 关闭弹窗
        document.body.removeChild(dialog);
        toast('保存成功, 请刷新页面生效');
      },
    });

    btnContainer.appendChild(cancelBtn);
    btnContainer.appendChild(confirmBtn);
    dialog.appendChild(btnContainer);

    document.body.appendChild(dialog);
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

  /** 创建复选框 */
  function createCheckboxElement({ labelText, value }) {
    const checkboxElement = document.createElement('div');
    checkboxElement.style.display = 'flex';
    checkboxElement.style.alignItems = 'center';

    if (labelText && labelText.trim() !== '') {
      const label = document.createElement('label');
      label.textContent = labelText;
      label.style.fontSize = '14px';
      label.style.marginRight = '8px';
      checkboxElement.appendChild(label);
    }

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = value || false;
    checkboxElement.appendChild(checkbox);

    return { checkboxElement, checkbox };
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
        const userVacations = ((userVacationRes || {}).Data || {}).Data || [];
        if (userVacations.length) {
          statistics.vacations = formVacations(userVacations[0] || {});
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
  function formVacations(userVacation = {}) {
    if (!userVacation || !Object.keys(userVacation).length) {
      return {};
    }
    const vacations = [];

    const annual = formVacationByKey(userVacation, '1');
    vacations.push({ key: 'annual', name: '剩余年假', value: annual.total - annual.used });
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

  /** =================================== 表单工具 ============================================ */

  /** 检查页面是否为表单编辑页面 */
  function isFormPage(templateId) {
    const currentUrl = getCurrentUrl();
    if (!currentUrl || !currentUrl.length) return false;

    if (templateId && templateId.length) {
      return currentUrl.includes(`templateId=${templateId}`);
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
    return pageUrl.includes('?q=') || pageUrl.includes('/auth-callback');
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
    if (!url || !url.length) return url;

    const Q_PARAM = '?q=';
    const qIndex = url.indexOf(Q_PARAM);
    if (qIndex === -1) return url;

    const key = getURLPassword();
    if (!key || !key.length) return url;

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
            Authorization: 'Bearer ' + settings.cozeAccessToken,
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
            Authorization: 'Bearer ' + settings.cozeAccessToken,
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
            Authorization: 'Bearer ' + settings.cozeAccessToken,
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
    if (debug) {
      console.log(...args);
    }
  }
})();
