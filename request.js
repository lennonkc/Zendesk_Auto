const axios = require('axios');
const moment = require('moment');
const ExcelJS = require('exceljs');
const readline = require('readline');
const getUserNameById = require('./reports/Users/getUserNameByID');

require('dotenv').config(); // 加载环境变量
let lastRunTime = null;

const ticketRequester = {

    defaultTimeZone: 'America/New_York',
  /**
   * Converts a date string in "YYYY-MM-DD" format to a UNIX timestamp.
   * @param {string} dateString - The input date string in "YYYY-MM-DD" format.
   * @param {string} [timeZone="America/New_York"] - The time zone to use for conversion (default is "America/New_York").
   * @returns {number} - The UNIX timestamp (in seconds) corresponding to the input date.
   */
  convertDateToTimestamp(dateString, timeZone = this.defaultTimeZone) {
    if (!moment(dateString, 'YYYY-MM-DD', true).isValid()) {
      throw new RangeError(`Invalid date format: "${dateString}". Expected format: "YYYY-MM-DD".`);
    }
  
    const date = new Date(dateString + "T00:00:00");
    if (isNaN(date.getTime())) {
      throw new RangeError(`Invalid date value: "${dateString}". Unable to convert to timestamp.`);
    }
  
    const formatter = new Intl.DateTimeFormat('en-US', {
      timeZone,
      year: 'numeric',
      month: 'numeric',
      day: 'numeric',
      hour: 'numeric',
      minute: 'numeric',
      second: 'numeric',
    });
  
    const parts = formatter.formatToParts(date);
    const dateInTimeZone = new Date(
      `${parts.find(p => p.type === 'year').value}-${parts.find(p => p.type === 'month').value.padStart(2, '0')}-${parts.find(p => p.type === 'day').value.padStart(2, '0')}T${parts.find(p => p.type === 'hour').value.padStart(2, '0')}:${parts.find(p => p.type === 'minute').value.padStart(2, '0')}:${parts.find(p => p.type === 'second').value.padStart(2, '0')}`
    );
  
    return Math.floor(dateInTimeZone.getTime() / 1000);
  },


  /**
   * Sends a request to the Zendesk API to retrieve tickets.
   * @param {string} startDate - The start date in "YYYY-MM-DD" format.
   */
  requestToZendesk(startDate) {
    const startTimestamp = this.convertDateToTimestamp(startDate);
    const config = {
      method: 'get',
      maxBodyLength: Infinity,
      url: `https://${process.env.ZENDESK_API_BaseURL}api/v2/incremental/tickets?start_time=${startTimestamp}`,
      headers: { 
        'Accept': 'application/json', 
        'Authorization': `Basic ${process.env.ZENDESK_API_TOKEN}`
      }
    };
    return axios.request(config)
      .then((response) => {
        JSON.stringify(response.data)
        console.log("successfully request");
        return response.data; // 返回数据
      })
      .catch((error) => {
        console.error(error);
        throw new Error("Failed to fetch data from Zendesk API");
      });
  },

};

const usersReuqester = {
    requestToZendeskUpdateAssigneeIDtoName(assignee_id){
        const config = {
            method: 'get',
            maxBodyLength: Infinity,
            url: `https://${process.env.ZENDESK_API_BaseURL}api/v2/users/${assignee_id}`,
            headers: { 
              'Accept': 'application/json', 
              'Authorization': `Basic ${process.env.ZENDESK_API_TOKEN}`
            }
          };
          return axios(config)
          .then(function (response) {
            // console.log(JSON.stringify(response.data));
            console.log(response.data.user.name)
            return response.data.user.name;
          })
          .catch(function (error) {
            console.log(error);
            throw new Error("Failed to query Assignee`s name from Zendesk API");
          });
      },
};

const zendeskAnalyzer = {
    async generateZendeskReport(startDate) {
      try {
        const response = await ticketRequester.requestToZendesk(startDate);
  
        if (!response || !response.tickets || response.tickets.length === 0) {
          console.log('No tickets found in the response.');
          return;
        }
  
        const { tickets, count, end_time } = response;
        lastRunTime = end_time;
        // console.log("updated lastRunTime",lastRunTime)
  
        const now = moment();
        const filteredTickets = tickets.filter(ticket => {
          const lastUpdated = moment(ticket.updated_at);
          const unrespondedHours = now.diff(lastUpdated, 'hours');
  
          if (ticket.status === 'new' || ticket.status === 'open') {
            return unrespondedHours > 24;
          }
  
          if (ticket.status === 'pending') {
            return unrespondedHours > 72;
          }
  
          return false;
        });
  
        // console.log(`Filtered Tickets Count: ${filteredTickets.length}`);
  
        const generateTime = moment.unix(end_time).format('YYYY-MM-DD HH:mm:ss');
        const reportHeader = [
          `Zendesk Report: Analyzed ${count} Tickets, Filtered Tickets: ${filteredTickets.length}`,
          `From: ${startDate} To: ${moment.unix(end_time).format('YYYY-MM-DD HH:mm:ss')}`,
          `Generate Time: ${generateTime}`
        ];
        console.log(reportHeader.join('\n'));
        
        const tableData = await Promise.all(
            filteredTickets.map(async (ticket) => ({
              ticketNumber: ticket.id,
              ticketTitle: ticket.subject,
              ticketContents: ticket.description.length > 100
                ? `${ticket.description.slice(0, 100)}...`
                : ticket.description,
              ticketStatus: ticket.status,
              unrespondedTime: `${moment().diff(moment(ticket.updated_at), 'hours')} hours`,
              assignee: (ticket.assignee_id && await getUserNameById(ticket.assignee_id)) || 'Unassigned',
            }))
          );
        console.table(tableData);

      // 为 Excel 文件准备数据（完整内容）
      const fileName = `./reports/${moment().format('YYYY-MM-DD HH:mm:ss')}.xlsx`;

      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Zendesk Report');

      // 添加列标题
      sheet.columns = [
        { header: 'Ticket Number', key: 'ticketNumber', width: 15 },
        { header: 'Ticket Title', key: 'ticketTitle', width: 30 },
        { header: 'Ticket Contents', key: 'ticketContents', width: 100 },
        { header: 'Ticket Status', key: 'ticketStatus', width: 15 },
        { header: 'Unresponded Time', key: 'unrespondedTime', width: 20 },
        { header: 'Assignee', key: 'assignee', width: 20 },
      ];

      // 添加行数据
      const excelData = await Promise.all(
        filteredTickets.map(async (ticket) => {
          const assigneeName = ticket.assignee_id
            ? await getUserNameById(ticket.assignee_id)
            : 'Unassigned';
      
          // 添加到表格
          sheet.addRow({
            ticketNumber: ticket.id,
            ticketTitle: ticket.subject,
            ticketContents: ticket.description, // 完整内容
            ticketStatus: ticket.status,
            unrespondedTime: `${moment().diff(moment(ticket.updated_at), 'hours')} hours`,
            assignee: assigneeName,
          });
        })
      );


      // 在第一行插入报告头部
      sheet.spliceRows(1, 0, [reportHeader]);
      const headerCell = sheet.getCell('A1');
      headerCell.font = { bold: true, size: 14 };
      headerCell.alignment = { vertical: 'middle', horizontal: 'left' };
      sheet.mergeCells(`A1:F1`);

      // 保存文件
      await workbook.xlsx.writeFile(fileName);
      console.log(`Zendesk Report saved to ${fileName}`);
    } catch (error) {
      console.error('Error generating Zendesk report:', error);
    }
  }
  };


// 提示用户输入日期
function promptForDate() {
    return new Promise((resolve) => {
      const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
      });
  
      rl.question('Please enter the start date (YYYY-MM-DD): ', (inputDate) => {
        rl.close();
        resolve(inputDate);
      });
    });
  }
  
// 检查 lastRunTime 的值并执行报告生成逻辑
async function manageZendeskAnalyzer() {
    if (!lastRunTime) {
        console.log('lastRunTime is empty.');
        const userDate = await promptForDate();

        // 验证输入格式是否正确
        if (!moment(userDate, 'YYYY-MM-DD', true).isValid()) {
        console.log('Invalid date format. Please use "YYYY-MM-DD".');
        return manageZendeskAnalyzer(); // 重新提示用户输入
        }

        lastRunTime = userDate; // 更新 lastRunTime
        console.log(`Setting lastRunTime to ${lastRunTime}`);
        await zendeskAnalyzer.generateZendeskReport(lastRunTime); // 立即运行
    }

    // 转换 lastRunTime 并开始循环执行
    const startDate = moment.unix(lastRunTime).format('YYYY-MM-DD');
    console.log(`lastRunTime is updated: ${moment.unix(lastRunTime).format('YYYY-MM-DD')}. Starting cyclic execution every 6 hours.`);

    async function cyclicExecution() {
        await zendeskAnalyzer.generateZendeskReport(startDate);
        console.log(`Next run scheduled in 6 hours.`);
    }

    // 立即执行一次
    // await cyclicExecution();

    // 每隔 6 小时循环执行
    setInterval(async () => {
        await cyclicExecution();
    }, 6 * 60 * 60 * 1000); // 6 小时
}

// 启动管理函数
manageZendeskAnalyzer();

module.exports = zendeskAnalyzer;
module.exports = ticketRequester;