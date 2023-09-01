import { Context, Schema, Logger, h } from 'koishi'

// 导入 exceljs 模块，用于操作 Excel 文件
import * as Excel from 'exceljs';

export const name = 'chat-log-excelizer'
export const logger = new Logger('ChatLogExcelizer')
export const usage = `## 🎮 使用

- 建议指令添加指令别名

## 📝 指令说明

本插件提供了以下几个指令：

- \`chatLogExcelizer\`：查看本插件的指令帮助。
- \`chatLogExcelizer.exporterAll\`：导出所有群组的聊天记录（不会自动清空数据表）
- \`chatLogExcelizer.exporter\`：导出当前群组的聊天记录为 Excel 文件，并根据设置决定是否发送到群组或清空数据表。
- \`chatLogExcelizer.clearAllData\`：清空所有群组的聊天记录数据表。
- \`chatLogExcelizer.clearData\`：清空当前群组的聊天记录数据表。`

export interface Config {
  saveDirectory: any
  autoClearDataTableEnabled: boolean
  autoClearAllDataTableEnabled: boolean
  sendFileToGroupEnabled: boolean
}

export const Config: Schema<Config> = Schema.object({
  saveDirectory: Schema.path().default('').description('文件保存目录（可选，默认为 Koishi 项目的根目录）'),
  autoClearDataTableEnabled: Schema.boolean().default(true).description('是否自动清空当前群组的聊天记录数据表（不作用于导出所有群组时）'),
  autoClearAllDataTableEnabled: Schema.boolean().default(true).description('是否自动清空所有群组的聊天记录数据表'),
  sendFileToGroupEnabled: Schema.boolean().default(true).description('是否将文件直接发送到群组（可能无效）'),
})

declare module 'koishi' {
  interface Tables {
    chat_log_excelizer_table: ChatLogExcelizer
  }
}

export interface ChatLogExcelizer {
  id: number
  guildId: string
  userId: string
  username: string
  time: string
  content: string
}

export function apply(ctx: Context, config: Config) {
  const {
    saveDirectory,
    sendFileToGroupEnabled,
    autoClearAllDataTableEnabled,
    autoClearDataTableEnabled,
  } = config
  ctx.model.extend('chat_log_excelizer_table', {
    id: 'unsigned',
    guildId: 'string',
    userId: 'string',
    username: 'string',
    time: 'string',
    content: 'text',
  })
  // 处理当私聊的时候没有 guildId 的情况
  ctx.on('message', (session) => {
    const {
      guildId,
      userId,
      username,
      timestamp,
      content,
    } = session

    let date = new Date(timestamp);

    let time = date.toLocaleString('zh-CN', {
      timeZone: 'Asia/Shanghai',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    });

    // 如果 guildId 不存在，可以设置为默认值或留空
    const defaultGuildId = 'N/A'; // 默认值为 'N/A'
    const resolvedGuildId = guildId || defaultGuildId;

    ctx.database.create('chat_log_excelizer_table', { guildId: resolvedGuildId, userId, username, time, content })
  })
  ctx.command('chatLogExcelizer', '查看chatLogExcelizer指令帮助 ')
    .action(({ session }) => {
      session.execute(`chatLogExcelizer -h`)
    })
  ctx.command('chatLogExcelizer.exporterAll', '导出所有群组的聊天记录（不会自动清空数据表）')
    .action(async ({ session }) => {
      const chatLogs = await ctx.database.get('chat_log_excelizer_table', {})
      let result: [boolean, string]
      result = await chatLogToExcel(chatLogs, saveDirectory)
      if (result[0]) {
        if (sendFileToGroupEnabled) {
          await session.send(h.file(result[1]))
        }
        if (autoClearAllDataTableEnabled) {
          await ctx.database.remove('chat_log_excelizer_table', {})
        }
        return `导出成功！
文件路径：${result[1]}`
      }
      logger.error(result[1])
    })
  ctx.command('chatLogExcelizer.exporter', '导出当前群组的聊天记录')
    .action(async ({ session }) => {
      const { guildId } = session
      const chatLogs = await ctx.database.get('chat_log_excelizer_table', { guildId })
      let result: [boolean, string]
      result = await chatLogToExcel(chatLogs, saveDirectory)
      if (result[0]) {
        if (sendFileToGroupEnabled) {
          await session.send(h.file(result[1]))
        }
        if (autoClearDataTableEnabled) {
          await ctx.database.remove('chat_log_excelizer_table', { guildId })
        }
        return `导出成功！
文件路径：${result[1]}`
      }
      logger.error(result[1])
    })
  ctx.command('chatLogExcelizer.clearAllData', '清空所有群组的聊天记录数据表 ')
    .action(async ({ session }) => {
      await ctx.database.remove('chat_log_excelizer_table', {})
      return '清空成功！'
    })
  ctx.command('chatLogExcelizer.clearData', '清空当前群组的聊天记录数据表 ')
    .action(async ({ session }) => {
      const { guildId } = session
      await ctx.database.remove('chat_log_excelizer_table', { guildId })
      return '清空成功！'
    })

  // 定义一个函数，将聊天记录作为参数，生成 Excel 表格文件并保存在指定的文件目录
  // 参数：chatLogs - 聊天记录数组；dir - 文件目录（可选，默认为项目根目录）
  // 返回值：一个元组，包含一个布尔值和一个字符串，分别表示是否成功生成文件和文件的完整路径
  async function chatLogToExcel(chatLogs: ChatLogExcelizer[], dir?: string): Promise<[boolean, string]> {
    // 创建一个新的工作簿
    const workbook = new Excel.Workbook();

    // 添加一个工作表，并命名为 Chat Logs
    const worksheet = workbook.addWorksheet(`Chat Logs`);

    // 在工作表中添加表头，包含聊天记录的字段名
    worksheet.columns = [
      { header: 'ID', key: 'id', width: 10 },
      { header: 'Guild ID', key: 'guildId', width: 10 },
      { header: 'User ID', key: 'userId', width: 10 },
      { header: 'Username', key: 'username', width: 10 },
      { header: 'Content', key: 'content', width: 30 },
      { header: 'Time', key: 'time', width: 20 }
    ];

    // 在工作表中添加聊天记录数据
    worksheet.addRows(chatLogs);

    // 如果没有指定文件目录，那么使用项目根目录
    if (!dir) {
      dir = '.';
    }

    // 定义文件名，使用当前时间戳作为唯一标识
    const date = new Date();
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const hour = date.getHours();
    const minute = date.getMinutes();
    const second = date.getSeconds();
    const filename = `chat_logs_${year}-${month}-${day}-${hour}-${minute}-${second}.xlsx`;

    // 定义文件路径，使用 path 模块拼接文件目录和文件名
    const path = require('path').join(dir, filename);

    // 尝试将工作簿保存到文件路径
    try {
      await workbook.xlsx.writeFile(path);
      // 如果成功，返回 true 和文件路径
      return [true, path];
    } catch (error) {
      // 如果失败，返回 false 和错误信息
      return [false, error.message];
    }
  }

}
