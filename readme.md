# koishi-plugin-chat-log-excelizer

[![npm](https://img.shields.io/npm/v/koishi-plugin-chat-log-excelizer?style=flat-square)](https://www.npmjs.com/package/koishi-plugin-chat-log-excelizer)

## 🎈 介绍

这是一个基于 **Koishi** 框架的机器人插件，可以将群聊的聊天记录导出为 **Excel** 表格文件，并保存在指定的文件目录。你可以使用这个插件来备份、分析或分享你的群聊历史，或者用它来做一些有趣的事情。😉

## 📦 安装

```
前往 Koishi 插件市场添加该插件即可
```

## 🎮 使用

- 建议指令添加指令别名

## 📝 指令说明

本插件提供了以下几个指令：

- `chatLogExcelizer`：查看本插件的指令帮助。
- `chatLogExcelizer.exporterAll`：导出所有群组的聊天记录（不会自动清空数据表）
- `chatLogExcelizer.exporter`：导出当前群组的聊天记录为 Excel 文件，并根据设置决定是否发送到群组或清空数据表。
- `chatLogExcelizer.clearAllData`：清空所有群组的聊天记录数据表。
- `chatLogExcelizer.clearData`：清空当前群组的聊天记录数据表。

## 🌠 后续计划

- 如有需要，则会支持私聊的聊天记录导出

## 🙏 致谢

* [Koishi](https://koishi.chat/) - 机器人框架
* [exceljs](https://github.com/exceljs/exceljs) - Excel 文件操作模块

## 📄 License

MIT License © 2023