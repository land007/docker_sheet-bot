const http = require('http');
const url = require('url');
const { addRowToEndWithSheet } = require('./doc_api');
const { sendMessage } = require('./bot_api');
const { convertToMarkdownTable } = require('./makedown_api');

const docId = process.env['doc_id'];
console.log('docId', docId);

//const sheetId = 'v5jt5g';//5月

const shareUrl = process.env['share_url'];
console.log('shareUrl', shareUrl);

const botKey = process.env['bot_key'];//糖猫
console.log('botKey', botKey);

// 定时器时间间隔（单位：毫秒）
const interval = 5000; // 5秒

// 缓存数组，用于积攒待发送的消息
const rowDataCache = [];

/**
 * 获取本地时间的中国格式时间戳（东八区）
 * @returns {string} 格式化后的时间戳，例如：2023-05-22 12:34:56.789
 */
const getLocalTimestamp = () => {
	const localTime = new Date();
	const timezoneOffset = 8 * 60; // 东八区偏移量为8小时
	const utcTime = new Date(localTime.getTime() + timezoneOffset * 60 * 1000);

	const year = utcTime.getFullYear();
	const month = (utcTime.getMonth() + 1).toString().padStart(2, '0');
	const day = utcTime.getDate().toString().padStart(2, '0');
	const hours = utcTime.getHours().toString().padStart(2, '0');
	const minutes = utcTime.getMinutes().toString().padStart(2, '0');
	const seconds = utcTime.getSeconds().toString().padStart(2, '0');
	const milliseconds = utcTime.getMilliseconds().toString().padStart(3, '0');

	return `${year}-${month}-${day} ${hours}:${minutes}:${seconds} ${milliseconds}`;
};

// 添加消息到缓存
const addToCache = (rowData) => {
	const timestamp = getLocalTimestamp();
	// 在数组的第一个位置增加一项可读时间戳
	for (let i = 0; i < rowData.length; i++) {
		rowData[i].values.unshift({
			"cell_value": {
				"text": timestamp
			}
		});
	}
	Array.prototype.push.apply(rowDataCache, rowData);
};

// 发送缓存中的消息
const sendRowDataCache = async () => {
	if (rowDataCache.length > 0) {
		// 合并缓存中的消息为一个大的消息对象
		try {
			console.log('rowDataCache', JSON.stringify(rowDataCache, null, 4));
			// 调用发送消息的方法，将缓存中的消息一次性发送出去
			const newRow = await addRowToEndWithSheet(undefined, docId, rowDataCache); // 使用 await 关键字等待异步操作的结果
			console.log('newRow:', newRow);

			// 清空缓存
			rowDataCache.length = 0;
		} catch (error) {
			console.error('Error:', error);
		}
	}
};

// 创建定时器，定期发送缓存中的消息
setInterval(sendRowDataCache, interval);

// 创建服务器
const server = http.createServer((req, res) => {
	// 解析请求的URL
	const parsedUrl = url.parse(req.url, true);
	const path = parsedUrl.pathname;
	const method = req.method;

	// 处理 POST 请求
	if (method === 'POST' && path === '/api/addRow') {
		// 获取请求的数据
		let requestBody = '';
		req.on('data', chunk => {
			requestBody += chunk;
		});

		req.on('end', async () => {
			// 解析请求体数据
			const requestData = JSON.parse(requestBody);

			// 调用 addRowToEnd 方法处理数据
			const rowData = requestData.rowData;
			addToCache(rowData);
			//const newRow = await addRowToEnd(undefined, docId, sheetId, rowData); // 使用 await 关键字等待异步操作的结果
			const formTitle = '订单列表';
			const formLink = '[' + formTitle + '](' + shareUrl + ')';
			const jsonFormatted = '```json\n' + JSON.stringify(rowData, null, 2) + '\n```';
			//const markdownTable = convertToMarkdownTable(rowData);
			const message = {
				msgtype: 'markdown',
				markdown: {
					content: formLink + '\n' + jsonFormatted
				}
			};
			const result = await sendMessage(botKey, message);
			// 返回响应
			res.statusCode = 200;
			res.setHeader('Content-Type', 'application/json');
			res.end(JSON.stringify({ success: true }));//, newRow: newRow
		});
	} else {
		// 处理未知路由
		res.statusCode = 404;
		res.setHeader('Content-Type', 'application/json');
		res.end(JSON.stringify({ success: false, message: 'Route not found' }));
	}
});

// 启动服务器
const port = 80;
server.listen(port, () => {
	console.log(`Server is running on port ${port}`);
});
