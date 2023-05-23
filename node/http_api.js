const https = require('https');

/**
 * 发送请求并返回 Promise
 * @param {object} options 请求选项
 * @param {string} requestBody 请求体
 * @returns {Promise<any>} 请求结果的 Promise 对象
 */
const sendRequest = async (options, requestBody) => {
	// 使用 https 模块发送请求
	return new Promise((resolve, reject) => {
		const req = https.request(options, (res) => {
			let responseData = '';

			res.on('data', (chunk) => {
				responseData += chunk;
			});

			res.on('end', () => {
				if (res.statusCode === 200) {
					const response = JSON.parse(responseData);
					console.log('response', JSON.stringify(response, null, 4));
					resolve(response);
				} else {
					reject(new Error(`请求失败，状态码：${res.statusCode}`));
				}
			});
		});

		req.on('error', (error) => {
			reject(error);
		});

		if (requestBody !== null) {
			req.write(requestBody);
		}
		req.end();
	});
};

module.exports = { sendRequest };
