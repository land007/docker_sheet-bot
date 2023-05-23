const { sendRequest } = require('./http_api');

const sendMessage = async (botKey, message) => {
	console.log('message', message);

	const requestData = JSON.stringify(message);

	const options = {
		method: 'POST',
		hostname: 'qyapi.weixin.qq.com',
		path: `/cgi-bin/webhook/send?key=${botKey}`,
		headers: {
			'Content-Type': 'application/json',
			'Content-Length': Buffer.byteLength(requestData)
		}
	};

	try {
		const response = await sendRequest(options, requestData);
		console.log('response', response);
		return response;
	} catch (error) {
		throw error;
	}
};

module.exports = {
	sendMessage
};
