const your_corp_id = process.env['your_corp_id'];
console.log('your_corp_id', your_corp_id);
const your_corp_secret = process.env['your_corp_secret'];
console.log('your_corp_secret', your_corp_secret);

const { sendRequest } = require('./http_api');

/**
 * 获取访问令牌
 * @param {string} corpId 企业 ID
 * @param {string} corpSecret 企业密钥
 * @returns {Promise<string>} 访问令牌
 */
const getAccessToken = async (corpId, corpSecret) => {
	const options = {
		method: 'GET',
		hostname: 'qyapi.weixin.qq.com',
		path: `/cgi-bin/gettoken?corpid=${corpId}&corpsecret=${corpSecret}`
	};

	return sendRequest(options, null)
		.then(response => {
			if (response && response.access_token) {
				return response.access_token;
			} else {
				throw new Error('无法获取访问令牌');
			}
		})
		.catch(error => {
			throw error;
		});
};

/**
 * 创建 WeDrive 空间
 * @param {string} spaceName 空间名称
 * @param {string} accessToken 访问令牌
 * @returns {Promise<string>} 空间 ID
 */
const createWeDriveSpace = async (spaceName, accessToken) => {
	const requestData = {
		space_name: spaceName,
		auth_info: [
			{ type: 1, userid: 'JiaYiQiu', auth: 7 },
			//{ type: 1, userid: 'QinXuYing', auth: 7 }
		],
		space_sub_type: 0
	};

	const options = {
		method: 'POST',
		hostname: 'qyapi.weixin.qq.com',
		path: `/cgi-bin/wedrive/space_create?access_token=${accessToken}`,
		headers: {
			'Content-Type': 'application/json'
		}
	};

	return sendRequest(options, JSON.stringify(requestData))
		.then(response => {
			console.log('response', response);
			if (response && response.spaceid) {
				return response.spaceid;
			} else {
				throw new Error('无法创建 WeDrive 空间');
			}
		})
		.catch(error => {
			throw error;
		});
};

/**
 * 解散空间
 * @param {string} accessToken 访问令牌
 * @param {string} spaceId 空间ID
 * @returns {Promise<Object>} 解散空间的结果
 */
const dismissSpace = async (accessToken, spaceId) => {
	// 构造请求数据
	const requestData = {
		spaceid: spaceId
	};
	// 构造请求选项
	const options = {
		method: 'POST',
		hostname: 'qyapi.weixin.qq.com',
		path: `/cgi-bin/wedrive/space_dismiss?access_token=${accessToken}`,
		headers: {
			'Content-Type': 'application/json'
		}
	};

	try {
		// 发送请求并等待结果
		const response = await sendRequest(options, JSON.stringify(requestData));

		// 检查结果是否包含错误码
		if (response.errcode === 0) {
			// 解散空间成功
			console.log('解散空间成功:', JSON.stringify(response));
		} else {
			// 解散空间失败
			console.error('解散空间失败:', response.errmsg);
		}

		// 返回解散空间的结果
		return response;
	} catch (error) {
		// 处理请求失败的情况
		console.error('解散空间请求失败:', error);
		throw error;
	}
};


/**
 * 创建 WeDrive 文件夹
 * @param {string} spaceId 空间 ID
 * @param {string} folderName 文件夹名称
 * @param {string} parentFolderId 父文件夹 ID
 * @param {number} folderType 文件夹类型
 * @param {string} accessToken 访问令牌
 * @returns {Promise<string>} 文件夹 ID
 */
const createWeDriveFolder = async (spaceId, folderName, parentFolderId, folderType, accessToken) => {
	const requestData = {
		spaceid: spaceId,
		file_name: folderName,
		fatherid: parentFolderId,
		file_type: folderType
	};

	const options = {
		method: 'POST',
		hostname: 'qyapi.weixin.qq.com',
		path: `/cgi-bin/wedrive/file_create?access_token=${accessToken}`,
		headers: {
			'Content-Type': 'application/json'
		}
	};

	return sendRequest(options, JSON.stringify(requestData))
		.then(response => {
			if (response && response.fileid) {
				return response.fileid;
			} else {
				throw new Error('无法创建 WeDrive 文件夹');
			}
		})
		.catch(error => {
			throw error;
		});
};

/**
 * 创建文档
 * @param {string} accessToken 访问令牌
 * @param {string} spaceId 空间 ID
 * @param {string} fatherId 父文件夹 ID
 * @param {number} docType 文档类型
 * @param {string} docName 文档名称
 * @param {Array<string>} adminUsers 管理员用户列表
 * @returns {Promise<{ docId: string, url: string }>} 文档 ID 和 URL
 */
const createDocument = async (accessToken, spaceId, fatherId, docType, docName, adminUsers) => {
	const requestData = {
		doc_name: docName,
		fatherid: fatherId,
		spaceid: spaceId,
		doc_type: docType,
		admin_users: adminUsers
	};

	const options = {
		method: 'POST',
		hostname: 'qyapi.weixin.qq.com',
		path: `/cgi-bin/wedoc/create_doc?access_token=${accessToken}`,
		headers: {
			'Content-Type': 'application/json'
		}
	};

	return sendRequest(options, JSON.stringify(requestData))
		.then(response => {
			if (response && response.docid && response.url) {
				return {
					docid: response.docid,
					url: response.url
				};
			} else {
				throw new Error('无法创建文档');
			}
		})
		.catch(error => {
			throw error;
		});
};

/**
 * 获取表格行列信息
 * @param {string} accessToken 访问令牌
 * @param {string} docId 文档 ID
 * @returns {Promise<Object[]>} 表格的工作表属性信息
 */
const getSheetProperties = async (accessToken, docId) => {
	const requestData = {
		docid: docId
	};

	const options = {
		method: 'POST',
		hostname: 'qyapi.weixin.qq.com',
		path: `/cgi-bin/wedoc/spreadsheet/get_sheet_properties?access_token=${accessToken}`,
		headers: {
			'Content-Type': 'application/json'
		}
	};

	try {
		const response = await sendRequest(options, JSON.stringify(requestData));
		if (response && response.errcode === 0 && response.properties) {
			return response.properties;
		} else {
			throw new Error('无法获取表格行列信息');
		}
	} catch (error) {
		throw error;
	}
};


/**
 * 读取表格数据
 * @param {string} accessToken 访问令牌
 * @param {string} docId 文档 ID
 * @param {string} sheetId 表格工作表 ID
 * @param {string} range 数据范围
 * @returns {Promise<object>} 表格数据
 */
const readSheetData = async (accessToken, docId, sheetId, range) => {
	const requestData = {
		docid: docId,
		sheet_id: sheetId,
		range: range
	};

	const options = {
		method: 'POST',
		hostname: 'qyapi.weixin.qq.com',
		path: `/cgi-bin/wedoc/spreadsheet/get_sheet_range_data?access_token=${accessToken}`,
		headers: {
			'Content-Type': 'application/json'
		}
	};

	return sendRequest(options, JSON.stringify(requestData))
		.then(response => {
			if (response && response.errcode === 0 && response.grid_data) {
				return response.grid_data;
			} else {
				throw new Error('无法读取表格数据');
			}
		})
		.catch(error => {
			throw error;
		});
};

/**
 * 添加表格工作表
 * @param {string} accessToken 访问令牌
 * @param {string} docId 文档 ID
 * @param {object} sheetData 表格工作表数据
 * @returns {Promise<string>} 表格工作表 ID
 */
const addSheet = async (accessToken, docId, sheetData) => {
	const requestData = {
		docid: docId,
		requests: [
			{
				add_sheet_request: sheetData
			}
		]
	};

	const options = {
		method: 'POST',
		hostname: 'qyapi.weixin.qq.com',
		path: `/cgi-bin/wedoc/spreadsheet/batch_update?access_token=${accessToken}`,
		headers: {
			'Content-Type': 'application/json'
		}
	};

	return sendRequest(options, JSON.stringify(requestData))
		.then(response => {
			if (response && response.responses && response.responses[0].add_sheet_response && response.responses[0].add_sheet_response.properties && response.responses[0].add_sheet_response.properties.sheet_id) {
				return response.responses[0].add_sheet_response.properties.sheet_id;
			} else {
				throw new Error('无法添加表格工作表');
			}
		})
		.catch(error => {
			throw error;
		});
};

/**
 * 删除表格工作表
 * @param {string} accessToken 访问令牌
 * @param {string} docId 文档 ID
 * @param {string} sheetId 表格工作表 ID
 */
const deleteSheet = async (accessToken, docId, sheetId) => {
	const requestData = {
		docid: docId,
		requests: [
			{
				delete_sheet_request: {
					sheet_id: sheetId
				}
			}
		]
	};

	const options = {
		protocol: 'https:',
		hostname: 'qyapi.weixin.qq.com',
		path: `/cgi-bin/wedoc/spreadsheet/batch_update?access_token=${accessToken}`,
		method: 'POST',
		headers: {
			'Content-Type': 'application/json'
		}
	};

	try {
		const response = await sendRequest(options, JSON.stringify(requestData));
		console.log('删除工作表成功:', JSON.stringify(response));
		//{"errcode":0,"errmsg":"ok","responses":[{"delete_sheet_response":{"sheet_id":"7z6qfv"}}]}
	} catch (error) {
		console.error('删除工作表失败:', error);
	}
};

/**
 * 更新表格范围数据
 * @param {string} accessToken 访问令牌
 * @param {string} docId 文档 ID
 * @param {string} sheetId 表格工作表 ID
 * @param {object} rangeData 范围数据
 * @returns {Promise<{ updatedCells: number }>} 更新的单元格数量
 */
const updateRange = async (accessToken, docId, sheetId, rangeData) => {
	const requestData = {
		docid: docId,
		requests: [
			{
				update_range_request: {
					sheet_id: sheetId,
					grid_data: rangeData
				}
			}
		]
	};

	const options = {
		method: 'POST',
		hostname: 'qyapi.weixin.qq.com',
		path: `/cgi-bin/wedoc/spreadsheet/batch_update?access_token=${accessToken}`,
		headers: {
			'Content-Type': 'application/json'
		}
	};

	return sendRequest(options, JSON.stringify(requestData))
		.then(response => {
			if (response && response.errcode === 0) {
				return {
					updatedCells: response.responses[0].update_range_response.updated_cells
				};
			} else {
				throw new Error('更新表格范围数据失败');
			}
		})
		.catch(error => {
			throw error;
		});
};

/**
 * 删除表格维度数据
 * @param {string} accessToken 访问令牌
 * @param {string} docId 文档 ID
 * @param {string} sheetId 表格工作表 ID
 * @param {object} dimensionData 维度数据
 * @returns {Promise<void>}
 */
const deleteDimension = async (accessToken, docId, sheetId, dimensionData) => {
	const requestData = {
		docid: docId,
		requests: [
			{
				delete_dimension_request: {
					sheet_id: sheetId,
					...dimensionData
				}
			}
		]
	};

	const options = {
		method: 'POST',
		hostname: 'qyapi.weixin.qq.com',
		path: `/cgi-bin/wedoc/spreadsheet/batch_update?access_token=${accessToken}`,
		headers: {
			'Content-Type': 'application/json'
		}
	};

	return sendRequest(options, JSON.stringify(requestData))
		.then(() => {
			// 成功删除表格维度数据，不返回任何结果
		})
		.catch(error => {
			throw error;
		});
};

/**
 * 在指定工作表的行末尾新增数据
 * @param {string} accessToken 访问令牌
 * @param {string} docId 文档 ID
 * @param {string} sheetId 工作表 ID
 * @param {Array} rowData 新增的行数据
 * @returns {Promise} 行数据新增结果
 */
const addRowToEnd = async (accessToken, docId, sheetId, rowData) => {
	try {
		console.log('rowData', JSON.stringify(rowData, null, 4));
		if (accessToken === undefined) {
			accessToken = await getAccessToken(your_corp_id, your_corp_secret);
		}
		console.log('accessToken', accessToken);

		// 获取工作表的行列信息
		const properties = await getSheetProperties(accessToken, docId);

		// 确定新增行的起始行号
		const startRow = properties.find(sheet => sheet.sheet_id === sheetId).row_count;

		//		const timestamp = getLocalTimestamp();
		//		// 在数组的第一个位置增加一项可读时间戳
		//		for (let i = 0; i < rowData.length; i++) {
		//			rowData[i].values.unshift({
		//				"cell_value": {
		//					"text": timestamp
		//				}
		//			});
		//		}

		const rangeData = {
			"start_row": startRow,
			"start_column": 0,
			"rows": rowData
		};

		// 更新指定范围的数据
		const updateResult = await updateRange(accessToken, docId, sheetId, rangeData);

		// 返回行数据新增结果
		return updateResult;
	} catch (error) {
		throw error;
	}
};

/**
 * 获取当前的年份和月份，并返回数字格式的字符串表示。
 * @returns {string} 当前年份和月份的数字字符串，例如：202305
 */
const getCurrentYearAndMonth = () => {
  const localTime = new Date();
  const timezoneOffset = 8 * 60; // 东八区偏移量为8小时
  const utcTime = new Date(localTime.getTime() + timezoneOffset * 60 * 1000);

  const year = utcTime.getFullYear().toString(); // 获取年份并转换为字符串格式
  const month = (utcTime.getMonth() + 1).toString().padStart(2, '0'); // 获取月份并补齐两位

  return year + month; // 返回年份和月份的数字字符串
};

/**
 * 在指定文档中的当前月份工作表的行末尾新增数据
 * @param {string} accessToken 访问令牌
 * @param {string} docId 文档 ID
 * @param {Array} rowData 新增的行数据
 * @returns {Promise} 行数据新增结果
 */
const addRowToEndWithSheet = async (accessToken, docId, rowData) => {
	try {
        console.log('rowData', JSON.stringify(rowData, null, 4));
		if (accessToken === undefined) {
			accessToken = await getAccessToken(your_corp_id, your_corp_secret);
		}
		// 获取工作表的行列信息
		const properties = await getSheetProperties(accessToken, docId);

		// 获取当前月份的工作表
        //const currentMonth = new Date().toLocaleString('default', { month: 'long' });
        //const currentMonth = new Date().toLocaleString('zh-CN', { month: 'long' });
        const currentMonth = getCurrentYearAndMonth();
		const currentSheet = properties.find(sheet => sheet.title === currentMonth);

		let sheetId;
		if (currentSheet) {
			sheetId = currentSheet.sheet_id;
		} else {
			// 创建当前月份的工作表
			const sheetData = {
				title: currentMonth,
				row_count: 10,
				column_count: 10
			};
			sheetId = await addSheet(accessToken, docId, sheetData);
		}

		// 添加数据到指定工作表的行末尾
		const newRow = await addRowToEnd(accessToken, docId, sheetId, rowData);
		return newRow;
	} catch (error) {
		throw error;
	}
};

module.exports = {
	addRowToEnd,
	addRowToEndWithSheet
};