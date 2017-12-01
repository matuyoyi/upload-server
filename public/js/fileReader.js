const xlsx = require('node-xlsx');
const nodemailer = require('nodemailer');

const workSheetsFromFile = xlsx.parse('./test.xlsx');
let rawData;					// 表格原始数据
let timeIndex;					// 时间列索引
const headerArr = []; 			// 表头
const fileData = [];			// 转换后数据
for (let i = 0; i < workSheetsFromFile.length; i++) {
	if(workSheetsFromFile[i].data.length > 0) {
		rawData = workSheetsFromFile[i].data;
	}
}
for (var i = 0; i < rawData.length; i++) {
	for (var j = 0; j < rawData[i].length; j++) {
		if(i === 0) {
			headerArr.push(rawData[i][j])
			if(rawData[i][j].indexOf('时间') > -1
			|| rawData[i][j].indexOf('日期') > -1
			|| rawData[i][j].indexOf('time') > -1
			|| rawData[i][j].indexOf('date') > -1) {
				timeIndex = j;
			}
		}
		if (timeIndex) {
			if(!isNaN(rawData[i][timeIndex])) {
				rawData[i][timeIndex] = (new Date(1900, 0, rawData[i][timeIndex])).toLocaleString();
			}
		}

	}
}
rawData.shift();
for (let i = 0; i < rawData.length; i++) {
	let excelRowData = {};
	for (let j = 0; j < rawData[i].length; j++) {
		excelRowData[headerArr[j]] = rawData[i][j];
	}
	generateData(excelRowData);
}
function generateData(data) {
	fileData.push(data);
}

for (let i = 0; i < fileData.length; i++) {
	for(let value in fileData[i]) {
		if(value === '时间') {
			if(fileData[i][value].length>8) {
				var data = fileData[i][value].replace(/[:\s-\\]/g, '')
				let time = data.slice(8)
				if(parseInt(time) >= 90000) {
					let emailAddress = fileData[i].邮箱;
					let currentNum = fileData[i].工号;
					let currentName = fileData[i].姓名;
					let currentTime = fileData[i].时间;
					mailer(emailAddress, currentNum, currentName, currentTime)
				}
			}
		}
	}
}

function mailer(emailAddress, currentNum, currentName, currentTime) {
	nodemailer.createTestAccount((err, account) => {

	    // create reusable transporter object using the default SMTP transport
	    let transporter = nodemailer.createTransport({
	        host: 'smtp.163.com',
	        port: 465,
	        //secure: false, // true for 465, false for other ports
	        auth: {
	            user: 'wtu_chennan@163.com', // generated ethereal user
	            pass: 'wtu4b520'  // generated ethereal password
	        }
	    });

	    // setup email data with unicode symbols
	    let mailOptions = {
	        from: '"chen nan 👻" <wtu_chennan@163.com>', // sender address
	        to: emailAddress, // list of receivers
	        subject: 'Hello ✔', // Subject line
	        text: 'Hello world?', // plain text body
	        html: `<p>工号${currentNum}-${currentName}打卡异常，</p><p>异常记录：<b>${currentTime}</b></p>` // html body
	    };

	    // send mail with defined transport object
	    transporter.sendMail(mailOptions, (error, response) => {
	        if (error) {
	            console.log(error);
	        } else {
	        	console.log(response)
	        }
	    });
	});
}
