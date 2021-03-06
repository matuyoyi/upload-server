const express = require('express');
const router = express.Router();
const fs = require('fs');
const xlsx = require('node-xlsx');
const _ = require('lodash');

/* GET home page. */
router.get('/upload-file', function (req, res, next) {
    res.send('index');
});
router.post('/upload-file', (req, res, next) => {
    const reqData = req.files.file;
    const fileName = reqData.name;
    const suffix = fileName.replace(/.+\./, '');
    if(suffix === 'csv') {
        const bufferList = reqData.data.toString().split(/\r\n+/); // convert buffer to String[]
        const fileData = [];
        bufferList.forEach(element => {
            const arr = element.split(',');
            fileData.push(element.split(','));
        });
        const clomuns = fileData.shift();
        const resData = [];
        for(let i=0;i<fileData.length;i++) {
            let itemObj = {}
            for(let j=0;j<fileData[i].length;j++) {
                const key = clomuns[j].replace(/(^\s*)|(\s*$)/g, '');
                const value = fileData[i][j].replace(/(^\s*)|(\s*$)/g, '');
                eval("itemObj." + key + "='" + value + "'");
                if(j === fileData[i].length - 1) {
                    resData.push(itemObj);
                }
            }
        }
        const data = {
            clomuns,
            data: resData
        }
        res.send(data);
    } else if (suffix === 'xlsx') {
        let originData = xlsx.parse(reqData.data);
        let timeIndex;					// Index of time column		
        const resData = [];             // responese data	
        let fileData;
        const sheetsList = _.filter(originData, function(data) {
            return data.data.length > 0;
        });
        const arr = []
        function getObject(arr1, arr2) {
            let obj = {}
            for(let i = 0;i < arr1.length; i++) {
                obj[arr1[i]] = arr2[i]
            }
            return obj;
        }
        _.each(sheetsList, (sheet) => {
            fileData = [];
            const columns = sheet.data.shift(); 			// columns
            _.each(columns, (item, i) => {
                if(item.indexOf('时间') > -1 || item.indexOf('日期') > -1 || item.indexOf('time') > -1 || item.indexOf('date') > -1 ) {
                    timeIndex = i;
                }
            })
            _.each(sheet.data, (row) => {
                let rowArr = [];
                _.each(row, (item, index) => {
                    index === timeIndex ? rowArr.push((new Date(1900, 0, item)).toLocaleString()) : rowArr.push(item);
                })
                const excelRowData = getObject(columns,rowArr);
                generateData(excelRowData);
            })
            resData.push({
                columns,
                name: sheet.name,
                data: fileData,
            })
        })
        
        function generateData(data) {
            fileData.push(data);
        }
        res.send(resData);
        
    } else if(suffix === 'docx') {
        const filePath = './public/upload/'+fileName;
        req.files.file.mv(filePath,  function(err) {
            if (err) {
                return res.status(500).send(err);
            }
            const AdmZip = require('adm-zip'); //引入查看zip文件的包
            const zip = new AdmZip(filePath); //filePath为文件路径
            let contentXml = zip.readAsText("word/document.xml");//将document.xml读取为text内容
            let str = "";
            contentXml.match(/<w:t>[\s\S]*?<\/w:t>/ig).forEach((item)=>{
                str += item.slice(5,-6)+'\r\n'
            });
            res.send(str);
        })
    } else {
        res.status(500).send(err);
    }
})
function sendMail(param) {
	nodemailer.createTestAccount((err, account) => {
	    // create reusable transporter object using the default SMTP transport
	    let transporter = nodemailer.createTransport({
	        host: 'smtp.163.com',
	        port: 465,
	        auth: {
	            user: 'wtu_chennan@163.com', // generated ethereal user
	            pass: 'wtu4b520'  // generated ethereal password
	        }
	    });

	    // setup email data with unicode symbols
	    let mailOptions = {
	        from: '"chen nan 👻" <wtu_chennan@163.com>', // sender address
	        to: param.emailAddress, // list of receivers
	        subject: 'Hello ✔', // Subject line
	        text: 'Hello world?', // plain text body
	        html: `<p>工号${param.currentNum}-${param.currentName} 记录异常，</p><p>异常记录：<b>${param.currentTime}</b></p>` // html body
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

module.exports = router;
