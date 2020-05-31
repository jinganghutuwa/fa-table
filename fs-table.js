/*
 * @Author: Jason
 * @Date: 2019-07-31 10:06:09
 * @version: 
 * @LastEditors: Jason
 * @LastEditTime: 2020-05-31 12:00:24
 * @Descripttion: 
 */
/* 
支持读写Excel的node.js模块
node-xlsx: 基于Node.js解析excel文件数据及生成excel文件，支持xls及xlsx格式文件；
excel-parser: 基于Node.js解析excel文件数据，支持xls及xlsx格式文件；
excel-export : 基于Node.js将数据生成导出excel文件，生成文件格式为xlsx；
node-xlrd: 基于node.js从excel文件中提取数据，仅支持xls格式文件。 
*/
const fs = require('fs')
const path = require("path")
const xlsx = require('node-xlsx')


fs.readdir(path.join(__dirname,"excel"), (err, file) => {
    if (err) {
        console.log('读取文件夹失败', err)
    } else {
        console.log('获取指定文件夹下信息：', file)
        let fileList = file.filter(i => {
            return /xlsx?$/g.test(i)
        }) // 也可以使用i.endsWith('.xls') || i.endsWith('.xlsx')判断  或者  path.extname(path)
        console.log('符合条件的文件列表：', fileList)
        if(!fileList || (fileList && fileList.length === 0)){
            console.log('无符合要求的文件')
            return
        }
        
        fileList.forEach(i => {
            let obj = xlsx.parse(path.join(__dirname,"excel",i));
            // let obj = xlsx.parse(__dirname + "/" + i);
            console.log("excel表格的所有数据",obj)
            obj.forEach(item => {
                let xlsxRes = item.data
                console.log(`读取的${i}的信息为：`, xlsxRes);
                let trHtml = ''
                xlsxRes.forEach((item, idx) => {
                    console.log(item)
                    let tdHtml = ''
                    item.forEach((j, idx) => {
                        tdHtml += `<td>${j}</td>`
                    })
                    trHtml += `<tr>${tdHtml}</tr>`
                })
                let html = `<table>${trHtml}</table>`
                console.log('获取到的html为', html)
                /* 
                str.substring(startIndex,endIndex)  忽略endIndex则返回从startIndex到字符串尾字符
                str.substr(startIndex,length)  忽略length则返回从startIndex到字符串尾字符 
                */
                let resultName = i.substring(0,i.indexOf('.')) + item.name
                fs.writeFile(path.join(__dirname, 'html', resultName + '.html'), html, (err) => {
                    if (err) {
                        console.log('写入文件出错', err)
                    } else {
                        console.log('写入完成')
                    }
                })
            })
            // let xlsxRes = obj[0].data
            // console.log(`读取的${i}的信息为：`, xlsxRes);
            //     let trHtml = ''
            //     xlsxRes.forEach((item, idx) => {
            //         console.log(item)
            //         let tdHtml = ''
            //         item.forEach((j, idx) => {
            //             tdHtml += `<td>${j}</td>`
            //         })
            //         trHtml += `<tr>${tdHtml}</tr>`
            //     })
            //     let html = `<table>${trHtml}</table>`
            //     console.log('获取到的html为', html)
            //     /* 
            //       str.substring(startIndex,endIndex)  忽略endIndex则返回从startIndex到字符串尾字符
            //       str.substr(startIndex,length)  忽略length则返回从startIndex到字符串尾字符 
            //     */
            //     let resultName = i.substring(0,i.indexOf('.'))
            //     fs.writeFile(__dirname + '/' + resultName + '.html', html, (err) => {
            //         if (err) {
            //             console.log('写入文件出错', err)
            //         } else {
            //             console.log('写入完成')
            //         }
            //     })
        });
    }
})
// let obj = xlsx.parse(__dirname + "/jc_rang.xls");
// let xlsxRes = obj[0].data
// console.log('读取的Excel的信息为：',xlsxRes);
// let trHtml = ''
// xlsxRes.forEach((item,idx)=>{
//     console.log(item)
//     let tdHtml = ''
//     item.forEach((i,idx)=>{
//         tdHtml += `<td>${i}</td>`
//     })
//     trHtml += `<tr>${tdHtml}</tr>`
// })
// let html = `<table>${trHtml}</table>`
// console.log('获取到的html为',html)
// fs.writeFile(__dirname + '/result.html',html,(err)=>{
//     if(err){
//         console.log('写入文件出错',err)
//     }else{
//         console.log('写入完成')
//     }
// })