const fs = require('fs');
const XLSX = require('xlsx');

// 定义需要读取的文件路径
const filePath = '../test.xlsx';

// 读取XLSX文件
const fileData = fs.readFileSync(filePath);
const workbook = XLSX.read(fileData, { type: 'buffer' });
const sheetName = workbook.SheetNames[0]; // 获取第一个Sheet的名称
const worksheet = workbook.Sheets[sheetName];

const columnBValues = [];

// 遍历工作表的单元格，提取指定列（这里是B列）的数据
for (const cellAddress in worksheet) {
  if (cellAddress.startsWith('G')) {
    const cell = worksheet[cellAddress];
    const cellValue = cell.v;
    columnBValues.push(cellValue);
  }
}

// 将B列的数据倒序排列

// // 将B列的内容进行倒序排列25/11/2015 22:50:30
let columnBValues_ = [];
    columnBValues.forEach(v=>{
      let _ = v.split(" ")[0];
      let _d = v.split(" ")[1];
      let n = _.split("/").reverse().join("/");
      // console.log(n)
      v = n + _d;
      columnBValues_.push(n + " " + _d);
      // console.log(v)
    });
// columnBValues.reverse();

// 更新工作表中B列的单元格数据
for (let i = 1; i < columnBValues_.length; i++) {
  const cellAddress = `G${i + 1}`;
  const cellValue = columnBValues_[i];
  // console.log(cellValue)
  worksheet[cellAddress].v = cellValue;
}

// 生成新的Workbook对象
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, worksheet, sheetName);

// 将新的Workbook对象写入文件
const newFilePath = './2.xlsx';
const newFileData = XLSX.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });
fs.writeFileSync(newFilePath, newFileData);

console.log('文件已更新');