const fs = require('fs-extra')
const xlsx = require('node-xlsx')

const readXls = (xlsBuffer, sheet = 0) => {
  const xlsArr = xlsx.parse(xlsBuffer)[sheet].data
  const res = []
  for (let i = 0; i < xlsArr.length; i++) {
    const line = xlsArr[i];
    res.push(line);
  }
  return res
}

const sortXls = (xlsArr, col = 1, row= 2) => {
  const saveHeader = xlsArr.splice(0, row)
  xlsArr.sort((a, b) => a[col] - b[col])
  return saveHeader.concat(xlsArr)
}

const splitXls = (xlsArr, splitCount = 5000, saveTitle = 2) => {
  const saveHeader = xlsArr.splice(0, saveTitle)
  let count = 0
  while (xlsArr.length) {
    let data = saveHeader.concat(xlsArr.splice(0, splitCount))
    var buffer = xlsx.build([{name: 'Sheet1', data: data}])
    fs.writeFileSync(`${++count}.xlsx`, buffer)
  }

}

let xls = fs.readFileSync('./test.xlsx')
const data = sortXls(readXls(xls))
splitXls(data)
// console.log(data.map(x => JSON.stringify(x)).join('\n'))
