const xlsx = require('node-xlsx')
const fs = require('fs')

const inputExcelPath = './input/工作簿3.xlsx'

const workSheetsFromFile = xlsx.parse(inputExcelPath)

fs.writeFileSync('./output/sql.txt', '', (error) => {
  if (error) {
    console.log(error)
  } else {
    console.log('清空')
  }
})

let successCount = 0

workSheetsFromFile[0].data.forEach((item) => {
  const fixed_assets_number = item[1]
  const assets_name = item[2]
  const assets_config = item[3]
  const count = item[4]
  const assets_cause = item[6]
  const address = item[11]
  const department = item[10]

  const sql = `update fixed_assets set assets_name='${assets_name}',assets_config='${assets_config}',
    count='${count}',assets_cause='${assets_cause}',address='${address}',department='${department}' 
    where fixed_assets_number = '${fixed_assets_number}';\n`

  fs.appendFileSync('./output/sql.txt', sql, (error) => {
    if (error) {
      console.log(error)
    } else {
      console.log('添加成功')
      successCount++
    }
  })
})

console.log(`共生成了${successCount}条sql`)
