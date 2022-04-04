const XLSX = require('xlsx')
const fs = require('fs')

const workbook = XLSX.readFile('./Tahsil_Codes.xls')

const data = workbook.Sheets.Sheet1
const states = []
const districts = []
const tehsils = []

for (let i = 1; i <= 35; i++) {
  let districtCodes = []
  let state = null

  for (let rowNumber = 7; rowNumber < 6198; rowNumber++) {
    if (+data['A' + rowNumber].w === i) {
      if (!state) {
        state = {
          stateCode: data['A' + rowNumber].w,
          name: data['D' + rowNumber].w,
        }
      }
      if (
        !districtCodes.find(dc => dc === data['B' + rowNumber].w) &&
        data['B' + rowNumber].w !== '00'
      ) {
        districtCodes.push(data['B' + rowNumber].w)
      }
    }
  }

  for (const dc of districtCodes) {
    for (let rowNumber = 7; rowNumber < 6198; rowNumber++) {
      if (+data['A' + rowNumber].w === i && data['B' + rowNumber].w === dc) {
        if (data['C' + rowNumber].w === '0000') {
          districts.push({
            stateCode: state.stateCode,
            districtCode: dc,
            name: data['D' + rowNumber].w,
          })
        } else {
          tehsils.push({
            stateCode: state.stateCode,
            districtCode: districts[districts.length - 1].districtCode,
            tehsilCode: data['C' + rowNumber].w,
            name: data['D' + rowNumber].w,
          })
        }
      }
    }
  }

  states.push(state)
}

fs.writeFile('./states.json', JSON.stringify(states), 'utf8', function (err) {
  if (err) {
    console.log('An error occured while writing JSON Object to File.')
    return console.log(err)
  }
  console.log('JSON file has been saved.')
})

fs.writeFile(
  './districts.json',
  JSON.stringify(districts),
  'utf8',
  function (err) {
    if (err) {
      console.log('An error occured while writing JSON Object to File.')
      return console.log(err)
    }
    console.log('JSON file has been saved.')
  }
)

fs.writeFile('./tehsils.json', JSON.stringify(tehsils), 'utf8', function (err) {
  if (err) {
    console.log('An error occured while writing JSON Object to File.')
    return console.log(err)
  }
  console.log('JSON file has been saved.')
})
