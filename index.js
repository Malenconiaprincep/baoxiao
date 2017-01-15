const co = require('co')
const fs = require('fs');
const xlsx = require('node-xlsx');
let data = xlsx.parse(fs.readFileSync('./kq.xlsx')); // parses a buffer
let destPath = 'zj.xls'
let destFile = xlsx.parse(fs.readFileSync(destPath)); // parses a buffer
let arr = []
let name = '汪波'

// 解析考勤
function parseList(data, sheetNum) {
  if (!sheetNum) {
    sheetNum = 0
  }
  let list = data[sheetNum].data
  return list
}

function filterList(list) {
  return list.filter(item => {
    return item[2] === name
  })
}

function getTime(time) {
  time = time.split(':')
  let hours = time[0]
  let minutes = time[1]
  let seconds = time[2]

  let date = new Date()
  date.setHours(hours)
  date.setMinutes(minutes)
  date.setSeconds(seconds)
  return Math.ceil(date.getTime() / 1000)
}

// 检测上班时间
function check(end, start, special) {
  // 过节至少4个小时
  if (special) {
    return (getTime(end) - getTime(start)) >= 4 * 3600
  }
  return (getTime(end) - getTime(start)) >= 8 * 3600
}

// 获取加班费档次
function level(end, start, special) {
  let endHours = end.split(':')[0]
  let price = 0

  // 平时
  if (endHours >= 20 && endHours < 22) {
      price = 20
  }

  if (endHours >= 22) {
      price = 35
  }

  // 休息
  if (special) {
    if ((getTime(end) - getTime(start)) >= 4 * 3600 && (getTime(end) - getTime(start)) < 8 * 3600 ) {
      price = 35
    } else {
      price = 50
    }
  }

  return price
}

function compose(item) {
  for (var i=0; i < item.length; i++) {
    if (!arr[i]) {
      arr[i] = []
    }
    if (!item[i]) {
      item[i] = ''
    }
    arr[i].push(item[i])
  }
}




co(function *(){
    let list = parseList(data)
    list = filterList(list)
    for(var i=0; i < list.length; i++) {
      let end = list[i][6]
      let start = list[i][4]
      let day = list[i][3]
      let special = false
      special = /六|日/.test(day)


      if(end && start && check(end, start, special)) {
          if(level(end, start, special)) {
            arr.push(list[i])
          }
      }
    }


    let originlist = parseList(destFile, 3)
    originlist = originlist.concat(arr)
    destFile[3].data = originlist
    
    let buffer = xlsx.build(destFile);
    fs.writeFile(destPath, buffer , function(err, data){
      console.log(err)
      console.log(data)
    })

}).catch(err => {
    console.log(err)
})
