const Excel = require('exceljs');
var workbook = new Excel.Workbook();
const filename = "task_act_time_2.xlsx"
const fs = require('fs');
var _ = require('lodash');
let data = []
function randomInteger(min, max) {
  // получить случайное число от (min-0.5) до (max+0.5)
  let rand = min - 0.5 + Math.random() * (max - min + 1);
  return Math.round(rand);
}

workbook.xlsx.readFile(filename)
  .then(function() {
    workbook.eachSheet(function(worksheet, sheetId) {
      let proccess = {}
      proccess.tasks = {}
      worksheet.eachRow(function(row, rowNumber) {
        if(row.values[1]!="name_lvl_0"&&row.values[1]!=""&&row.values[1]!=" "){
          // deadline_1 -> name_1 -> actor_1 -> deadline_2 -> name_2 -> actor_2
          const deadline_1 = row.values[6].toString()
          const name_1 = row.values[4]
          const actor_1 = row.values[5]
          const deadline_2 = row.values[9].toString()
          const name_2 = row.values[7]
          const actor_2 = row.values[8]
          
          proccess.name = row.values[1]
          proccess.depName = row.values[2]
          proccess.deadline = row.values[3]
          if (proccess.tasks[deadline_1] == undefined) proccess.tasks[deadline_1] = {}
          if (proccess.tasks[deadline_1][name_1] == undefined ) proccess.tasks[deadline_1][name_1] = {}
          if (proccess.tasks[deadline_1][name_1][actor_1] == undefined ) proccess.tasks[deadline_1][name_1][actor_1] = {}
          if (proccess.tasks[deadline_1][name_1][actor_1][deadline_2] == undefined ) proccess.tasks[deadline_1][name_1][actor_1][deadline_2] = {}
          if (proccess.tasks[deadline_1][name_1][actor_1][deadline_2][name_2] == undefined ) proccess.tasks[deadline_1][name_1][actor_1][deadline_2][name_2] = {}
          if (proccess.tasks[deadline_1][name_1][actor_1][deadline_2][name_2][actor_2] == undefined ) proccess.tasks[deadline_1][name_1][actor_1][deadline_2][name_2][actor_2] = {}
          // if (proccess.tasks[row.values[4]] == undefined) {
          //   proccess.tasks[row.values[4]] = { deps: { }, deadline: row.values[6] }
          //   proccess.tasks[row.values[4]].deps[row.values[5]]={tasks:{}}
          // }
          // if (proccess.tasks[row.values[4]].deps[row.values[5]].tasks[row.values[6]] == undefined) {proccess.tasks[row.values[4]].deps[row.values[5]].tasks[row.values[6]] = {}}
          // if (proccess.tasks[row.values[4]].deps[row.values[5]].tasks[row.values[6]][row.values[7]] == undefined) proccess.tasks[row.values[4]].deps[row.values[5]].tasks[row.values[6]][row.values[7]] = {}
          // if (proccess.tasks[row.values[4]].deps[row.values[5]].tasks[row.values[6]][row.values[7]] == undefined) proccess.tasks[row.values[4]].deps[row.values[5]].tasks[row.values[6]][row.values[7]] = 
          // if (proccess.tasks[row.values[4]].tasks[row.values[7]] == undefined) proccess.tasks[row.values[4]].tasks[row.values[7]] = { dep: [row.values[8]], deadline: row.values[9], childes:[] }
          // else {
          //   proccess.tasks[row.values[4]].dep.push(row.values[5])
          //   proccess.tasks[row.values[4]].dep = _.uniq(proccess.tasks[row.values[4]].dep)
          // }
          // for(let i=0; i<randomInteger(2,200); i++){
          //   proccess.tasks[row.values[4]].tasks[row.values[7]].childes.push({name: "Учреждение " + (i + 1), status: Math.random()<=0.95?"success":"fail"})
          // }
        }
      });
      data.push(proccess)
      console.log('file done')
    });
    fs.writeFile("data.json", JSON.stringify(data), function(err) {
      if(err) {
          return console.log(err);
      }
      console.log("The file was saved!");
  }); 
  });