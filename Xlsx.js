const XLSX = require ('xlsx')
const moment = require('moment')
const fs = require('fs');
const _ = require('underscore')
const { replaceOne } = require('../json data/mongoose1/data')
const workbook = XLSX.readFile('Assignment_Timecard.xlsx')
// console.log('Current working directory:', process.cwd());

// const worksheet = workbook.Sheets[workbook.SheetNames[0]]
const worksheet = workbook.Sheets[workbook.SheetNames[0]]
// const employeeData = JSON.parse(worksheet)

const employeeData = XLSX.utils.sheet_to_json(worksheet)
// console.log(employeeData);
// console.log(employeeData);


// const data = [employeeData]

// console.log(data);

// const filterebyDays = data.filter(WorkTime => WorkTime.TimecardHours>14)

// console.log(filterebyDays);
const employees = []
for(i=0; i<employeeData.length; i++){
    
    const Time = employeeData[i]['Time'];
  const TimeOut = employeeData[i]['Time Out'];
      const PositionID =   employeeData[i]['Position ID'];
    const PositionStatus =  employeeData[i]['Position Status'];
    const TimeCardHours =  employeeData[i]['Timecard Hours (as Time)'];
    const PayCycleStartDate =  employeeData[i]['Pay Cycle Start Date'];
    const PayCycleEndDate = employeeData[i]['Pay Cycle End Date'];
    const EmployeeName =  employeeData[i]['Employee Name'];
    const FileNumber =  employeeData[i]['File Number'];
    //employees.push(employee)
    function excelTimeToJSDate(excelTime) {
    
        var millisecondsPerDay = 24 * 60 * 60 * 1000;
        var unixTime = (excelTime - 1) * millisecondsPerDay;
        var jsDate = new Date(unixTime);
    
        return jsDate;
    }

    
    const employee = {Time, TimeOut, PositionID, TimeCardHours, EmployeeName, TimeDate: excelTimeToJSDate(Time), TimeOutDate: excelTimeToJSDate(TimeOut)};

    employees.push(employee)
}


const userWorkedFor7Days = []
const shiftOver14 = []
const userLevelData = _.groupBy(employees, record => {
    return record.EmployeeName
})

const shiftDifference = []


for(let i = 0; i< Object.keys(userLevelData).length; i++) {
//    if(records)
    const records = userLevelData[Object.keys(userLevelData)[i]];
  if(Object.keys(userLevelData)[i] != 'REsaXiaWE, XAis') {
    //continue;
  }
    consecative = 0;
    for(let j = 0 ; j < records.length - 1; j++) {
        //console.log(records[j+1].TimeDate, records[j].TimeDate, moment(moment(records[j+1].TimeDate).startOf('d')).diff(moment(moment(records[j].TimeDate).startOf('d')), 'd'))
     //console.log(moment(records[j+1].TimeDate).diff(moment(records[j].TimeOutDate), 'h'))
        if(moment(records[j+1].TimeDate).diff(moment(records[j].TimeOutDate), 'h') >  1 &&
         moment(records[j+1].TimeDate).diff(moment(records[j].TimeOutDate), 'h') <= 10  ) {
            shiftDifference.push(Object.keys(userLevelData)[i]);
         }
//console.log(records[j].TimeCardHours)
const time = records[j].TimeCardHours.split(":")[0]
console.log(time)  
if(time > 14) {
            shiftOver14.push(Object.keys(userLevelData)[i])
         }

        
        if(moment(moment(records[j+1].TimeDate).startOf('d')).diff(moment(moment(records[j].TimeDate).startOf('d')), 'd') == 1) {
            consecative++
        } else if(moment(moment(records[j+1].TimeDate).startOf('d')).diff(moment(moment(records[j].TimeDate).startOf('d')), 'd') >  1) {
            consecative = 0
        }
    if(consecative == 7) {
        userWorkedFor7Days.push(Object.keys(userLevelData)[i]);
        //break
    }
       
    }

   

  
}
const content = shiftDifference.join('\n')
const content2 = userWorkedFor7Days.join('\n')
const content3 = shiftOver14.join('\n')

const filePath = 'output.txt'

fs.writeFileSync(filePath , content + '\n', 'utf-8')
console.log(`Data has been written to ${filePath}`);
fs.appendFileSync(filePath, content2 + '\n', 'utf-8');
console.log(`Data has been written to ${filePath}`);
fs.appendFileSync(filePath, content3 + '\n', 'utf-8');
console.log(`Data has been written to ${filePath}`);
console.log(shiftDifference)
console.log(userWorkedFor7Days)
 
  console.log(shiftOver14)