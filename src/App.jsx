import { useState } from 'react'
import './App.css'
import * as XLSX from 'xlsx'
import { AddTripDetailFromExcel } from './function/tripdetail';

function App() {
  const [excelFile, setExcelFile]=useState(null);
  const [excelFileError, setExcelFileError]=useState(null); 

  const [excelData, setExcelData]=useState(null);

  const fileType=['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
  const handleFile = (e)=>{
    let selectedFile = e.target.files[0];
    if(selectedFile){
      console.log(selectedFile.type);
      if(selectedFile&&fileType.includes(selectedFile.type)){
        let reader = new FileReader();
        reader.readAsArrayBuffer(selectedFile);
        reader.onload=(e)=>{
          setExcelFileError(null);
          setExcelFile(e.target.result);
        } 
      }
      else{
        setExcelFileError('Please select only excel file types');
        setExcelFile(null);
      }
    }
    else{
      console.log('plz select your file');
    }
  }

  const handleSubmit = async (e) => {
    e.preventDefault();
    try {
      if(excelFile!==null){
        const workbook = XLSX.read(excelFile,{type:'buffer'});
        const worksheetName = workbook.SheetNames[2];
        const worksheet=workbook.Sheets[worksheetName];
  
        let check = ''
        let count = 3
        while (check != undefined) {
          check = worksheet[`D${count + 1}`];
          count = count + 1
        }
  
        let allTripData = []
        
        for (let index = 4; index < 100; index++) {
          // ดึงข้อมูลออกมาจากไฟล์ทีละบรรทัด
          const date = worksheet[`D${index}`];
          const customer = worksheet[`E${index}`];
          const type = worksheet[`I${index}`];
          const serviceType = worksheet[`J${index}`];
          const vehicleType = worksheet[`K${index}`];
          const plateNumber = worksheet[`L${index}`];
          const driverOne = worksheet[`M${index}`];
          const driverTwo = worksheet[`N${index}`];
          const numberOfTrip = worksheet[`Q${index}`];
          const team = worksheet[`R${index}`];
          const network = worksheet[`S${index}`];
          const totalDistance = worksheet[`T${index}`];
          const remark = worksheet[`AB${index}`];
  
          // ถ้าข้อมูลมีการเว้นว่าง ให้เปลี่ยนเป็น null เพื่อป้องกัน Error
          let serviceTypeData
          let vehicleTypeData
          let totalDistanceData
          let remarkData
          let driverOneData
          let driverTwoData
          let numberOfTripData
          let networkData
  
          if (serviceType == undefined) {
            serviceTypeData = null
          } else {
            serviceTypeData = serviceType.v
          }
  
          if (vehicleType == undefined) {
            vehicleTypeData = null
          } else {
            vehicleTypeData = vehicleType.v
          }
  
          if (driverOne == undefined) {
            driverOneData = null
          } else {
            driverOneData = driverOne.v
          }
  
          if (driverTwo == undefined) {
            driverTwoData = null
          } else {
            driverTwoData = driverTwo.v
          }
  
          if (totalDistance == undefined) {
            totalDistanceData = null
          } else {
            totalDistanceData = totalDistance.v
          }
  
          if (remark == undefined) {
            remarkData = null
          } else {
            remarkData = remark.v
          }
  
          if (numberOfTrip == undefined) {
            numberOfTripData = 1
          } else {
            if (numberOfTrip.v > 1) {
              numberOfTripData = numberOfTrip.v
            } else {
              numberOfTripData = 1
            }
          }
  
          if (network == undefined) {
            networkData = team.v.slice(3)
          } else {
            networkData = network.v
          }
  
          console.log({
            date: date.w,
            customer: customer.v,
            type: type.v,
            serviceType: serviceTypeData,
            vehicleType: vehicleTypeData,
            plateNumber: plateNumber.v,
            team: team.v,
            network: networkData,
            numberOfTrip: numberOfTripData,
            totalDistance: totalDistanceData,
            remark: remarkData,
            driverOne: driverOneData,
            driverTwo: driverTwoData,
          });
  
          let tripData = ({
            date: date.w,
            customer: customer.v,
            type: type.v,
            serviceType: serviceTypeData,
            vehicleType: vehicleTypeData,
            plateNumber: plateNumber.v,
            team: team.v,
            network: networkData,
            numberOfTrip: numberOfTripData,
            totalDistance: totalDistanceData,
            remark: remarkData,
            driverOne: driverOneData,
            driverTwo: driverTwoData,
            createBy: 'test',
            // updateBy: user.user.name
          });
  
          allTripData.push(tripData)
          console.log(allTripData);
        }
  
        AddTripDetailFromExcel(allTripData)
          .then((res) => {
            let alertMsg = res.data
            if (alertMsg[0] == 'เ') {
              // console.log(alertMsg)
              alert(alertMsg)
            } else {
              let alertStr = ''
              alertMsg.map((msg) => {
                alertStr = alertStr + msg
                //console.log(msg);
              })
              // console.log(alertStr);
              alert(alertStr)
            }
          })
          .catch((err) => {
            console.log(err)
          })
        // console.log('Add TripDetail Success');
      }
      else{
        setExcelData(null);
      }
    } catch (error) {
      if (error = "TypeError: Cannot read properties of undefined (reading 'v')") {
        console.log('เลือก Page ไม่ถูกต้องหรือ Column ไม่ตรงกับที่ระบบกำหนดไว้');
        alert('เลือก Page ไม่ถูกต้องหรือ Column ไม่ตรงกับที่ระบบกำหนดไว้')
      } else {
        alert(error)
      }
    }
  }

  console.log(excelFile);
  console.log(excelData);

  return (
    <div className="container p-5">
      {/* upload file section */}
      <div className='form'>
        <form className='form-group' autoComplete="off" onSubmit={handleSubmit}>
          <label><h5>Upload Excel file</h5></label>
          <br></br>
          <input type='file' className='form-control' onChange={handleFile} required></input>        
          <button type='submit' className='btn btn-success' style={{marginTop:5+'px'}}>Submit</button>          
        </form>
      </div>
    </div>
  )
}

export default App
