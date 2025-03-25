import React from "react";
import * as XLSX from "xlsx";
import { useState } from "react";
import { useEffect } from "react";

export const RETAILTIMESHEET = () => {
  const [file, setFile] = useState(null);
  const [json, setJson] = useState([]);
  const [FinalDisplayArray, setFinalDisplayArray] = useState([]);
  const [u, setu] = useState([])
  console.log(json);



  function getUniqueTickets(arr) {
    const ticketMap = new Map();

    arr.forEach(obj => {
      const ticketNumber = obj["Ticket Number"];
      if (!ticketMap.has(ticketNumber)) {
        ticketMap.set(ticketNumber, obj);
      }
    });
    return Array.from(ticketMap.values());

  }

  // Array of names to filter
  const namesToFilter = [
    "Harishwar Nesamani",
    "KNVANSABARISHGUPTA OBILISETTI",
    "Akash Sampath",
    "Naveen Kumar Anandan",
    "Jayasree Mohanakrishnan",
    "Deepika Sampath Kumar",
    "Tarun Akash Pazhani S",
    "Tamilarasi Balamurugan",
    "Jeevanandam Ruthramurthy",
    "Sudha Birendarkumar",
    "Dhuruva Gowshik Ganesan",
    "Karthik Govindasamy",
    "Meenakshi Maragathavel",
    "Veeravisvavinayagam Kumaravelu",
    "Vedhasree Manivannan",
    "Janani Venkatesalu",
    "Priyadharshini Mohan",
    "Moneshwar Devaraj",
    "Mohammedumarmuqthar Mansoor",
    "Divya Shree",
    "Sindhuja Prabakaran",
    "Kishore Ganesan",
    "Nitish Kumar D",
    "Yuvaraj Selvam",
    "Rojini.S Sathish Kumar",
    "Priyadharshini James",
    "Arul Mani Joseph",
    "Anitha Ananthan",
    "Vishwa Alagiri",
    "Kirthika Jayaraman",
    "Jenithson Thommai",
    "Karthikeyan Panachavaranam",
    "Najir Hussain Nashim Miyan",
    "Ayyapparaj Dhamodhaan",
    "Epsi Surendran",
    "Vishnu Bose",
    "Shalini Subramanian",
    "Divya Barani Karthikeyean",
    "Manoj Rajasekaran",
    "Kowsalya G",
    "Pooja Raghavendra",
    "Saranya Selvamani",
    "Sruthi Mathivanan",
    "Goutham Sakthivel",
    "Sneha Hari Doss",
    "Prasanth Rajendran",
    "Ramprakash Rajan",
    "Sandhiya Kollapuri",
    "Dilip Suresh",
    "Kishore Sivalingam",
    "Vidhul Jothi Senthil Nathan",
    "Bhargavi Baskaran",
    "Rangarajan Basker",
    "Tharun Kumar V",
    "Logeshwari S Sundaramoorthy",
    "Rajeshwari Rajagopal",
    "Shantha Kumar Saravanan",
    "Akash N Natarajan C",
    "Ashwin Kumar S",
    "Pradeep Joel Xavier",
    "Ali Mehran Kandrikar",
    "Saranesh Duraisamy",
    "Sivasankari Arumugam",
    "Saran Kumar G",
    "Shifhana Banu Usain",
    "Ranjana Mohan",
    "Rex Fleming",
    "Harshaavardhan Subramani",
    "Rajamouli R",
    "Priya Dharshini K",
    "Siddharthan Mayilsamy",
    "Madhumitha.C Chandhiran.N",
    "Naveen Srinivasan",
    "Sathish Kumar Venkatesan",
    "Keerthana Ganesh",
    "Sathish Kumar Sankaranagappan",
    "Ritesh Suresh",
    "Bhuvan Balasubramanian",
    "Kiranraj Ravichandran",
    "Shanmuga Priya Ramesh",
    "Prabhakaran Sekar",
    "Manoj Thiruppathi",
    "Priyea Dharshani B",
    "Tanya Jackson",
    "Swetha Mani",
    "Durairaj Saravanakumar",
    "Avi Sharma",
    "Saquib Tanweer",
    "Sam Turner",
    "Aarthi Madhan",
    "Lakshmi Aishwarya Ratakondala",
    "Dhruv Doshi",
    "Augustina Albert Sagayaraj",
    "Gnana Wilciya",
    "Palani Raja Vellaisamy",
    "Dhanalakshmi Sundar",
    "Nithish Thivya",
    "Vijayalakshmi Janakiraman",
    "Balaji Ashok Kumar",
    "Harihara Ponnaiah",
    "Kamaleeshwari Sasi Kapoor Singh",
    "Akshay Kumar P",
    "Savitha Panneerselvan",
    "Yuvasree Balasubramaniam",
    "Vijay R Kumar",
    "Naveen Kumar Sankar",
    "Arun Sajeev",
    "Aswini Haribabu",
    "Praveen Kumar Thanigaiarasu",
    "Amrutha Rajan",
    "Pragadeeshwaran Ganesan",
    "Govarthan Mohan",
    "Keerthana j",
    "Sathish E",
    "Mahalakshmi Gopi",
    "Monisha Babu",
    "Devakumar Y"
  ];

  // Object array to filter


  // Filter the array based on the `name` property
  const filteredArray = json.filter(item => namesToFilter.includes(item["Status Modifier"]));

  // Log the result





  const handleConvert = () => {
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: "binary" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);
        setJson(json);
      };
      reader.readAsBinaryString(file);
    }
  };

  const handleExport = () => {
    // Create a worksheet from the data
    const worksheet = XLSX.utils.json_to_sheet(FinalArr);

    // Create a new workbook
    const workbook = XLSX.utils.book_new();

    // Append the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    // Export the workbook as an Excel file
    XLSX.writeFile(workbook, "MyData.xlsx");
  };

  function excelDateToJsDate(excelDate) {
    // Excel's epoch start (January 1, 1900)
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));

    // Convert the serial number to milliseconds
    const msSinceEpoch = excelDate * 24 * 60 * 60 * 1000;

    // Add the milliseconds to the epoch
    const jsDate = new Date(excelEpoch.getTime() + msSinceEpoch);

    // Format the date (e.g., "9-9-2024")
    const day = jsDate.getDate();
    const month = jsDate.getMonth() + 1; // Months are zero-indexed
    const year = jsDate.getFullYear();

    return `${month}-${day}-${year}`;
  }

  function excelDateToJsMonth(excelDate) {
    // Excel's epoch start (January 1, 1900)
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    // Convert the serial number to milliseconds
    const msSinceEpoch = excelDate * 24 * 60 * 60 * 1000;
    // Add the milliseconds to the epoch
    const jsDate = new Date(excelEpoch.getTime() + msSinceEpoch);
    // Month names array
    const monthNames = [
      "Jan",
      "Feb",
      "Mar",
      "Apr",
      "May",
      "Jun",
      "Jul",
      "Aug",
      "Sep",
      "Oct",
      "Nov",
      "Dec",
    ];

    // Extract month and year
    const month = monthNames[jsDate.getMonth()]; // Zero-indexed
    const year = jsDate.getFullYear();

    return `${month}`;
  }

  function timCounter(str) {
    let count = 0;

    // Loop through the string character by character
    for (let i = 0; i < str.length; i++) {
      // Check if the character is a digit (0-9) using a regular expression
      if (/[0-9]/.test(str[i])) {
        count++;
      }
    }

    return Math.round(count / 7);
  }

  function dexCounter(str = "") {
    let count = 0;

    // Loop through the string character by character
    for (let i = 0; i < str.length; i++) {
      // Check if the character is a digit (0-9) using a regular expression
      if (/[0-9]/.test(str[i])) {
        count++;
      }
    }

    return Math.round(count / 6);
  }
  function convertMinutesToHours(minutes) {
    const hours = Math.floor(minutes / 60); // Whole hours
    const remainingMinutes = minutes % 60; // Remaining minutes
    return `${hours} hour(s) and ${remainingMinutes} minute(s)`;
  }

  let TotalTurnAroundTime = 0
  let FinalArr = [];
  getUniqueTickets(filteredArray).map((m) => {
    let currentUniqueArray = []
    let currentTicket = filteredArray.filter((f) => f["Ticket Number"] == m["Ticket Number"])

    console.log(currentTicket);


    let arr = {};
    arr["Date"] = excelDateToJsDate(m["Status Change Time"]);
    arr["Month"] = excelDateToJsMonth(m["Status Change Time"]);
    arr["Name"] = m["Status Modifier"];
    arr["TS Ticket #"] = m["Ticket Number"];
    arr["Ticket Product"] = m["Ticket Type"];
    arr["Tim count"] = timCounter(m["TIM #(s)"]);
    arr["Dex count"] = dexCounter(m["Deal ID(s)"])
    arr["Client name"] = m["Ticket Name"];
    arr["Turn around time"] = currentTicket
      .map(item => item["Time In Old Status"])
      .reduce((acc, curr) => acc + curr, 0) + 1

    arr["back and forth"] = currentTicket.length

    TotalTurnAroundTime += m["Time In Old Status"]

    FinalArr.push(arr)





  })









  const uniqueTickets = Array.from(
    FinalArr
      .reduce((map, item) => {
        if (!map.has(item["TS Ticket #"])) {
          map.set(item["TS Ticket #"], item);
        }
        return map;
      }, new Map())
      .values()
  );


  let TotalTimHandled = 0
  let TotalTicketHandled = uniqueTickets.length
  let AverageTimeForTim = 0


  uniqueTickets.map((m) => {
    TotalTimHandled += m["Tim count"]



  })
  AverageTimeForTim = TotalTurnAroundTime / TotalTimHandled
  let AverageTimeForTicket = TotalTurnAroundTime / TotalTicketHandled






  return (
    <div>
      <h1>names and tickets data</h1>
      <input
        type="file"
        accept=".xls,.xlsx"
        onChange={(e) => setFile(e.target.files[0])}
      />
      <button onClick={handleConvert}>Convert</button>
      <button onClick={handleExport}>Export to Excel</button>
      <p>Total Tim Handled: {TotalTimHandled}</p>
      <p>Total Ticket Handled: {TotalTicketHandled}</p>
      <p>Total Turn around Time: {TotalTurnAroundTime}</p>
      <p>Average Time Taken For Tim: {AverageTimeForTim}</p>
      <p>Average Time Taken For Ticket: {AverageTimeForTicket}</p>
      <table>
        <tr>
          <th>Date</th>
          <th>Month</th>
          <th>Name</th>
          <th>TS Ticket #</th>
          <th>Ticket Product</th>
          <th>TIMs Uploaded</th>
          <th>Dex Orders Uploaded</th>
          <th>Turn around time</th>
          <th>Client Name</th>
        </tr>
        {FinalArr.map((m) => {
          return (
            <tr>
              <td>{m["Date"]}</td>
              <td>{m["Month"]}</td>
              <td>{m["Name"]}</td>
              <td>{m["TS Ticket #"]}</td>
              <td>{m["Ticket Product"]}</td>
              <td>{m["Tim count"]}</td>
              <td>{m["Dex count"]}</td>
              <td>{m["Turn around time"]}</td>
              <td>{m["Client name"]}</td>
            </tr>
          )
        })}
      </table>
    </div>
  );
};
