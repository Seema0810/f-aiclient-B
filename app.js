import express from 'express';
import cors from 'cors';
import cookieParser from 'cookie-parser';
import morgan from 'morgan'
import userRouter from './routes/user.routes.js';
import companyRoutes from './routes/user.company.routes.js';
import reportRoutes from './routes/report.routes.js';
import fileRoutes from './routes/fileRoutes.js';
import xlsx from "xlsx";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const app = express();
app.use(morgan('dev'));
app.use(
    cors({
      // origin: [],
      origin: ["http://localhost:3000","https://f-ai-424219.el.r.appspot.com","https://f-aiclient-pf5ho7tiaq-el.a.run.app"],

      credentials: true
    })
  );
  
app.get('/', (req,res)=>{
    res.send('Hello There from my deployed Server')
})
app.use(express.json());
app.use(cookieParser());
app.use(express.urlencoded({ extended: true, limit: '32kb' }));
app.use(express.static('public'));


app.use('/api/user', userRouter);
app.use('/api/companies', companyRoutes);
app.use('/api/files', fileRoutes);
app.use('/api/user/report', reportRoutes);


const __dirname = dirname(fileURLToPath(import.meta.url));

app.use(express.static(join(__dirname, "./build")));
app.get("/dashboard", (req, res) => {
  res.sendFile(join(__dirname, "./build/index.html"));
});


app.post("/updateCell", (req, res) => {
  const { newValue } = req.body;
  // console.log(newValue)
  if (!newValue) {
    return res.status(400).send("No new value provided.");
  }

  try {
    const workbook = xlsx.readFile("personalFinnace.xlsx");
    const worksheet = workbook.Sheets["Sheet1"];

    if (!worksheet) {
      console.error("Worksheet not found.");
      return res.status(404).send("Worksheet not found");
    }

    const cellRef = "B7";
    if (!worksheet[cellRef]) {
      worksheet[cellRef] = {};
    }
    worksheet[cellRef].v = newValue;

    xlsx.writeFile(workbook, "Updated_PersonalFinnace.xlsx");
    res.download("Updated_personalFinnace.xlsx", "personalFinnace.xlsx");

    res.status(200).send("Cell updated successfully");
  } catch (error) {
    console.error("Failed to update cell:", error);
    res.status(500).send("Failed to update the cell");
  }
});

app.get("/workbookToJson", (req, res) => {
  try {
    const workbook = xlsx.readFile("Updated_PersonalFinnace.xlsx");

    const sheetNames = workbook.SheetNames;

    const jsonData = {};

    sheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];

      const sheetData = xlsx.utils.sheet_to_json(worksheet);

      jsonData[sheetName] = sheetData;
    });

    res.send(jsonData);
  } catch (error) {
    console.error("Failed to convert workbook to JSON:", error);
    res.status(500).send("Failed to convert workbook to JSON");
  }
});
app.post("/updateMarketingCell", (req, res) => {
  const { newMarketValue } = req.body;
  console.log(newMarketValue)
  if (!newMarketValue) {
    return res.status(400).send("No new value provided.");
  }

  try {
    const workbook = xlsx.readFile("Marketing.xlsx");
    const worksheet = workbook.Sheets["DASHBOARD"];

    if (!worksheet) {
      console.error("Worksheet not found.");
      return res.status(404).send("Worksheet not found");
    }

    const cellRef = "P15";
    if (!worksheet[cellRef]) {
      worksheet[cellRef] = {};
    }
    worksheet[cellRef].v = newMarketValue;

    xlsx.writeFile(workbook, "Updated_Marketing.xlsx");
    res.download("Updated_Marketing.xlsx", "Marketing.xlsx");

    res.status(200).send("Cell updated successfully");
  } catch (error) {
    console.error("Failed to update cell:", error);
    res.status(500).send("Failed to update the cell");
  }
});
app.get("/MarketingworkbookToJson", (req, res) => {
  try {
    const workbook = xlsx.readFile("Updated_Marketing.xlsx");

    const sheetNames = workbook.SheetNames;

    const jsonData = {};

    sheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];

      const sheetData = xlsx.utils.sheet_to_json(worksheet);

      jsonData[sheetName] = sheetData;
    });

    res.send(jsonData);
  } catch (error) {
    console.error("Failed to convert workbook to JSON:", error);
    res.status(500).send("Failed to convert workbook to JSON");
  }
});

app.post("/updateSalesData", (req, res) => {
  const { newSalesValue } = req.body;
  console.log(newSalesValue)
  if (!newSalesValue) {
    return res.status(400).send("No new value provided.");
  }

  try {
    const workbook = xlsx.readFile("Sales.xlsx");
    const worksheet = workbook.Sheets["DASHBOARD"];

    if (!worksheet) {
      console.error("Worksheet not found.");
      return res.status(404).send("Worksheet not found");
    }

    const cellRef = "R12";
    if (!worksheet[cellRef]) {
      worksheet[cellRef] = {};
    }
    worksheet[cellRef].v = newSalesValue;

    xlsx.writeFile(workbook, "Updated_Sales.xlsx");
    res.download("Updated_Sales.xlsx", "Sales.xlsx");

    res.status(200).send("Cell updated successfully");
  } catch (error) {
    console.error("Failed to update cell:", error);
    res.status(500).send("Failed to update the cell");
  }
});

app.get("/SalesworkbookToJson", (req, res) => {
  try {
    const workbook = xlsx.readFile("Updated_Sales.xlsx");

    const sheetNames = workbook.SheetNames;

    const jsonData = {};

    sheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];

      const sheetData = xlsx.utils.sheet_to_json(worksheet);

      jsonData[sheetName] = sheetData;
    });

    res.send(jsonData);
  } catch (error) {
    console.error("Failed to convert workbook to JSON:", error);
    res.status(500).send("Failed to convert workbook to JSON");
  }
});







app.post("/EmployeePerformance", (req, res) => {
  const { newPerformance } = req.body;
  console.log(newPerformance)
  if (!newPerformance) {
    return res.status(400).send("No new value provided.");
  }

  try {
    const workbook = xlsx.readFile("Employe.xlsx");
    const worksheet = workbook.Sheets["Sheet1"];

    if (!worksheet) {
      console.error("Worksheet not found.");
      return res.status(404).send("Worksheet not found");
    }

    const cellRef = "D10";
    if (!worksheet[cellRef]) {
      worksheet[cellRef] = {};
    }
    worksheet[cellRef].v = newPerformance;

    xlsx.writeFile(workbook, "Updated_Employe.xlsx");
    res.download("Updated_Employe.xlsx", "updatedEmployee.xlsx");

    res.status(200).send("Cell updated successfully");
  } catch (error) {
    console.error("Failed to update cell:", error);
    res.status(500).send("Failed to update the cell");
  }
});

app.get("/EmployeJson", (req, res) => {
  try {
    const workbook = xlsx.readFile("Updated_Employe.xlsx");

    const sheetNames = workbook.SheetNames;

    const jsonData = {};

    sheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];

      const sheetData = xlsx.utils.sheet_to_json(worksheet);

      jsonData[sheetName] = sheetData;
    });

    res.send(jsonData);
  } catch (error) {
    console.error("Failed to convert workbook to JSON:", error);
    res.status(500).send("Failed to convert workbook to JSON");
  }
});


export default app;