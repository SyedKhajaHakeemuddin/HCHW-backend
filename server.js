const express = require("express");
const xlsx = require("xlsx");
const cors = require("cors");
const bodyParser = require("body-parser");
const fs = require("fs");
const path = require("path");
const multer = require("multer");

const app = express();
const PORT = 5000;

// Middleware
app.use(cors());
app.use(bodyParser.json());
app.use("/uploads", express.static("uploads")); // Serve uploaded files
app.use("/uploads/photos", express.static("uploads/photos")); // Serve student photos

// Paths
const excelFilePath = path.join(__dirname, "uploads", "student_data.xlsx");
const photosDir = path.join(__dirname, "uploads", "photos");
const placeholderImage = "http://localhost:5000/uploads/photos/placeholder.jpg"; // Default image

// ✅ Ensure directories exist
if (!fs.existsSync("uploads")) fs.mkdirSync("uploads", { recursive: true });
if (!fs.existsSync(photosDir)) fs.mkdirSync(photosDir, { recursive: true });

// ✅ Function to read data from Excel
const readExcelData = () => {
  if (!fs.existsSync(excelFilePath)) {
    throw new Error("Excel file not found. Please upload a file.");
  }
  const workbook = xlsx.readFile(excelFilePath);
  const allData = {};

  workbook.SheetNames.forEach((sheetName) => {
    const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
    allData[sheetName] = sheetData;
  });

  return allData;
};

// ✅ Normalize column names for consistency
const normalizeColumns = (entry) => {
  const normalizedEntry = {};
  Object.keys(entry).forEach((key) => {
    normalizedEntry[key.toLowerCase()] = entry[key];
  });
  return normalizedEntry;
};

// ✅ Endpoint to search student by admission number
app.get("/search/admission", (req, res) => {
  try {
    const { admissionNumber } = req.query;
    if (!admissionNumber) {
      return res.status(400).json({ error: "Missing admissionNumber" });
    }

    const allData = readExcelData();
    let foundStudent = null;

    Object.keys(allData).forEach((classname) => {
      if (!foundStudent) {
        const classData = allData[classname].map(normalizeColumns);
        foundStudent = classData.find((entry) => entry["admin_no"] == admissionNumber);
        if (foundStudent) {
          foundStudent.classname = classname;
        }
      }
    });

    if (!foundStudent) {
      return res.status(404).json({ error: "Student not found" });
    }

    // ✅ Fetch photo from /uploads/photos/
    const photoPath = path.join(photosDir, `${foundStudent.admin_no}.jpg`);
    foundStudent.photo = fs.existsSync(photoPath)
      ? `http://localhost:${PORT}/uploads/photos/${foundStudent.admin_no}.jpg`
      : placeholderImage;

    res.json(foundStudent);
  } catch (error) {
    console.error("Error in /search/admission:", error.message);
    res.status(500).json({ error: "Internal Server Error" });
  }
});

// ✅ Endpoint to search student by name with live suggestions
app.get("/search/name", (req, res) => {
  try {
    const { name } = req.query;
    if (!name) {
      return res.status(400).json({ error: "Missing name parameter" });
    }

    const allData = readExcelData();
    let suggestions = [];

    Object.keys(allData).forEach((classname) => {
      const classData = allData[classname].map(normalizeColumns);
      const matches = classData.filter((entry) => {
        const studentName = entry["student_name"];
        return studentName && studentName.toLowerCase().includes(name.toLowerCase());
      });

      // ✅ Include admission number for fetching full student details
      suggestions = [
        ...suggestions,
        ...matches.map((student) => ({
          student_name: student.student_name,
          admin_no: student.admin_no,
        })),
      ];
    });

    res.json(suggestions.length > 0 ? suggestions : []);
  } catch (error) {
    console.error("Error in /search/name:", error.message);
    res.status(500).json({ error: "Internal Server Error" });
  }
});
// ✅ Download Student Data (Excel)
app.get("/download-student-data", (req, res) => {
  try {
    if (fs.existsSync(excelFilePath)) {
      res.download(excelFilePath, "student_data.xlsx");
    } else {
      res.status(404).json({ error: "Student data file not found." });
    }
  } catch (error) {
    console.error("❌ Error in /download-student-data:", error.message);
    res.status(500).json({ error: "Internal Server Error" });
  }
});


// ✅ Upload Student Data (Excel)
const storage = multer.diskStorage({
  destination: "uploads/",
  filename: (req, file, cb) => {
    cb(null, "student_data.xlsx"); // Always save as student_data.xlsx
  },
});
const upload = multer({ storage: storage, fileFilter: (req, file, cb) => {
    if (path.extname(file.originalname) !== ".xlsx") {
      return cb(new Error("Only .xlsx files are allowed!"), false);
    }
    cb(null, true);
  },
});

app.post("/upload", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "Please upload an Excel file." });
    }
    res.json({ message: "✅ File uploaded successfully!" });
  } catch (error) {
    console.error("Error in /upload:", error.message);
    res.status(500).json({ error: "Internal Server Error" });
  }
});


// ✅ Endpoint to delete student data (Excel)
// ✅ Delete Student Data (Excel)
app.delete("/delete", (req, res) => {
  try {
    if (fs.existsSync(excelFilePath)) {
      fs.unlinkSync(excelFilePath);
      res.json({ message: "✅ Excel file deleted successfully!" });
    } else {
      res.status(404).json({ error: "No file to delete." });
    }
  } catch (error) {
    console.error("Error in /delete:", error.message);
    res.status(500).json({ error: "Internal Server Error" });
  }
});


// ✅ School Activities Storage in JSON
const activitiesFilePath = path.join(__dirname, "uploads", "school_activities.json");

const readActivities = () => {
  if (!fs.existsSync(activitiesFilePath)) {
    fs.writeFileSync(activitiesFilePath, JSON.stringify([]));
  }
  return JSON.parse(fs.readFileSync(activitiesFilePath, "utf-8"));
};

const writeActivities = (data) => {
  fs.writeFileSync(activitiesFilePath, JSON.stringify(data, null, 2));
};

app.post("/activities", (req, res) => {
  try {
    writeActivities(req.body);
    res.json({ message: "Activities saved successfully!" });
  } catch (error) {
    console.error("Error saving activities:", error.message);
    res.status(500).json({ error: "Failed to save activities." });
  }
});

app.get("/activities", (req, res) => {
  try {
    res.json(readActivities());
  } catch (error) {
    console.error("Error fetching activities:", error.message);
    res.status(500).json({ error: "Failed to fetch activities." });
  }
});

// ✅ Upload Multiple Photos (Folder Upload)
const uploadPhotos = multer({ dest: "uploads/photos/" });

app.post("/upload-photos", uploadPhotos.array("photos"), (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: "No photos uploaded." });
    }

    req.files.forEach((file) => {
      const ext = path.extname(file.originalname).toLowerCase();
      if (ext === ".jpg" || ext === ".png") {
        const newFilePath = path.join(photosDir, file.originalname);
        fs.renameSync(file.path, newFilePath);
      } else {
        fs.unlinkSync(file.path); // Delete invalid file
      }
    });

    res.json({ message: "✅ Photos uploaded successfully!" });
  } catch (error) {
    console.error("❌ Error uploading photos:", error);
    res.status(500).json({ error: "Failed to upload photos." });
  }
});

// ✅ Delete All Photos
app.delete("/delete-photos", (req, res) => {
  try {
    fs.readdir(photosDir, (err, files) => {
      if (err) {
        return res.status(500).json({ error: "Error reading photos directory." });
      }
      files.forEach((file) => fs.unlinkSync(path.join(photosDir, file)));
      res.json({ message: "✅ All photos deleted successfully!" });
    });
  } catch (error) {
    console.error("❌ Error deleting photos:", error);
    res.status(500).json({ error: "Failed to delete photos." });
  }
});


// ✅ Start server
app.listen(PORT, () => {
  console.log(`✅ Server running on http://localhost:${PORT}`);
});
