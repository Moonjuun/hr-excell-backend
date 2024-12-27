const express = require("express");
const cors = require("cors");
const fileUpload = require("express-fileupload");
const xlsx = require("xlsx");
const mysql = require("mysql2/promise");
require("dotenv").config(); // Make sure this is at the top of your file

const app = express();

// Middleware 설정
app.use(cors());
app.use(fileUpload());

// MySQL 연결 설정
const db = mysql.createPool({
  host: process.env.DB_HOST,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_DATABASE,
  port: process.env.DB_PORT,
});

// 서버 실행
app.listen(4000, () => {
  console.log("Server is running on http://localhost:4000");
});

app.post("/api/excell/data", async (req, res) => {
  try {
    // 업로드된 파일 가져오기
    if (!req.files || Object.keys(req.files).length === 0) {
      return res.status(400).send("파일이 없습니다.");
    }

    const file = req.files.file;

    // 엑셀 파일을 워크북으로 파싱
    const workbook = xlsx.read(file.data, { type: "buffer" });

    // 첫 번째 시트의 데이터를 가져옵니다.
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // 시트 데이터를 JSON으로 변환
    const jsonData = xlsx.utils.sheet_to_json(sheet);

    // MySQL에 저장할 데이터 준비
    const insertPromises = jsonData.map(async (row) => {
      const {
        이름,
        지원부서,
        나이,
        성별,
        전화번호,
        이메일,
        학력,
        경력,
        입사희망일,
        자격증,
        현재상태,
        희망연봉,
        면접점수,
        평가자,
        면접날짜,
      } = row;

      // 성별을 'M', 'F', 'Other'로 변환
      let genderValue = "Other"; // 기본값 설정
      if (성별 === "남") {
        genderValue = "M";
      } else if (성별 === "여") {
        genderValue = "F";
      }

      try {
        // MySQL에 데이터 삽입
        await db.query(
          `INSERT INTO applicants (
              name, department, age, gender, phone_number, email, education, experience,
              desired_join_date, certification, current_status, desired_salary, interview_score,
              evaluator, interview_date
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
          [
            이름,
            지원부서,
            나이,
            genderValue, // 변환된 성별 값 사용
            전화번호,
            이메일,
            학력,
            경력,
            입사희망일,
            자격증,
            현재상태,
            희망연봉,
            면접점수,
            평가자,
            면접날짜,
          ]
        );
      } catch (error) {
        console.error("MySQL에 삽입 중 오류 발생:", error);
      }
    });
    await Promise.all(insertPromises);
    console.log("데이터 Insert 완료");
    res.send("엑셀 데이터가 성공적으로 저장되었습니다.");
  } catch (error) {
    console.error("Error processing file:", error);
    res.status(500).send("파일 처리 중 오류가 발생했습니다.");
  }
});
