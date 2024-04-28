package main

import (
	//"bytes"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"io"

	"log"
	"net/http"
	"net/smtp"
	"os"
	"strings"

	//"golang.org/x/text/encoding/simplifiedchinese"
	//"golang.org/x/text/transform"

	"gorm.io/driver/postgres"
	"gorm.io/gorm"

	"github.com/xuri/excelize/v2"
)

type LeaveRequest struct {
	LeaveRequestID string `gorm:"column:lr_id;type:uuid;default:uuid_generate_v4();primaryKey"`
	StudentID      string `gorm:"column:student_id"`
	StudentName    string `gorm:"column:s_name"`
	StudentEmail   string `gorm:"column:s_email"`
	StudentDept    string `gorm:"column:s_dept"`
	CourseName     string `gorm:"column:course_name"`
	// 其他需要存儲的欄位
	ClassRoom    string `gorm:"column:class_room"`
	ClassTime    string `gorm:"column:class_time"`
	ClassCampus  string `gorm:"column:class_campus"`
	TeacherEmail string `gorm:"column:t_email"`
}

var db *gorm.DB

func initDB() error {
	var err error
	dsn := "host=localhost user=postgres password=Kk918523 dbname=bird2024 port=5432 sslmode=disable TimeZone=Asia/Shanghai"
	db, err = gorm.Open(postgres.Open(dsn), &gorm.Config{})
	if err != nil {
		return err
	}
	return nil
}

func main() {
	err := initDB()
	if err != nil {
		log.Fatal("Failed to connect to database:", err)
	}

	// 註冊路由處理器
	http.HandleFunc("/importExcel", importExcelHandler)
	http.HandleFunc("/sendEmail", sendEmailHandler)

	// 監聽端口
	fmt.Println("Server started on port 8080")
	log.Fatal(http.ListenAndServe(":8080", nil))
}

// 處理匯入 Excel 請求
func importExcelHandler(w http.ResponseWriter, r *http.Request) {
	// 解析 form 表單
	err := r.ParseMultipartForm(10 << 20) // 設置最大檔案大小為 10MB
	if err != nil {
		log.Println("Error parsing form:", err)
		http.Error(w, "Internal server error", http.StatusInternalServerError)
		return
	}

	// 從 form 表單中獲取檔案
	file, handler, err := r.FormFile("file") // 表單中檔案欄位的名稱為 "file"
	if err != nil {
		log.Println("Error retrieving file from form:", err)
		http.Error(w, "Bad request", http.StatusBadRequest)
		return
	}
	defer file.Close()

	// 指定 Excel 檔案的暫存位置
	excelFilePath := "./excelfiles/uploaded.xlsx"

	// 將檔案寫入暫存位置
	excelFile, err := os.Create(excelFilePath)
	if err != nil {
		log.Println("Error creating temporary Excel file:", err)
		http.Error(w, "Internal server error", http.StatusInternalServerError)
		return
	}
	defer excelFile.Close()

	// 將上傳的 Excel 檔案內容寫入暫存檔案
	_, err = io.Copy(excelFile, file)
	if err != nil {
		log.Println("Error writing Excel file:", err)
		http.Error(w, "Internal server error", http.StatusInternalServerError)
		return
	}

	// 檢查檔案類型是否為 Excel
	contentType := handler.Header.Get("Content-Type")
	if !strings.Contains(contentType, "spreadsheet") {
		log.Println("Error: Not an Excel file")
		http.Error(w, "Unsupported media type", http.StatusUnsupportedMediaType)
		return
	}

	// 指定 CSV 檔案的存儲位置
	csvFilePath := "./csvfiles/uploaded.csv"

	// 將 Excel 檔案轉換為 UTF-8 編碼的 CSV 檔案
	if err := convertExcelToCSV(excelFilePath, csvFilePath); err != nil {
		log.Println("Error converting Excel to CSV:", err)
		http.Error(w, "Internal server error", http.StatusInternalServerError)
		return
	}

	// 解析 CSV 檔案並存入資料庫
	err = parseCSVAndSaveToDB(csvFilePath)
	if err != nil {
		log.Println("Error parsing CSV and saving to database:", err)
		http.Error(w, "Internal server error", http.StatusInternalServerError)
		return
	}

	// 返回成功的回應
	response := map[string]interface{}{
		"status":  "success",
		"message": "Excel imported successfully",
	}
	json.NewEncoder(w).Encode(response)
}

func convertExcelToCSV(excelFilePath string, csvFilePath string) error {
	// 讀取 Excel 檔案
	xlsx, err := excelize.OpenFile(excelFilePath)
	if err != nil {
		return fmt.Errorf("error opening Excel file: %v", err)
	}

	// 獲取所有工作表名稱
	sheetMap := xlsx.GetSheetMap()

	// 檢查是否有工作表
	if len(sheetMap) == 0 {
		return fmt.Errorf("no sheets found in Excel file")
	}

	// 選擇第一個工作表
	sheetName := sheetMap[1]

	rows, err := xlsx.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("error getting rows from Excel file: %v", err)
	}

	// 開啟指定的 CSV 檔案來寫入轉換後的內容
	csvFile, err := os.Create(csvFilePath)
	if err != nil {
		return fmt.Errorf("error creating CSV file: %v", err)
	}
	defer csvFile.Close()

	// 將每一列寫入 CSV 檔案
	writer := csv.NewWriter(csvFile)
	defer writer.Flush()

	for _, row := range rows {
		if err := writer.Write(row); err != nil {
			return fmt.Errorf("error writing CSV file: %v", err)
		}
	}

	return nil
}

func parseCSVAndSaveToDB(csvFilePath string) error {
	csvFile, err := os.Open(csvFilePath)
	if err != nil {
		return err
	}
	defer csvFile.Close()

	// 創建 CSV Reader，並設置逗號作為字段分隔符號
	reader := csv.NewReader(csvFile)
	reader.Comma = ','
	reader.LazyQuotes = true

	// 讀取 CSV 的標題行，獲取每個欄位的索引
	header, err := reader.Read()
	if err != nil {
		return err
	}

	fmt.Println("CSV Header:", header)

	// 確保 CSV 標題行中包含了必要的欄位
	requiredFields := []string{"學號", "姓名", "學生信箱", "學生班級", "科目", "上課校區", "上課教室", "上課時間", "授課教師信箱"}

	for _, field := range requiredFields {
		if !contains(header, field) {
			return fmt.Errorf("CSV file is missing required field: %s", field)
		}
	}

	for {
		record, err := reader.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			// 錯誤處理，可以跳過該行繼續解析下一行
			fmt.Println("Error reading record:", err)
			continue
		}

		// 使用欄位名稱來存取每個欄位的值
		studentID := record[getFieldIndex(header, "學號")]
		studentName := record[getFieldIndex(header, "姓名")]
		studentEmail := record[getFieldIndex(header, "學生信箱")]
		studentDept := record[getFieldIndex(header, "學生班級")]
		courseName := record[getFieldIndex(header, "科目")]
		classcampus := record[getFieldIndex(header, "上課校區")]
		classroom := record[getFieldIndex(header, "上課教室")]
		classtime := record[getFieldIndex(header, "上課時間")]
		teacheremail := record[getFieldIndex(header, "授課教師信箱")]

		// 將解析到的資料存入資料庫
		leaveRequest := LeaveRequest{
			StudentID:    studentID,
			StudentName:  studentName,
			StudentEmail: studentEmail,
			StudentDept:  studentDept,
			CourseName:   courseName,
			ClassCampus:  classcampus,
			ClassRoom:    classroom,
			ClassTime:    classtime,
			TeacherEmail: teacheremail,
		}
		// 呼叫資料庫操作函式，將資料存入資料庫，這裡假設使用 Gorm
		if err := db.Create(&leaveRequest).Error; err != nil {
			return err
		}
	}

	return nil
}

// helper 函式，用來確定欄位名稱在標題行中的索引位置
func getFieldIndex(fields []string, field string) int {
	for i, f := range fields {
		if f == field {
			return i
		}
	}
	return -1 // 如果欄位名稱不存在，返回 -1
}

// helper 函式，用來檢查切片中是否包含特定元素
func contains(slice []string, str string) bool {
	for _, s := range slice {
		if s == str {
			return true
		}
	}
	return false
}

// 處理發送郵件請求
func sendEmailHandler(w http.ResponseWriter, r *http.Request) {
	// 解析請求參數
	teacherEmail := r.FormValue("teacher_email") // 老師的信箱

	// 在資料庫中查找符合老師信箱的學生請假資訊
	leaveRequests, err := findLeaveRequestsByTeacherEmail(teacherEmail)
	if err != nil {
		log.Println("Error finding leave requests:", err)
		http.Error(w, "Internal server error", http.StatusInternalServerError)
		return
	}

	// 動態生成郵件內容
	emailBody := generateEmailBody(leaveRequests)

	// 發送郵件
	err = sendEmail(teacherEmail, "學生請假名單", emailBody)

	if err != nil {
		log.Println("Error sending email:", err)
		http.Error(w, "Failed to send email", http.StatusInternalServerError)
		return
	}

	// 返回成功的回應
	response := map[string]interface{}{
		"status":  "success",
		"message": "Email sent successfully",
	}
	json.NewEncoder(w).Encode(response)
}

// 動態生成郵件內容
func generateEmailBody(leaveRequests []LeaveRequest) string {
	emailBody := "老師您好,\n\n以下為您目前開課的學生請假名單，請查閱:\n"

	// 使用指定格式列出學生資訊
	for _, request := range leaveRequests {
		// 使用管道符號分隔不同欄位
		studentInfo := fmt.Sprintf("·學號:%s\n·姓名:%s\n·學生信箱:%s\n·學生系所:%s\n·課程名稱:%s\n·上課時間:%s\n·上課校區:%s\n·上課教室:%s\n", request.StudentID, request.StudentName, request.StudentEmail, request.StudentDept, request.CourseName, request.ClassTime, request.ClassCampus, request.ClassRoom)
		emailBody += studentInfo + "\n"
	}

	emailBody += "如有問題，煩請老師寫信告知詢問!\n\n感謝您,\n送信鳥擇像發信系統團隊"
	return emailBody
}

// 在資料庫中查找符合老師信箱的學生請假資訊
func findLeaveRequestsByTeacherEmail(teacherEmail string) ([]LeaveRequest, error) {
	var leaveRequests []LeaveRequest
	// 使用 GORM 查詢符合條件的學生請假資訊
	if err := db.Where("t_email = ?", teacherEmail).Find(&leaveRequests).Error; err != nil {
		return nil, err
	}
	return leaveRequests, nil
}

// 發送郵件的函式
func sendEmail(recipient string, subject string, body string) error {
	from := "skes1114@gmail.com"   // 你的 Gmail 信箱
	password := "ckczffhmujrpwlzz" // 你的 Gmail 密碼

	// SMTP 設定
	smtpHost := "smtp.gmail.com"
	smtpPort := "587"

	// 設定身份驗證信息
	auth := smtp.PlainAuth("", from, password, smtpHost)

	// 郵件內容
	message := []byte("To: " + recipient + "\r\n" +
		"Subject: " + subject + "\r\n" + // 使用動態設定的主題
		"\r\n" +
		body + "\r\n")

	// 發送郵件
	err := smtp.SendMail(smtpHost+":"+smtpPort, auth, from, []string{recipient}, message)
	if err != nil {
		log.Println("Error sending email:", err)
		return err
	}
	fmt.Println("Email sent to", recipient)
	return nil
}
