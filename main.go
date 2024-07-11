package main

import (
	"bytes"
	"database/sql"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"os"

	_ "github.com/lib/pq"
	"github.com/tealeg/xlsx"
	"gopkg.in/gomail.v2"
)

type Query struct {
	AttachmentName string `json:"attachmentName"`
	Query          string `json:"query"`
}
type Config struct {
	DatabaseURI  string `json:"databaseURI"`
	FromEmail    string `json:"fromEmail"`
	ToEmail      string `json:"toEmail"`
	EmailSubject string `json:"emailSubject"`
	Password     string `json:"password"`
	EmailBody    string `json:"emailBody"`
	SMTPHost     string `json:"smtpHost"`
	SMTPPort     int    `json:"smtpPort"`
	Queries      []Query
}

type FileAttachment struct {
	AttachmentName string
	Content        []byte
}

func main() {

	logFile, err := os.OpenFile("app.log", os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	defer logFile.Close()
	log.SetOutput(logFile)

	log.Println("Script invoked...")

	configFile, err := os.Open("config.json")
	if err != nil {
		log.Fatalf("Error opening config file: %v", err)
	}

	var config Config

	if err := json.NewDecoder(configFile).Decode(&config); err != nil {
		log.Fatalf("failed to decode config.json: %v", err)
	}

	db, err := sql.Open("postgres", config.DatabaseURI)
	if err != nil {
		panic(err)
	}

	defer db.Close()

	// rows, err := db.Query(config.Queries[0].Query)
	// CheckError(err)

	// defer rows.Close()

	for _, query := range config.Queries {
		attachment := FileAttachment{AttachmentName: query.AttachmentName, Content: getAttachments(db, &query)}
		sendEmail(&config, &attachment)
	}

}

func sendEmail(config *Config, attachment *FileAttachment) {
	m := gomail.NewMessage()
	m.SetHeader("From", config.FromEmail)
	m.SetHeader("To", config.ToEmail)
	m.SetHeader("Subject", config.EmailSubject)
	m.SetBody("text/plain", config.EmailBody)
	// Attach the Excel file from the buffer
	m.Attach(attachment.AttachmentName, gomail.SetCopyFunc(func(w io.Writer) error {
		_, err := w.Write(attachment.Content)
		return err
	}))
	// Send the email
	d := gomail.NewDialer(config.SMTPHost, config.SMTPPort, config.FromEmail, config.Password)
	if err := d.DialAndSend(m); err != nil {
		log.Fatal(err)
	}
	log.Println("Email sent successfully with the Excel file.")
}

func getAttachments(db *sql.DB, query *Query) []byte {
	rows, err := db.Query(query.Query)
	CheckError(err)
	defer rows.Close()

	// Get column names
	cols, err := rows.Columns()
	if err != nil {
		log.Fatal(err)
	}
	values := make([]interface{}, len(cols))
	valuePtrs := make([]interface{}, len(cols))
	for i := range values {
		valuePtrs[i] = &values[i]
	}

	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		log.Fatal(err)
	}

	headerRow := sheet.AddRow()
	for _, colName := range cols {
		cell = headerRow.AddCell()
		cell.Value = colName
	}

	for rows.Next() {
		values := make([]interface{}, len(cols))
		valuePtrs := make([]interface{}, len(cols))
		for i := range values {
			valuePtrs[i] = &values[i]
		}
		if err := rows.Scan(valuePtrs...); err != nil {
			log.Fatal(err)
		}
		row = sheet.AddRow()

		for _, col := range values {
			cell := row.AddCell()
			// Convert all types of values to strings
			cell.Value = fmt.Sprintf("%v", col)
		}
	}

	var buf bytes.Buffer
	if err = file.Write(&buf); err != nil {
		log.Fatal(err)
	}

	return buf.Bytes()
}

func CheckError(err error) {
	if err != nil {
		log.Fatalf("Error: %v", err)
		panic(err)
	}
}
