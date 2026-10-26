package main

import (
	"context"
	"database/sql"
	"encoding/json"
	"fmt"
	"net"
	"net/url"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/joho/godotenv"
	_ "github.com/microsoft/go-mssqldb"
)

type job struct {
	JobID          string    `json:"job_id"`
	Email          string    `json:"email"`
	TotalJob       int       `json:"total_job"`
	ReceiveStatus  string    `json:"job_status_receive"`
	EmailStatus    string    `json:"job_status_send_email"`
	CreatedAt      time.Time `json:"created_at"`
}

func main() {
	if len(os.Args) != 2 {
		panic("env path required")
	}
	values, err := godotenv.Read(os.Args[1])
	if err != nil {
		panic(err)
	}
	server := strings.TrimSpace(values["DB_SERVER"])
	user := strings.TrimSpace(values["DB_USER"])
	password := values["DB_PASSWORD"]
	database := strings.TrimSpace(values["DB_NAME"])
	if database == "" { database = "guest_wifi" }
	if server == "" || user == "" || password == "" { panic("database configuration incomplete") }

	dsnURL := &url.URL{Scheme: "sqlserver", User: url.UserPassword(user, password)}
	switch {
	case strings.Contains(server, ","):
		parts := strings.SplitN(server, ",", 2)
		port, parseErr := strconv.Atoi(strings.TrimSpace(parts[1]))
		if parseErr != nil { panic(parseErr) }
		dsnURL.Host = net.JoinHostPort(strings.TrimSpace(parts[0]), strconv.Itoa(port))
	case strings.Contains(server, `\`):
		parts := strings.SplitN(server, `\`, 2)
		dsnURL.Host = strings.TrimSpace(parts[0])
		dsnURL.Path = "/" + strings.TrimSpace(parts[1])
	default:
		dsnURL.Host = net.JoinHostPort(server, "1433")
	}
	query := dsnURL.Query()
	query.Set("database", database)
	query.Set("encrypt", "disable")
	query.Set("TrustServerCertificate", "true")
	dsnURL.RawQuery = query.Encode()

	db, err := sql.Open("sqlserver", dsnURL.String())
	if err != nil { panic(err) }
	defer db.Close()
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Second)
	defer cancel()
	rows, err := db.QueryContext(ctx, `SELECT job_id, email, total_job,
		job_status_receive, job_status_send_email, created_at
		FROM dbo.job_ticket ORDER BY created_at, job_id`)
	if err != nil { panic(err) }
	defer rows.Close()
	jobs := make([]job, 0)
	for rows.Next() {
		var item job
		if err := rows.Scan(&item.JobID, &item.Email, &item.TotalJob, &item.ReceiveStatus, &item.EmailStatus, &item.CreatedAt); err != nil { panic(err) }
		jobs = append(jobs, item)
	}
	if err := rows.Err(); err != nil { panic(err) }
	payload, err := json.Marshal(jobs)
	if err != nil { panic(err) }
	fmt.Print(string(payload))
}
