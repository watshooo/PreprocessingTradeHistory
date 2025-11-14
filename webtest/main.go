package main

import (
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"time"

	"github.com/gorilla/mux"
	"github.com/rs/cors"
)

type Config struct {
	RateSpot   float64 `json:"rate_spot"`
	RateRemote float64 `json:"rate_remote"`
}

type ProcessRequest struct {
	TradeHistoryFiles []string `json:"trade_history_files"`
	JisdorFile        string   `json:"jisdor_file"`
	Config            Config   `json:"config"`
}

type ProcessResponse struct {
	Success    bool     `json:"success"`
	Message    string   `json:"message"`
	OutputFile string   `json:"output_file,omitempty"`
	Error      string   `json:"error,omitempty"`
	Logs       []string `json:"logs,omitempty"`
}

const (
	UploadDir   = "./uploads"
	OutputDir   = "./outputs"
	MaxFileSize = 50 << 20 // 50 MB
)

func main() {
	// Setup directories
	os.MkdirAll(UploadDir, 0755)
	os.MkdirAll(OutputDir, 0755)

	// Auto cleanup old files on startup
	log.Printf("🧹 Cleaning up old files...")
	cleanupOldFiles(UploadDir, 1*time.Hour)   // Delete files > 1 hour old in uploads
	cleanupOldFiles(OutputDir, 24*time.Hour)  // Delete files > 24 hours old in outputs

	router := mux.NewRouter()

	// API endpoints
	router.HandleFunc("/api/health", healthCheck).Methods("GET")
	router.HandleFunc("/api/upload", uploadFile).Methods("POST")
	router.HandleFunc("/api/process", processData).Methods("POST")
	router.HandleFunc("/api/download/{filename}", downloadFile).Methods("GET")
	router.HandleFunc("/api/files", listUploadedFiles).Methods("GET")
	router.HandleFunc("/api/outputs", listOutputFiles).Methods("GET")
	router.HandleFunc("/api/cleanup", cleanupFilesHandler).Methods("DELETE")

	// Serve static files (frontend) - HARUS PALING AKHIR
	router.PathPrefix("/").Handler(http.FileServer(http.Dir("./static")))

	// CORS configuration
	c := cors.New(cors.Options{
		AllowedOrigins: []string{"*"},
		AllowedMethods: []string{"GET", "POST", "DELETE", "OPTIONS"},
		AllowedHeaders: []string{"Content-Type", "Authorization"},
	})

	handler := c.Handler(router)

	port := os.Getenv("PORT")
	if port == "" {
		port = "8080"
	}

	log.Printf("🚀 Server starting on port %s", port)
	log.Printf("📍 Open browser: http://localhost:%s", port)
	log.Printf("📂 Upload dir: %s", UploadDir)
	log.Printf("📂 Output dir: %s", OutputDir)
	log.Fatal(http.ListenAndServe(":"+port, handler))
}

func healthCheck(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]string{
		"status":  "healthy",
		"service": "trade-history-dashboard",
		"time":    time.Now().Format(time.RFC3339),
	})
}

func uploadFile(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")

	// Parse multipart form
	r.ParseMultipartForm(MaxFileSize)
	file, handler, err := r.FormFile("file")
	if err != nil {
		log.Printf("❌ Error reading uploaded file: %v", err)
		json.NewEncoder(w).Encode(ProcessResponse{
			Success: false,
			Error:   "Failed to read uploaded file: " + err.Error(),
		})
		return
	}
	defer file.Close()

	// Validate file extension
	ext := filepath.Ext(handler.Filename)
	if ext != ".xlsx" && ext != ".xls" && ext != ".csv" {
		log.Printf("❌ Invalid file extension: %s", ext)
		json.NewEncoder(w).Encode(ProcessResponse{
			Success: false,
			Error:   "Only Excel files (.xlsx, .xls) or CSV are allowed",
		})
		return
	}

	// Generate unique filename
	timestamp := time.Now().Unix()
	filename := fmt.Sprintf("%d_%s", timestamp, handler.Filename)
	filePath := filepath.Join(UploadDir, filename)

	// Save file
	dst, err := os.Create(filePath)
	if err != nil {
		log.Printf("❌ Error creating file: %v", err)
		json.NewEncoder(w).Encode(ProcessResponse{
			Success: false,
			Error:   "Failed to save file: " + err.Error(),
		})
		return
	}
	defer dst.Close()

	if _, err := io.Copy(dst, file); err != nil {
		log.Printf("❌ Error writing file: %v", err)
		json.NewEncoder(w).Encode(ProcessResponse{
			Success: false,
			Error:   "Failed to write file: " + err.Error(),
		})
		return
	}

	log.Printf("✅ File uploaded: %s (size: %.2f MB)", filename, float64(handler.Size)/(1024*1024))

	json.NewEncoder(w).Encode(ProcessResponse{
		Success:    true,
		Message:    "File uploaded successfully",
		OutputFile: filename,
	})
}

func processData(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")

	var req ProcessRequest
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		log.Printf("❌ Error decoding request: %v", err)
		json.NewEncoder(w).Encode(ProcessResponse{
			Success: false,
			Error:   "Invalid request body: " + err.Error(),
		})
		return
	}

	log.Printf("📋 Process request received")
	log.Printf("   JISDOR file: %s", req.JisdorFile)
	log.Printf("   Trade files: %v", req.TradeHistoryFiles)
	log.Printf("   Rate spot: %.0f", req.Config.RateSpot)
	log.Printf("   Rate remote: %.0f", req.Config.RateRemote)

	// Validate files exist
	if req.JisdorFile == "" {
		json.NewEncoder(w).Encode(ProcessResponse{
			Success: false,
			Error:   "JISDOR file is required",
		})
		return
	}

	if len(req.TradeHistoryFiles) == 0 {
		json.NewEncoder(w).Encode(ProcessResponse{
			Success: false,
			Error:   "At least one trade history file is required",
		})
		return
	}

	// Prepare Python script arguments
	outputFilename := fmt.Sprintf("dashboard_%d.xlsx", time.Now().Unix())
	outputPath := filepath.Join(OutputDir, outputFilename)

	args := []string{
		"python/processor.py",
		"--jisdor", filepath.Join(UploadDir, req.JisdorFile),
		"--output", outputPath,
		"--rate-spot", fmt.Sprintf("%.0f", req.Config.RateSpot),
		"--rate-remote", fmt.Sprintf("%.0f", req.Config.RateRemote),
	}

	for _, file := range req.TradeHistoryFiles {
		args = append(args, "--trade-file", filepath.Join(UploadDir, file))
	}

	log.Printf("⚙️  Executing Python processor with arguments:")
	for i, arg := range args {
		log.Printf("   [%d] %s", i, arg)
	}

	// Execute Python script
	cmd := exec.Command("python", args...)
	output, err := cmd.CombinedOutput()

	outputStr := string(output)
	log.Printf("📝 Python output:\n%s", outputStr)

	if err != nil {
		log.Printf("❌ Processing error: %v", err)
		json.NewEncoder(w).Encode(ProcessResponse{
			Success: false,
			Error:   fmt.Sprintf("Processing failed: %v", err),
			Logs:    []string{outputStr},
		})
		return
	}

	log.Printf("✅ Processing completed successfully")

	json.NewEncoder(w).Encode(ProcessResponse{
		Success:    true,
		Message:    "Data processed successfully",
		OutputFile: outputFilename,
		Logs:       []string{outputStr},
	})
}

func downloadFile(w http.ResponseWriter, r *http.Request) {
	vars := mux.Vars(r)
	filename := vars["filename"]

	filePath := filepath.Join(OutputDir, filename)

	// Check if file exists
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		log.Printf("❌ File not found: %s", filename)
		http.Error(w, "File not found", http.StatusNotFound)
		return
	}

	w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", filename))
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

	log.Printf("📥 Downloading file: %s", filename)
	http.ServeFile(w, r, filePath)
}

func listUploadedFiles(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")

	files, err := os.ReadDir(UploadDir)
	if err != nil {
		log.Printf("❌ Error reading upload directory: %v", err)
		json.NewEncoder(w).Encode(map[string]interface{}{
			"success": false,
			"error":   "Failed to read upload directory: " + err.Error(),
		})
		return
	}

	var fileList []map[string]interface{}
	for _, file := range files {
		if !file.IsDir() {
			info, _ := file.Info()
			fileList = append(fileList, map[string]interface{}{
				"name":         file.Name(),
				"size":         info.Size(),
				"uploaded_at":  info.ModTime().Format(time.RFC3339),
			})
		}
	}

	log.Printf("📂 Listed %d uploaded files", len(fileList))
	json.NewEncoder(w).Encode(map[string]interface{}{
		"success": true,
		"files":   fileList,
	})
}

func listOutputFiles(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")

	files, err := os.ReadDir(OutputDir)
	if err != nil {
		log.Printf("❌ Error reading output directory: %v", err)
		json.NewEncoder(w).Encode(map[string]interface{}{
			"success": false,
			"error":   "Failed to read output directory: " + err.Error(),
		})
		return
	}

	var fileList []map[string]interface{}
	for _, file := range files {
		if !file.IsDir() {
			info, _ := file.Info()
			fileList = append(fileList, map[string]interface{}{
				"name":       file.Name(),
				"size":       info.Size(),
				"created_at": info.ModTime().Format(time.RFC3339),
			})
		}
	}

	log.Printf("📂 Listed %d output files", len(fileList))
	json.NewEncoder(w).Encode(map[string]interface{}{
		"success": true,
		"files":   fileList,
	})
}

// NEW: Cleanup handler for DELETE API
func cleanupFilesHandler(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")

	var deletedCount int

	// Cleanup uploads folder
	uploadFiles, err := os.ReadDir(UploadDir)
	if err == nil {
		for _, file := range uploadFiles {
			if !file.IsDir() {
				filePath := filepath.Join(UploadDir, file.Name())
				if err := os.Remove(filePath); err == nil {
					deletedCount++
					log.Printf("🗑️  Deleted upload: %s", file.Name())
				}
			}
		}
	}

	// Cleanup outputs folder
	outputFiles, err := os.ReadDir(OutputDir)
	if err == nil {
		for _, file := range outputFiles {
			if !file.IsDir() {
				filePath := filepath.Join(OutputDir, file.Name())
				if err := os.Remove(filePath); err == nil {
					deletedCount++
					log.Printf("🗑️  Deleted output: %s", file.Name())
				}
			}
		}
	}

	log.Printf("✅ Cleanup completed: %d files deleted", deletedCount)
	json.NewEncoder(w).Encode(map[string]interface{}{
		"success": true,
		"message": "Cleanup completed",
		"deleted": deletedCount,
	})
}

// NEW: Helper function to cleanup old files
func cleanupOldFiles(dirPath string, maxAge time.Duration) {
	files, err := os.ReadDir(dirPath)
	if err != nil {
		log.Printf("Error reading directory: %v", err)
		return
	}

	now := time.Now()
	deletedCount := 0

	for _, file := range files {
		if !file.IsDir() {
			info, _ := file.Info()
			fileAge := now.Sub(info.ModTime())

			if fileAge > maxAge {
				filePath := filepath.Join(dirPath, file.Name())
				err := os.Remove(filePath)
				if err == nil {
					deletedCount++
					log.Printf("   🗑️  Cleaned up: %s (age: %v)", file.Name(), fileAge)
				}
			}
		}
	}

	if deletedCount > 0 {
		log.Printf("✅ Auto-cleanup completed in %s: %d files deleted", dirPath, deletedCount)
	}
}
