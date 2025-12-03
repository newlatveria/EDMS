package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net"
	"net/http"
	"sort"
	"strings"
	"sync"

	"github.com/xuri/excelize/v2"
)

// --- Global Data Structures (In-Memory Database) ---
var (
	dataStore = make(map[string]SheetData)
	storeMutex sync.RWMutex
)

type SheetData struct {
	Headers []string
	Rows    [][]string 
}

// ---------------------------------------------------------------------
// --- Utility Functions ---
// ---------------------------------------------------------------------

// getOutboundIP attempts to find the primary non-loopback private IP address.
func getOutboundIP() string {
    // Attempt to connect to a public DNS server (does not send data)
    conn, err := net.Dial("udp", "8.8.8.8:80")
    if err != nil {
        log.Printf("WARN: Could not determine local IP from dialing. Falling back to 127.0.0.1. Error: %v", err)
        return "127.0.0.1"
    }
    defer conn.Close()

    localAddr := conn.LocalAddr().(*net.UDPAddr)
    return localAddr.IP.String()
}

// standardKey creates a case-insensitive, trimmed key for basic matching.
func standardKey(val string) string {
	return strings.TrimSpace(strings.ToLower(val))
}

// levenshteinDistance calculates the Levenshtein distance (edit distance).
func levenshteinDistance(s1, s2 string) int {
	if s1 == s2 { return 0 }
	if len(s1) == 0 { return len(s2) }
	if len(s2) == 0 { return len(s1) }

	v0 := make([]int, len(s2)+1)
	v1 := make([]int, len(s2)+1)

	for i := range v0 { v0[i] = i }

	for i := 1; i <= len(s1); i++ {
		v1[0] = i
		for j := 1; j <= len(s2); j++ {
			cost := 1
			if s1[i-1] == s2[j-1] { cost = 0 }
			v1[j] = min(v1[j-1]+1, v0[j]+1, v0[j-1]+cost)
		}
		copy(v0, v1)
	}
	return v1[len(s2)]
}

func min(a, b, c int) int {
	if a < b {
		if a < c { return a }
		return c
	}
	if b < c { return b }
	return c
}

func max(a, b int) int {
	if a > b { return a }
	return b
}

// isFuzzyMatch checks if two values are a fuzzy match based on the threshold.
func isFuzzyMatch(val1, val2 string, threshold int) bool {
	s1 := standardKey(val1)
	s2 := standardKey(val2)
	if s1 == s2 { return true }
	if s1 == "" || s2 == "" { return false }

	maxLen := max(len(s1), len(s2))
	if maxLen == 0 { return true }

	dist := levenshteinDistance(s1, s2)
	
	return dist*100 <= maxLen*threshold
}

// ---------------------------------------------------------------------
// --- API Data Structures (Unchanged) ---
// ---------------------------------------------------------------------

type MatchRequest struct {
	Sheet1           string `json:"sheet1"`
	Sheet2           string `json:"sheet2"`
	UseFuzzy         bool   `json:"useFuzzy"`
	FuzzyThreshold   int    `json:"fuzzyThreshold"`
	IsTargeted       bool   `json:"isTargeted"` 
}

type MatchResult struct {
	OriginalRow1 int    `json:"originalRow1"`
	OriginalRow2 int    `json:"originalRow2"`
	Val1         string `json:"val1"`
	Val2         string `json:"val2"`
	IsFuzzy      bool   `json:"isFuzzy"`
}

type MatchGroup struct {
	Tab1    string        `json:"tab1"`
	Tab2    string        `json:"tab2"`
	Header1 string        `json:"header1"`
	Header2 string        `json:"header2"`
	Matches []MatchResult `json:"matches"`
}

// ---------------------------------------------------------------------
// --- API Endpoint Handlers ---
// ---------------------------------------------------------------------

// uploadHandler handles file input, parsing, and data storage.
func uploadHandler(w http.ResponseWriter, r *http.Request) {
	log.Printf("INFO: Handling file upload request.")
	if r.Method != "POST" {
		log.Printf("ERROR: Method not allowed for /api/upload: %s", r.Method)
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}
	
	storeMutex.Lock()
	dataStore = make(map[string]SheetData)
	storeMutex.Unlock()
	log.Printf("DEBUG: In-memory data store cleared.")

	file, header, err := r.FormFile("excelFile")
	if err != nil {
		log.Printf("ERROR: Failed to retrieve file from form: %v", err)
		http.Error(w, fmt.Sprintf("Error retrieving file: %v", err), http.StatusBadRequest)
		return
	}
	defer file.Close()
	log.Printf("INFO: Received file: %s (%d bytes)", header.Filename, header.Size)

	buf := bytes.NewBuffer(nil)
	if _, err := io.Copy(buf, file); err != nil {
		log.Printf("ERROR: Failed to read file content: %v", err)
		http.Error(w, "Error reading file content", http.StatusInternalServerError)
		return
	}

	f, err := excelize.OpenReader(buf)
	if err != nil {
		log.Printf("ERROR: Failed to open Excel file with excelize: %v", err)
		http.Error(w, fmt.Sprintf("Error opening Excel file: %v", err), http.StatusInternalServerError)
		return
	}

	sheetNames := f.GetSheetMap()
	storeMutex.Lock()
	defer storeMutex.Unlock()

	names := make([]string, 0, len(sheetNames))
	for _, sheetName := range sheetNames {
		names = append(names, sheetName)
		
		rows, err := f.GetRows(sheetName)
		if err != nil || len(rows) == 0 {
			log.Printf("WARN: Skipping empty or unreadable sheet: %s", sheetName)
			continue
		}

		headers := rows[0]
		dataRows := make([][]string, len(rows)-1)
		
		for i, row := range rows[1:] {
			data := make([]string, len(row))
			for j, cell := range row {
				data[j] = cell
			}
			dataRows[i] = data
		}

		dataStore[sheetName] = SheetData{
			Headers: headers,
			Rows:    dataRows,
		}
		log.Printf("DEBUG: Parsed sheet '%s' with %d data rows and %d columns.", sheetName, len(dataRows), len(headers))
	}
	
	sort.Strings(names)
	log.Printf("INFO: File processing complete. %d sheets stored.", len(names))

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]interface{}{
		"sheetNames": names,
		"message":    "File parsed and stored successfully.",
	})
}

// matchHandler executes the all-to-all column comparison logic.
func matchHandler(w http.ResponseWriter, r *http.Request) {
	log.Printf("INFO: Handling matching request.")
	if r.Method != "POST" {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	var req MatchRequest
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		log.Printf("ERROR: Invalid match request body: %v", err)
		http.Error(w, fmt.Sprintf("Invalid request body: %v", err), http.StatusBadRequest)
		return
	}
	
	log.Printf("DEBUG: Matching sheets '%s' vs '%s'. Fuzzy: %t (Threshold: %d)", req.Sheet1, req.Sheet2, req.UseFuzzy, req.FuzzyThreshold)

	storeMutex.RLock()
	sheet1Data, ok1 := dataStore[req.Sheet1]
	sheet2Data, ok2 := dataStore[req.Sheet2]
	storeMutex.RUnlock()

	if !ok1 || !ok2 {
		log.Printf("ERROR: One or both sheets not found: %s, %s", req.Sheet1, req.Sheet2)
		http.Error(w, "One or both sheets not found in store.", http.StatusBadRequest)
		return
	}
	
	allMatches := make([]MatchGroup, 0)
	numCols1 := len(sheet1Data.Headers)
	numCols2 := len(sheet2Data.Headers)
	totalComparisons := 0

	matchedPairs := make(map[string]struct{}) 

	for c1 := 0; c1 < numCols1; c1++ {
		for c2 := 0; c2 < numCols2; c2++ {
			totalComparisons++
			matches := make([]MatchResult, 0)
			
			keyMap2 := make(map[string][]int) 
			for r2, row2 := range sheet2Data.Rows {
				if c2 < len(row2) {
					key := standardKey(row2[c2])
					if key != "" {
						keyMap2[key] = append(keyMap2[key], r2 + 2)
					}
				}
			}

			for r1, row1 := range sheet1Data.Rows {
				if c1 >= len(row1) { continue }
				val1 := row1[c1]
				key1 := standardKey(val1)
				row1Idx := r1 + 2

				// 1. Exact/Standard Match
				if row2Indices, ok := keyMap2[key1]; ok {
					for _, row2Idx := range row2Indices {
						pairKey := fmt.Sprintf("%d-%d", row1Idx, row2Idx)
						if _, exists := matchedPairs[pairKey]; exists { continue }
						
						val2 := sheet2Data.Rows[row2Idx-2][c2] 
						
						matches = append(matches, MatchResult{
							OriginalRow1: row1Idx,
							OriginalRow2: row2Idx,
							Val1: val1,
							Val2: val2,
							IsFuzzy: false,
						})
						matchedPairs[pairKey] = struct{}{}
					}
				}

				// 2. Fuzzy Match (Only if enabled)
				if req.UseFuzzy {
					for r2, row2 := range sheet2Data.Rows {
						row2Idx := r2 + 2
						pairKey := fmt.Sprintf("%d-%d", row1Idx, row2Idx)
						if _, exists := matchedPairs[pairKey]; exists { continue } 

						if c2 >= len(row2) { continue }
						val2 := row2[c2]

						if isFuzzyMatch(val1, val2, req.FuzzyThreshold) {
							matches = append(matches, MatchResult{
								OriginalRow1: row1Idx,
								OriginalRow2: row2Idx,
								Val1: val1,
								Val2: val2,
								IsFuzzy: true,
							})
							matchedPairs[pairKey] = struct{}{}
						}
					}
				}
			}
			
			if len(matches) > 0 {
				header1 := sheet1Data.Headers[c1]
				header2 := sheet2Data.Headers[c2]
				allMatches = append(allMatches, MatchGroup{
					Tab1: req.Sheet1, Tab2: req.Sheet2,
					Header1: header1, Header2: header2,
					Matches: matches,
				})
			}
		}
	}
	
	log.Printf("INFO: Matching complete. Ran %d column pair comparisons, found %d match groups.", totalComparisons, len(allMatches))

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(allMatches)
}

// dataHandler retrieves the full data for a specific sheet.
func dataHandler(w http.ResponseWriter, r *http.Request) {
	pathParts := strings.Split(r.URL.Path, "/")
	if len(pathParts) < 4 {
		http.Error(w, "Sheet name not specified.", http.StatusBadRequest)
		return
	}
	sheetName := pathParts[3]

	storeMutex.RLock()
	data, ok := dataStore[sheetName]
	storeMutex.RUnlock()

	if !ok {
		log.Printf("WARN: Data request failed. Sheet not found: %s", sheetName)
		http.Error(w, "Sheet not found.", http.StatusNotFound)
		return
	}
	log.Printf("INFO: Serving raw data for sheet: %s", sheetName)

	response := struct {
		Headers []string     `json:"headers"`
		Rows    [][]string   `json:"rows"`
	}{
		Headers: data.Headers,
		Rows:    data.Rows,
	}

	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(response)
}

// serveFile is a helper to serve static files from the root directory.
func serveFile(w http.ResponseWriter, r *http.Request, filename string, contentType string) {
	w.Header().Set("Content-Type", contentType)
	http.ServeFile(w, r, filename)
}

func main() {
	// --- Static File Handlers ---
	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path == "/" {
			serveFile(w, r, "index.html", "text/html")
		} else {
			http.ServeFile(w, r, r.URL.Path[1:])
		}
	})

	http.HandleFunc("/style.css", func(w http.ResponseWriter, r *http.Request) { serveFile(w, r, "style.css", "text/css") })
	http.HandleFunc("/app.js", func(w http.ResponseWriter, r *http.Request) { serveFile(w, r, "app.js", "application/javascript") })

	// --- API Handlers ---
	http.HandleFunc("/api/upload", uploadHandler)
	http.HandleFunc("/api/match", matchHandler)
	http.HandleFunc("/api/data/", dataHandler)

	port := "8080"
	ip := getOutboundIP()
	
	log.Printf("=====================================================")
	log.Printf("INFO: Server starting on port %s.", port)
	log.Printf("INFO: Access at: http://%s:%s", ip, port)
	log.Printf("INFO: Also available at: http://localhost:%s", port)
	log.Printf("=====================================================")
	log.Fatal(http.ListenAndServe(":"+port, nil))
}