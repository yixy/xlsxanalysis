package main

import (
	"database/sql"
	"io/ioutil"
	"os"
	"path/filepath"
	"testing"
	"time"

	_ "modernc.org/sqlite"
	"github.com/tealeg/xlsx"
)

// TestGetXlsxCellCoords tests the getXlsxCellCoords function
func TestGetXlsxCellCoords(t *testing.T) {
	tests := []struct {
		input    string
		expected []int // [col, row]
	}{
		{"A1", []int{0, 0}},
		{"B2", []int{1, 1}},
		{"Z10", []int{25, 9}},
		{"AA1", []int{26, 0}},
		{"IV20", []int{255, 19}},
	}

	for _, test := range tests {
		col, row := getXlsxCellCoords(test.input)
		if col != test.expected[0] || row != test.expected[1] {
			t.Errorf("getXlsxCellCoords(%s) = (%d, %d), expected (%d, %d)", 
				test.input, col, row, test.expected[0], test.expected[1])
		}
	}
}

// TestGetColumnIndex tests the getColumnIndex function
func TestGetColumnIndex(t *testing.T) {
	tests := []struct {
		input    string
		expected int
	}{
		{"A", 0},
		{"B", 1},
		{"Z", 25},
		{"AA", 26},
		{"AB", 27},
		{"BA", 52},
	}

	for _, test := range tests {
		result := getColumnIndex(test.input)
		if result != test.expected {
			t.Errorf("getColumnIndex(%s) = %d, expected %d", test.input, result, test.expected)
		}
	}
}

// TestLoadConfig tests the loadConfig function
func TestLoadConfig(t *testing.T) {
	// Create a temporary config file
	tempDir := t.TempDir()
	configPath := filepath.Join(tempDir, "test_config.yaml")
	
	configContent := `
db_path: "./test.db"
source_dir: "./xlsx_files"
batch_size: 50
export_path: "./export.xlsx"
tables:
  - table_name: "test_table"
    sheet_index: 0
    columns:
      - filename_source: "name"
        excel_col: "A"
        db_col: "name"
        db_type: "TEXT"
`
	
	err := ioutil.WriteFile(configPath, []byte(configContent), 0644)
	if err != nil {
		t.Fatalf("Failed to write temp config file: %v", err)
	}

	config, err := loadConfig(configPath)
	if err != nil {
		t.Fatalf("loadConfig failed: %v", err)
	}

	if config.DBPath != "./test.db" {
		t.Errorf("Expected DBPath to be './test.db', got '%s'", config.DBPath)
	}
	
	if config.BatchSize != 50 {
		t.Errorf("Expected BatchSize to be 50, got %d", config.BatchSize)
	}
	
	if len(config.Tables) != 1 {
		t.Errorf("Expected 1 table, got %d", len(config.Tables))
	}
	
	if config.Tables[0].TableName != "test_table" {
		t.Errorf("Expected table name 'test_table', got '%s'", config.Tables[0].TableName)
	}
}

// TestGetXlsxFiles tests the getXlsxFiles function
func TestGetXlsxFiles(t *testing.T) {
	// Create a temporary directory with some test files
	tempDir := t.TempDir()
	
	// Create test files
	xlsxFile1 := filepath.Join(tempDir, "test1.xlsx")
	xlsFile1 := filepath.Join(tempDir, "test2.xls")
	txtFile1 := filepath.Join(tempDir, "test3.txt")
	
	// Create empty files
	_, err := os.Create(xlsxFile1)
	if err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}
	_, err = os.Create(xlsFile1)
	if err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}
	_, err = os.Create(txtFile1)
	if err != nil {
		t.Fatalf("Failed to create test file: %v", err)
	}

	files, err := getXlsxFiles(tempDir)
	if err != nil {
		t.Fatalf("getXlsxFiles failed: %v", err)
	}

	if len(files) != 2 {
		t.Errorf("Expected 2 files, got %d", len(files))
	}

	foundXlsx := false
	foundXls := false
	for _, file := range files {
		if file.Name == "test1.xlsx" {
			foundXlsx = true
		}
		if file.Name == "test2.xls" {
			foundXls = true
		}
	}

	if !foundXlsx {
		t.Error("Did not find test1.xlsx in results")
	}
	
	if !foundXls {
		t.Error("Did not find test2.xls in results")
	}
}

// TestCreateTable tests the createTable function
func TestCreateTable(t *testing.T) {
	// Create a temporary SQLite database
	tempDir := t.TempDir()
	dbPath := filepath.Join(tempDir, "test.db")
	
	db, err := sql.Open("sqlite", dbPath)
	if err != nil {
		t.Fatalf("Failed to open database: %v", err)
	}
	defer db.Close()

	tableConfig := &Table{
		TableName: "test_table",
		Columns: []Column{
			{
				ExcelCol: "A",
				DBCol:    "name",
				DBType:   "TEXT",
			},
			{
				ExcelCol: "B",
				DBCol:    "value",
				DBType:   "INTEGER",
			},
		},
	}

	err = createTable(db, tableConfig)
	if err != nil {
		t.Fatalf("createTable failed: %v", err)
	}

	// Check if table was created
	var tableName string
	err = db.QueryRow("SELECT name FROM sqlite_master WHERE type='table' AND name='test_table'").Scan(&tableName)
	if err != nil {
		t.Fatalf("Table was not created: %v", err)
	}
	
	if tableName != "test_table" {
		t.Errorf("Expected table name 'test_table', got '%s'", tableName)
	}
}

// TestBatchInsertData tests the batchInsertData function
func TestBatchInsertData(t *testing.T) {
	// Create a temporary SQLite database
	tempDir := t.TempDir()
	dbPath := filepath.Join(tempDir, "test.db")
	
	db, err := sql.Open("sqlite", dbPath)
	if err != nil {
		t.Fatalf("Failed to open database: %v", err)
	}
	defer db.Close()

	tableConfig := &Table{
		TableName: "test_table",
		Columns: []Column{
			{
				ExcelCol: "A",
				DBCol:    "name",
				DBType:   "TEXT",
			},
			{
				ExcelCol: "B",
				DBCol:    "value",
				DBType:   "INTEGER",
			},
		},
	}

	// Create the table first
	err = createTable(db, tableConfig)
	if err != nil {
		t.Fatalf("createTable failed: %v", err)
	}

	// Prepare test data
	rows := [][]string{
		{"John", "25"},
		{"Jane", "30"},
		{"Bob", "35"},
	}

	sourceFileName := "test_file.xlsx"
	batchSize := 100

	err = batchInsertData(db, tableConfig, rows, sourceFileName, batchSize)
	if err != nil {
		t.Fatalf("batchInsertData failed: %v", err)
	}

	// Verify data was inserted
	var count int
	err = db.QueryRow("SELECT COUNT(*) FROM test_table").Scan(&count)
	if err != nil {
		t.Fatalf("Failed to count records: %v", err)
	}

	if count != 3 {
		t.Errorf("Expected 3 records, got %d", count)
	}

	// Check if source file name was recorded
	var sourceFile string
	err = db.QueryRow("SELECT DISTINCT source_file FROM test_table").Scan(&sourceFile)
	if err != nil {
		t.Fatalf("Failed to get source file: %v", err)
	}

	if sourceFile != sourceFileName {
		t.Errorf("Expected source file '%s', got '%s'", sourceFileName, sourceFile)
	}
}

// TestExportToExcel tests the exportToExcel function
func TestExportToExcel(t *testing.T) {
	// Create a temporary SQLite database with test data
	tempDir := t.TempDir()
	dbPath := filepath.Join(tempDir, "test.db")
	
	db, err := sql.Open("sqlite", dbPath)
	if err != nil {
		t.Fatalf("Failed to open database: %v", err)
	}
	defer db.Close()

	tableConfig := Table{
		TableName: "test_export_table",
		Columns: []Column{
			{
				ExcelCol: "A",
				DBCol:    "name",
				DBType:   "TEXT",
			},
			{
				ExcelCol: "B",
				DBCol:    "value",
				DBType:   "INTEGER",
			},
		},
	}

	// Create the table and insert test data
	err = createTable(db, &tableConfig)
	if err != nil {
		t.Fatalf("createTable failed: %v", err)
	}

	// Insert some test data
	_, err = db.Exec(`INSERT INTO test_export_table (source_file, processed_time, name, value) 
	                  VALUES (?, ?, ?, ?)`, "test.xlsx", time.Now(), "Test Name", 123)
	if err != nil {
		t.Fatalf("Failed to insert test data: %v", err)
	}

	// Export to Excel
	exportPath := filepath.Join(tempDir, "export.xlsx")
	tables := []Table{tableConfig}
	
	err = exportToExcel(db, tables, exportPath)
	if err != nil {
		t.Fatalf("exportToExcel failed: %v", err)
	}

	// Verify the Excel file was created
	if _, err := os.Stat(exportPath); os.IsNotExist(err) {
		t.Fatalf("Exported Excel file does not exist: %v", err)
	}

	// Open and verify the content
	xlFile, err := xlsx.OpenFile(exportPath)
	if err != nil {
		t.Fatalf("Failed to open exported Excel file: %v", err)
	}

	if len(xlFile.Sheets) != 1 {
		t.Errorf("Expected 1 sheet, got %d", len(xlFile.Sheets))
	}

	sheet := xlFile.Sheets[0]
	if sheet.Name != "test_export_table" {
		t.Errorf("Expected sheet name 'test_export_table', got '%s'", sheet.Name)
	}

	if len(sheet.Rows) < 2 { // Header + data row
		t.Errorf("Expected at least 2 rows, got %d", len(sheet.Rows))
	}
}

// TestMainFunctionIntegration performs an integration test of the main flow
func TestMainFunctionIntegration(t *testing.T) {
	// Create a temporary directory for our test
	tempDir := t.TempDir()
	
	// Create a config file
	configDir := filepath.Join(tempDir, "config")
	err := os.MkdirAll(configDir, 0755)
	if err != nil {
		t.Fatalf("Failed to create config dir: %v", err)
	}
	
	configPath := filepath.Join(configDir, "config.yaml")
	configContent := `
db_path: "` + filepath.Join(tempDir, "test.db") + `"
source_dir: "` + filepath.Join(tempDir, "xlsx_files") + `"
batch_size: 10
export_path: "` + filepath.Join(tempDir, "export.xlsx") + `"
tables:
  - table_name: "integration_test_table"
    sheet_index: 0
    columns:
      - excel_col: "A"
        db_col: "name"
        db_type: "TEXT"
      - excel_col: "B"
        db_col: "value"
        db_type: "INTEGER"
`
	
	err = ioutil.WriteFile(configPath, []byte(configContent), 0644)
	if err != nil {
		t.Fatalf("Failed to write config: %v", err)
	}

	// Create source directory
	sourceDir := filepath.Join(tempDir, "xlsx_files")
	err = os.MkdirAll(sourceDir, 0755)
	if err != nil {
		t.Fatalf("Failed to create source dir: %v", err)
	}

	// Create a test Excel file
	excelPath := filepath.Join(sourceDir, "test.xlsx")
	file := xlsx.NewFile()
	sheet, _ := file.AddSheet("Sheet1")
	
	row := sheet.AddRow()
	cell := row.AddCell()
	cell.Value = "Name"
	cell = row.AddCell()
	cell.Value = "Value"
	
	row = sheet.AddRow()
	cell = row.AddCell()
	cell.Value = "Test Name"
	cell = row.AddCell()
	cell.Value = "100"
	
	err = file.Save(excelPath)
	if err != nil {
		t.Fatalf("Failed to save Excel file: %v", err)
	}

	// Load config and test
	config, err := loadConfig(configPath)
	if err != nil {
		t.Fatalf("Failed to load config: %v", err)
	}

	// Connect to database
	db, err := sql.Open("sqlite", config.DBPath)
	if err != nil {
		t.Fatalf("Failed to connect to database: %v", err)
	}
	defer db.Close()

	// Process the file
	for _, tableConfig := range config.Tables {
		// Create table
		err = createTable(db, &tableConfig)
		if err != nil {
			t.Fatalf("Failed to create table: %v", err)
		}

		// Get Excel files
		xlsxFiles, err := getXlsxFiles(config.SourceDir)
		if err != nil {
			t.Fatalf("Failed to get XLSX files: %v", err)
		}

		// Process each file
		for _, fileInfo := range xlsxFiles {
			xlFile, err := xlsx.OpenFile(fileInfo.Path)
			if err != nil {
				t.Fatalf("Failed to open file: %v", err)
			}

			sheet := xlFile.Sheets[tableConfig.SheetIndex]

			// Convert to string arrays
			rows := make([][]string, 0)
			for _, row := range sheet.Rows {
				if row != nil {
					rowStr := make([]string, 0)
					for _, cell := range row.Cells {
						rowStr = append(rowStr, cell.String())
					}
					rows = append(rows, rowStr)
				}
			}

			// Skip header row
			if len(rows) > 1 {
				rows = rows[1:]
			}

			err = batchInsertData(db, &tableConfig, rows, fileInfo.Name, config.BatchSize)
			if err != nil {
				t.Fatalf("Failed to insert data: %v", err)
			}
		}
	}

	// Verify data was inserted
	var count int
	err = db.QueryRow("SELECT COUNT(*) FROM integration_test_table").Scan(&count)
	if err != nil {
		t.Fatalf("Failed to count records: %v", err)
	}

	if count != 1 {
		t.Errorf("Expected 1 record, got %d", count)
	}
}