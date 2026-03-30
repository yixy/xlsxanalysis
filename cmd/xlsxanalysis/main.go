package main

import (
	"database/sql"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
	"gopkg.in/yaml.v3"
	_ "modernc.org/sqlite" // 纯 Go 实现的 SQLite 驱动，无需 CGO
)

// Config 对应 config.yaml 结构
type Config struct {
	DBPath     string        `yaml:"db_path"`
	SourceDir  string        `yaml:"source_dir"`
	BatchSize  int           `yaml:"batch_size"`
	ExportPath string        `yaml:"export_path"`
	Tables     []TableConfig `yaml:"tables"`
}

type TableConfig struct {
	TableName  string           `yaml:"table_name"`
	SheetIndex int              `yaml:"sheet_index"`
	Columns    []ColumnProperty `yaml:"columns"`
}

// ColumnProperty 兼容 YAML 中混合定义的列属性
type ColumnProperty struct {
	FilenameSource string `yaml:"filename_source,omitempty"`
	FilenameCell   string `yaml:"filename_cell,omitempty"`
	ExcelCol       string `yaml:"excel_col,omitempty"`
	DBCol          string `yaml:"db_col,omitempty"`
	DBType         string `yaml:"db_type,omitempty"`
}

func main() {
	// 1. 加载配置
	config, err := loadConfig("./config/config.yaml")
	if err != nil {
		log.Fatalf("无法加载配置文件: %v", err)
	}

	// 2. 初始化数据库
	db, err := sql.Open("sqlite", config.DBPath)
	if err != nil {
		log.Fatalf("无法打开数据库: %v", err)
	}
	defer db.Close()

	if err := initDatabase(db, config); err != nil {
		log.Fatalf("数据库初始化失败: %v", err)
	}

	// 3. 扫描目录并处理文件
	files, err := os.ReadDir(config.SourceDir)
	if err != nil {
		log.Fatalf("无法读取源目录: %v", err)
	}

	for _, file := range files {
		if !file.IsDir() && strings.HasSuffix(file.Name(), ".xlsx") {
			filePath := filepath.Join(config.SourceDir, file.Name())
			fmt.Printf("正在处理文件: %s\n", file.Name())
			if err := processExcel(db, filePath, config); err != nil {
				log.Printf("处理文件 %s 出错: %v", file.Name(), err)
			}
		}
	}

	// 4. 导出到 Excel (如果配置了 export_path)
	if config.ExportPath != "" {
		fmt.Printf("正在导出数据到: %s\n", config.ExportPath)
		if err := exportToExcel(db, config); err != nil {
			log.Fatalf("导出 Excel 失败: %v", err)
		}
	}

	fmt.Println("任务处理完成！")
}

func loadConfig(path string) (*Config, error) {
	buf, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	var conf Config
	err = yaml.Unmarshal(buf, &conf)
	return &conf, err
}

func initDatabase(db *sql.DB, conf *Config) error {
	for _, table := range conf.Tables {
		var colDefs []string
		// 固定添加 source_file 和 processed_at 字段用于追踪来源和处理时间
		colDefs = append(colDefs, "source_file TEXT", "processed_at TEXT")

		for _, col := range table.Columns {
			if col.DBCol != "" {
				colDefs = append(colDefs, fmt.Sprintf("%s %s", col.DBCol, col.DBType))
			}
		}
		query := fmt.Sprintf("CREATE TABLE IF NOT EXISTS %s (%s)", table.TableName, strings.Join(colDefs, ", "))
		_, err := db.Exec(query)
		if err != nil {
			return err
		}
	}
	return nil
}

func processExcel(db *sql.DB, filePath string, conf *Config) error {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return err
	}
	defer f.Close()

	for _, tableConf := range conf.Tables {
		sheetList := f.GetSheetList()
		if tableConf.SheetIndex >= len(sheetList) {
			return fmt.Errorf("sheet 索引 %d 超出范围", tableConf.SheetIndex)
		}
		sheetName := sheetList[tableConf.SheetIndex]

		// 获取数据来源标识
		var sourceValue string
		var mappingCols []ColumnProperty

		for _, col := range tableConf.Columns {
			if col.FilenameSource != "" {
				if col.FilenameSource == "cell" {
					// 稍后根据下面的 FilenameCell 获取
				} else if col.FilenameSource == "name" {
					sourceValue = filepath.Base(filePath)
				}
			}
			if col.FilenameCell != "" {
				val, _ := f.GetCellValue(sheetName, col.FilenameCell)
				sourceValue = val
			}
			if col.ExcelCol != "" {
				mappingCols = append(mappingCols, col)
			}
		}
		if sourceValue == "" {
			sourceValue = filepath.Base(filePath) // 兜底使用文件名
		}

		// 读取所有行
		rows, err := f.GetRows(sheetName)
		if err != nil {
			return err
		}

		// 准备 SQL
		var dbCols []string
		var placeholders []string
		dbCols = append(dbCols, "source_file", "processed_at")
		placeholders = append(placeholders, "?", "?")
		for _, col := range mappingCols {
			dbCols = append(dbCols, col.DBCol)
			placeholders = append(placeholders, "?")
		}

		insertSQL := fmt.Sprintf("INSERT INTO %s (%s) VALUES (%s)",
			tableConf.TableName,
			strings.Join(dbCols, ","),
			strings.Join(placeholders, ","))

		// 开启事务批量写入
		tx, err := db.Begin()
		if err != nil {
			return err
		}

		stmt, err := tx.Prepare(insertSQL)
		if err != nil {
			tx.Rollback()
			return err
		}

		count := 0
		for i, row := range rows {
			// 假设第一行通常是标题，或者根据实际逻辑跳过
			if i == 0 {
				continue
			}

			values := []interface{}{sourceValue, time.Now().Format("2006-01-02 15:04:05")}
			//hasData 的作用是过滤掉无效行。它确保程序只向数据库插入那些“在你关心的列上确实存在数据（或者至少该行长度覆盖到了你关心的列）”的行，防止数据库被大量的空记录填满。
			hasData := false
			for _, col := range mappingCols {
				colIdx, _ := excelize.ColumnNameToNumber(col.ExcelCol)
				val := ""
				if colIdx-1 < len(row) {
					val = row[colIdx-1]
					hasData = true
				}
				values = append(values, val)
			}

			if !hasData {
				continue
			}

			_, err := stmt.Exec(values...)
			if err != nil {
				log.Printf("写入行 %d 失败: %v", i, err)
				continue
			}

			count++
			if count%conf.BatchSize == 0 {
				// 提交当前批次（此处为了简化，实际生产可调整事务粒度）
			}
		}

		stmt.Close()
		err = tx.Commit()
		if err != nil {
			return err
		}
		fmt.Printf("表 %s 已写入 %d 条数据\n", tableConf.TableName, count)
	}

	return nil
}

func exportToExcel(db *sql.DB, conf *Config) error {
	f := excelize.NewFile()
	defer f.Close()

	for i, tableConf := range conf.Tables {
		sheetName := tableConf.TableName
		// 如果是第一个表格，重命名默认的 Sheet1，否则新建 Sheet
		if i == 0 {
			f.SetSheetName("Sheet1", sheetName)
		} else {
			f.NewSheet(sheetName)
		}

		// 构造表头
		headers := []string{"source_file", "processed_at"}
		for _, col := range tableConf.Columns {
			if col.DBCol != "" {
				headers = append(headers, col.DBCol)
			}
		}

		// 写入表头
		for colIdx, header := range headers {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
			f.SetCellValue(sheetName, cell, header)
		}

		// 查询数据
		query := fmt.Sprintf("SELECT %s FROM %s", strings.Join(headers, ", "), tableConf.TableName)
		rows, err := db.Query(query)
		if err != nil {
			return err
		}
		defer rows.Close()

		rowIdx := 2
		for rows.Next() {
			// 准备接收数据的切片
			values := make([]interface{}, len(headers))
			valuePtrs := make([]interface{}, len(headers))
			for i := range values {
				valuePtrs[i] = &values[i]
			}

			if err := rows.Scan(valuePtrs...); err != nil {
				return err
			}

			// 写入行数据
			for colIdx, val := range values {
				cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowIdx)
				f.SetCellValue(sheetName, cell, val)
			}
			rowIdx++
		}
		fmt.Printf("表 %s 已导出 %d 条数据\n", sheetName, rowIdx-2)
	}

	return f.SaveAs(conf.ExportPath)
}
