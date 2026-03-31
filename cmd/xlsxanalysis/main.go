package main

import (
	"database/sql"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
	"gopkg.in/yaml.v2"
	_ "modernc.org/sqlite"
)

// Config 配置文件结构
type Config struct {
	DBPath     string  `yaml:"db_path"`
	SourceDir  string  `yaml:"source_dir"`
	BatchSize  int     `yaml:"batch_size"`
	ExportPath string  `yaml:"export_path"`
	Tables     []Table `yaml:"tables"`
}

// Table 表配置结构
type Table struct {
	TableName  string   `yaml:"table_name"`
	SheetIndex int      `yaml:"sheet_index"`
	Columns    []Column `yaml:"columns"`
}

// Column 列配置结构
type Column struct {
	FilenameSource string `yaml:"filename_source"` // 指定文件名来源类型 (name/cell)
	FilenameCell   string `yaml:"filename_cell"`   // 如果FilenameSource是cell，则从此单元格获取文件名
	ExcelCol       string `yaml:"excel_col"`       // Excel列名
	DBCol          string `yaml:"db_col"`          // 数据库字段名
	DBType         string `yaml:"db_type"`         // 数据库类型
}

// XlsxFileInfo 存储Excel文件信息
type XlsxFileInfo struct {
	Name string
	Path string
}

// main 是 XLSX 文件分析程序的入口点。
// 该程序会读取配置文件，连接到 SQLite 数据库，并将 XLSX 文件中的数据导入到数据库中。
// 支持批量导入以及从多个 Excel 文件中提取数据并存储到指定的数据库表中。
// 可选地，可以将数据库中的数据导出回 Excel 文件。
func main() {
	fmt.Println("开始解析XLSX文件...")

	// 读取配置文件
	config, err := loadConfig("./config/config.yaml")
	if err != nil {
		fmt.Printf("加载配置文件失败: %v\n", err)
		return
	}

	// 确保数据库目录存在
	dbDir := filepath.Dir(config.DBPath)
	if err := os.MkdirAll(dbDir, 0755); err != nil {
		fmt.Printf("创建数据库目录失败: %v\n", err)
		return
	}

	// 连接SQLite数据库
	db, err := sql.Open("sqlite", config.DBPath)
	if err != nil {
		fmt.Printf("连接数据库失败: %v\n", err)
		return
	}
	defer db.Close()

	// 获取要处理的xlsx文件列表
	xlsxFiles, err := getXlsxFiles(config.SourceDir)
	if err != nil {
		fmt.Printf("获取XLSX文件列表失败: %v\n", err)
		return
	}

	fmt.Printf("找到 %d 个XLSX文件\n", len(xlsxFiles))

	// 处理每个表配置
	for _, tableConfig := range config.Tables {
		// 创建表
		err = createTable(db, &tableConfig)
		if err != nil {
			fmt.Printf("创建表 %s 失败: %v\n", tableConfig.TableName, err)
			continue
		}

		// 处理每个XLSX文件
		for _, fileInfo := range xlsxFiles {
			fmt.Printf("正在处理文件: %s\n", fileInfo.Name)

			// 解析Excel文件
			xlFile, err := xlsx.OpenFile(fileInfo.Path)
			if err != nil {
				fmt.Printf("打开文件失败 %s: %v\n", fileInfo.Name, err)
				continue
			}

			// 检查是否有足够的工作表
			if len(xlFile.Sheets) <= tableConfig.SheetIndex {
				fmt.Printf("工作表索引 %d 超出范围，文件只有 %d 个工作表\n", tableConfig.SheetIndex, len(xlFile.Sheets))
				continue
			}

			sheet := xlFile.Sheets[tableConfig.SheetIndex]

			// 转换为字符串数组
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

			// 跳过空表
			if len(rows) <= 1 {
				fmt.Printf("文件 %s 的工作表为空或只包含标题行\n", fileInfo.Name)
				continue
			}

			// 跳过标题行
			rows = rows[1:]

			// 找到 filename source 列的配置
			var filenameCol *Column
			for i, col := range tableConfig.Columns {
				if col.FilenameSource != "" {
					filenameCol = &tableConfig.Columns[i]
					break
				}
			}

			// 确定源文件名，可能是文件名或单元格内容
			sourceFileName := fileInfo.Name
			if filenameCol != nil && filenameCol.FilenameSource == "cell" && filenameCol.FilenameCell != "" {
				// 从单元格获取文件名
				cellCoord := filenameCol.FilenameCell
				colIdx, rowIdx := getXlsxCellCoords(cellCoord)
				if rowIdx < len(sheet.Rows) && colIdx < len(sheet.Rows[rowIdx].Cells) {
					sourceFileName = sheet.Rows[rowIdx].Cells[colIdx].String()
				}
			}

			err = batchInsertData(db, &tableConfig, rows, sourceFileName, config.BatchSize)
			if err != nil {
				fmt.Printf("插入数据失败 %s: %v\n", fileInfo.Name, err)
			}
		}
	}

	// 如果配置了导出路径，则导出数据库表到Excel文件
	if config.ExportPath != "" {
		err = exportToExcel(db, config.Tables, config.ExportPath)
		if err != nil {
			fmt.Printf("导出到Excel失败: %v\n", err)
		} else {
			fmt.Printf("数据已导出到: %s\n", config.ExportPath)
		}
	}

	fmt.Println("所有XLSX文件处理完成!")
}

// exportToExcel 将数据库表导出到Excel文件
func exportToExcel(db *sql.DB, tables []Table, exportPath string) error {
	// 确保导出目录存在
	exportDir := filepath.Dir(exportPath)
	if err := os.MkdirAll(exportDir, 0755); err != nil {
		return fmt.Errorf("创建导出目录失败: %v", err)
	}

	// 创建Excel文件
	xlsxFile := xlsx.NewFile()

	for _, tableConfig := range tables {
		sheet, err := xlsxFile.AddSheet(tableConfig.TableName)
		if err != nil {
			return fmt.Errorf("创建工作表 %s 失败: %v", tableConfig.TableName, err)
		}

		// 查询表数据
		query := fmt.Sprintf(`SELECT source_file, processed_time`)

		// 添加配置中定义的其他列
		for _, col := range tableConfig.Columns {
			if col.FilenameSource != "" || col.FilenameCell != "" {
				continue
			}
			query += fmt.Sprintf(", %s", col.DBCol)
		}

		query += fmt.Sprintf(` FROM "%s" ORDER BY id`, tableConfig.TableName)

		rows, err := db.Query(query)
		if err != nil {
			return fmt.Errorf("查询表 %s 数据失败: %v", tableConfig.TableName, err)
		}
		defer rows.Close()

		// 获取列信息
		columns, err := rows.Columns()
		if err != nil {
			return fmt.Errorf("获取列信息失败: %v", err)
		}

		// 创建标题行
		headerRow := sheet.AddRow()
		for _, colName := range columns {
			headerRow.AddCell().SetValue(colName)
		}

		// 遍历结果并添加到Excel
		for rows.Next() {
			row := sheet.AddRow()
			values := make([]interface{}, len(columns))
			valuePtrs := make([]interface{}, len(columns))

			for i := range values {
				valuePtrs[i] = &values[i]
			}

			err = rows.Scan(valuePtrs...)
			if err != nil {
				return fmt.Errorf("扫描行数据失败: %v", err)
			}

			for _, val := range values {
				cell := row.AddCell()

				// 处理不同类型的值
				switch v := val.(type) {
				case []byte:
					cell.SetValue(string(v))
				case time.Time:
					cell.SetValue(v.Format("2006-01-02 15:04:05"))
				default:
					cell.SetValue(val)
				}
			}
		}
	}

	// 保存Excel文件
	err := xlsxFile.Save(exportPath)
	if err != nil {
		return fmt.Errorf("保存Excel文件失败: %v", err)
	}

	return nil
}

// getXlsxCellCoords 将Excel单元格坐标转换为行列索引（从0开始）
func getXlsxCellCoords(cell string) (col, row int) {
	col = 0
	i := 0
	for ; i < len(cell) && cell[i] >= 'A' && cell[i] <= 'Z'; i++ {
		col = col*26 + int(cell[i]-'A'+1)
	}
	col-- // 转换为从0开始的索引

	rowStr := cell[i:]
	fmt.Sscanf(rowStr, "%d", &row)
	row-- // 转换为从0开始的索引

	return col, row
}

// loadConfig 加载配置文件
func loadConfig(path string) (*Config, error) {
	data, err := ioutil.ReadFile(path)
	if err != nil {
		return nil, err
	}

	var config Config
	err = yaml.Unmarshal(data, &config)
	if err != nil {
		return nil, err
	}

	// 设置默认值
	if config.BatchSize <= 0 {
		config.BatchSize = 100
	}

	return &config, nil
}

// getXlsxFiles 获取指定目录下的所有xlsx文件
func getXlsxFiles(dir string) ([]XlsxFileInfo, error) {
	var files []XlsxFileInfo

	err := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		if !info.IsDir() && (strings.HasSuffix(strings.ToLower(info.Name()), ".xlsx") ||
			strings.HasSuffix(strings.ToLower(info.Name()), ".xls")) {
			files = append(files, XlsxFileInfo{
				Name: info.Name(),
				Path: path,
			})
		}

		return nil
	})

	return files, err
}

// createTable 创建数据库表
func createTable(db *sql.DB, tableConfig *Table) error {
	// 构建列定义
	columnDefs := []string{}

	// 添加文件名字段
	columnDefs = append(columnDefs, `"source_file" TEXT`)

	// 添加处理时间字段
	columnDefs = append(columnDefs, `"processed_time" DATETIME`)

	// 添加配置中定义的其他列
	for _, col := range tableConfig.Columns {
		// 跳过filename相关的配置项，因为它们是特殊处理的
		if col.FilenameSource != "" || col.FilenameCell != "" {
			continue
		}
		columnDefs = append(columnDefs, fmt.Sprintf(`"%s" %s`, col.DBCol, col.DBType))
	}

	query := fmt.Sprintf(`CREATE TABLE IF NOT EXISTS "%s" (
		"id" INTEGER PRIMARY KEY AUTOINCREMENT,
		%s
	)`, tableConfig.TableName, strings.Join(columnDefs, ",\n		"))

	_, err := db.Exec(query)
	return err
}

// batchInsertData 批量插入数据
func batchInsertData(db *sql.DB, tableConfig *Table, rows [][]string, sourceFileName string, batchSize int) error {
	if len(rows) == 0 {
		return nil
	}

	// 准备插入语句 - 只包括实际的列，排除filename相关的配置
	columnNames := []string{"source_file", "processed_time"}
	valuePlaceholders := []string{"?", "?"}

	for _, col := range tableConfig.Columns {
		if col.FilenameSource != "" || col.FilenameCell != "" {
			continue
		}
		// 确保列名不是空字符串
		if col.DBCol != "" {
			columnNames = append(columnNames, fmt.Sprintf(`"%s"`, col.DBCol))
			valuePlaceholders = append(valuePlaceholders, "?")
		}
	}

	// 如果没有有效的列，跳过插入
	if len(columnNames) <= 2 { // 只有source_file和processed_time
		fmt.Printf("没有找到有效的数据列，跳过插入\n")
		return nil
	}

	insertQuery := fmt.Sprintf(
		`INSERT INTO "%s" (%s) VALUES (%s)`,
		tableConfig.TableName,
		strings.Join(columnNames, ", "),
		strings.Join(valuePlaceholders, ", "),
	)

	tx, err := db.Begin()
	if err != nil {
		return err
	}
	defer tx.Rollback()

	stmt, err := tx.Prepare(insertQuery)
	if err != nil {
		return err
	}
	defer stmt.Close()

	now := time.Now().Format("2006-01-02 15:04:05")
	batchCount := 0

	for _, row := range rows {
		values := []interface{}{sourceFileName, now}

		for _, col := range tableConfig.Columns {
			if col.FilenameSource != "" || col.FilenameCell != "" {
				continue
			}

			// 确保列名不是空字符串
			if col.DBCol != "" {
				// 根据列索引获取单元格值
				colIndex := getColumnIndex(col.ExcelCol)
				if colIndex < len(row) {
					values = append(values, strings.TrimSpace(row[colIndex]))
				} else {
					values = append(values, "")
				}
			}
		}

		_, err := stmt.Exec(values...)
		if err != nil {
			return fmt.Errorf("执行插入语句失败: %v", err)
		}

		batchCount++
		if batchCount >= batchSize {
			err = tx.Commit()
			if err != nil {
				return err
			}

			tx, err = db.Begin()
			if err != nil {
				return err
			}
			defer tx.Rollback()

			stmt.Close()
			stmt, err = tx.Prepare(insertQuery)
			if err != nil {
				return err
			}
			defer stmt.Close()

			batchCount = 0
		}
	}

	// 提交剩余的数据
	if batchCount > 0 {
		err = tx.Commit()
		if err != nil {
			return err
		}
	}

	return nil
}

// getColumnIndex 将Excel列字母转换为索引（从0开始）
func getColumnIndex(colName string) int {
	result := 0
	for _, c := range strings.ToUpper(colName) {
		if c < 'A' || c > 'Z' {
			panic(fmt.Sprintf("无效的Excel列名: %s", colName))
		}
		result = result*26 + int(c-'A'+1)
	}
	return result - 1
}
