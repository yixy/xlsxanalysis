# 变量定义
BINARY_NAME=xlsxanalysis
MAIN_PACKAGE=./cmd/xlsxanalysis
GO_CMD=go1.20.14
# CGO_ENABLED=0 GOOS=windows GOARCH=amd64 GOAMD64=v1 go build -ldflags="-s -w" -trimpath
GO_BUILD=$(GO_CMD) build -ldflags="-s -w" -trimpath
GO_CLEAN=$(GO_CMD) clean
GO_TEST=$(GO_CMD) test
GO_TIDY=$(GO_CMD) mod tidy
TARGET_DIR=target
RESOURCES=config xlsx_files

.PHONY: all build build-linux build-windows build-darwin build-all clean test run tidy help

# 默认目标：执行编译
all: build

# 编译项目
build:
	@echo "正在编译 $(BINARY_NAME)..."
	@mkdir -p $(TARGET_DIR)
	@cp -r $(RESOURCES) $(TARGET_DIR)/
	$(GO_BUILD) -o $(TARGET_DIR)/$(BINARY_NAME) $(MAIN_PACKAGE)
	@echo "编译完成。"

# 交叉编译 Linux (amd64)
build-linux:
	@echo "正在编译 Linux 版本..."
	@mkdir -p $(TARGET_DIR)/linux
	@cp -r $(RESOURCES) $(TARGET_DIR)/linux/
	CGO_ENABLED=0 GOOS=linux GOARCH=amd64 $(GO_BUILD) -o $(TARGET_DIR)/linux/$(BINARY_NAME) $(MAIN_PACKAGE)

# 交叉编译 Windows (amd64) - 兼容 Go 1.20 和 Windows 7
# 注意：Go 1.21+ 不再支持 Windows 7。此版本使用 Go 1.20，可以支持 Win7。
build-windows:
	@echo "正在编译 Windows 版本 (兼容 Windows 7)..."
	@mkdir -p $(TARGET_DIR)/windows
	@cp -r $(RESOURCES) $(TARGET_DIR)/windows/
	CGO_ENABLED=0 GOOS=windows GOARCH=amd64 $(GO_BUILD) -o $(TARGET_DIR)/windows/$(BINARY_NAME).exe $(MAIN_PACKAGE)

# 交叉编译 Windows 32 位版本 (386) - 兼容 Windows 7 32 位系统
build-windows-386:
	@echo "正在编译 Windows 32 位版本 (兼容 Windows 7 32 位)..."
	@mkdir -p $(TARGET_DIR)/windows_386
	@cp -r $(RESOURCES) $(TARGET_DIR)/windows_386/
	CGO_ENABLED=0 GOOS=windows GOARCH=386 $(GO_BUILD) -o $(TARGET_DIR)/windows_386/$(BINARY_NAME).exe $(MAIN_PACKAGE)

# 交叉编译 Darwin (macOS amd64)
build-darwin:
	@echo "正在编译 macOS 版本..."
	@mkdir -p $(TARGET_DIR)/darwin
	@cp -r $(RESOURCES) $(TARGET_DIR)/darwin/
	CGO_ENABLED=0 GOOS=darwin GOARCH=amd64 $(GO_BUILD) -o $(TARGET_DIR)/darwin/$(BINARY_NAME) $(MAIN_PACKAGE)

# 编译所有平台
build-all: build-linux build-windows build-windows-386 build-darwin

# 清理编译产物
clean:
	@echo "正在清理..."
	$(GO_CLEAN)
	rm -rf $(TARGET_DIR)
	@echo "清理完毕。"

# 运行测试
test:
	$(GO_TEST) -v ./...

# 编译并运行
run: build
	./$(TARGET_DIR)/$(BINARY_NAME)

# 整理依赖
tidy:
	$(GO_TIDY)

# 帮助说明
help:
	@echo "使用说明:"
	@echo "  make build  - 编译项目二进制文件"
	@echo "  make clean  - 移除编译生成的二进制文件"
	@echo "  make test   - 执行所有单元测试"
	@echo "  make run    - 编译并直接运行程序"
	@echo "  make tidy   - 整理 go.mod 依赖"
	@echo "  make build-linux        - 编译 Linux 版本"
	@echo "  make build-windows      - 编译 Windows 64 位版本 (兼容 Win7)"
	@echo "  make build-windows-386  - 编译 Windows 32 位版本 (兼容 Win7)"
	@echo "  make build-darwin       - 编译 macOS 版本"
	@echo "  make build-all          - 编译所有平台版本"