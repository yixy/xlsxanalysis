# 变量定义
BINARY_NAME=xlsxanalysis
MAIN_PACKAGE=./cmd/xlsxanalysis
GO_CMD=go
GO_BUILD=$(GO_CMD) build
GO_CLEAN=$(GO_CMD) clean
GO_TEST=$(GO_CMD) test
GO_TIDY=$(GO_CMD) mod tidy

.PHONY: all build build-linux build-windows build-darwin build-all clean test run tidy help

# 默认目标：执行编译
all: build

# 编译项目
build:
	@echo "正在编译 $(BINARY_NAME)..."
	mkdir -p target
	cp -r config target/
	cp -r xlsx_files target/

	$(GO_BUILD) -o target/$(BINARY_NAME) $(MAIN_PACKAGE)
	@echo "编译完成。"

# 交叉编译 Linux (amd64)
build-linux:
	@echo "正在编译 Linux 版本..."
	mkdir -p target/linux
	cp -r config target/linux/
	cp -r xlsx_files target/linux/
	CGO_ENABLED=0 GOOS=linux GOARCH=amd64 $(GO_BUILD) -o target/linux/$(BINARY_NAME) $(MAIN_PACKAGE)

# 交叉编译 Windows (amd64)
build-windows:
	@echo "正在编译 Windows 版本..."
	mkdir -p target/windows
	cp -r config target/windows/
	cp -r xlsx_files target/windows/
	CGO_ENABLED=0 GOOS=windows GOARCH=amd64 $(GO_BUILD) -o target/windows/$(BINARY_NAME).exe $(MAIN_PACKAGE)

# 交叉编译 Darwin (macOS amd64)
build-darwin:
	@echo "正在编译 macOS 版本..."
	mkdir -p target/darwin
	cp -r config target/darwin/
	cp -r xlsx_files target/darwin/
	CGO_ENABLED=0 GOOS=darwin GOARCH=amd64 $(GO_BUILD) -o target/darwin/$(BINARY_NAME) $(MAIN_PACKAGE)

# 编译所有平台
build-all: build-linux build-windows build-darwin

# 清理编译产物
clean:
	@echo "正在清理..."
	$(GO_CLEAN)
	rm -rf target/*
	@echo "清理完毕。"

# 运行测试
test:
	$(GO_TEST) -v ./...

# 编译并运行
run: build
	./$(BINARY_NAME)

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
	@echo "  make build-linux   - 编译 Linux 版本"
	@echo "  make build-windows - 编译 Windows 版本"
	@echo "  make build-darwin  - 编译 macOS 版本"
	@echo "  make build-all     - 编译所有平台版本"
