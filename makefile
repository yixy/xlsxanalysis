# 变量定义
BINARY_NAME=xlsxanalysis
MAIN_PACKAGE=./cmd/xlsxanalysis
GO_CMD=go
GO_BUILD=$(GO_CMD) build
GO_CLEAN=$(GO_CMD) clean
GO_TEST=$(GO_CMD) test
GO_TIDY=$(GO_CMD) mod tidy

.PHONY: all build clean test run tidy help

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
