.PHONY: help clean package package-windows

BINARY_NAME=ttcn-report
OUTPUT_DIR=output

help:
	@echo "可用命令:"
	@echo "  make help             显示帮助信息"
	@echo "  make clean            清理构建产物"
	@echo "  make package          构建可执行文件"
	@echo "  make package-windows  构建 Windows x86_64 可执行文件"

clean:
	@echo "清理构建产物..."
	@rm -rf $(OUTPUT_DIR)

package:
	@echo "构建可执行文件..."
	@mkdir -p $(OUTPUT_DIR)
	@go build -ldflags "-s -w" -o $(OUTPUT_DIR)/$(BINARY_NAME) .

package-windows:
	@echo "构建 Windows x86_64 可执行文件..."
	@mkdir -p $(OUTPUT_DIR)
	@GOOS=windows GOARCH=amd64 go build -ldflags "-s -w -H=windowsgui" -o $(OUTPUT_DIR)/$(BINARY_NAME).exe .
