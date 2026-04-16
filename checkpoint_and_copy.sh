#!/bin/bash

# SQLite WAL Checkpoint 并备份数据库脚本
# 功能：将 WAL 文件数据写入主数据库，然后复制到目标目录

set -e

# 配置
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DB_DIR="${SCRIPT_DIR}/data"
DB_NAME="benchmark.db"
DB_PATH="${DB_DIR}/${DB_NAME}"
WAL_PATH="${DB_DIR}/${DB_NAME}-wal"
SHM_PATH="${DB_DIR}/${DB_NAME}-shm"
TARGET_DIR="/mnt/d/Data"

# 颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

log_info() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

log_warn() {
    echo -e "${YELLOW}[WARN]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# 检查 sqlite3 是否安装
check_sqlite3() {
    if ! command -v sqlite3 &> /dev/null; then
        log_error "sqlite3 未安装，请先安装 sqlite3"
        exit 1
    fi
}

# 检查数据库文件是否存在
check_db_files() {
    if [[ ! -f "$DB_PATH" ]]; then
        log_error "数据库文件不存在: $DB_PATH"
        exit 1
    fi
    log_info "找到数据库文件: $DB_PATH"

    if [[ -f "$WAL_PATH" ]]; then
        WAL_SIZE=$(du -h "$WAL_PATH" | cut -f1)
        log_info "找到 WAL 文件: $WAL_PATH (大小: $WAL_SIZE)"
    else
        log_warn "WAL 文件不存在: $WAL_PATH"
    fi
}

# 检查目标目录
check_target_dir() {
    if [[ ! -d "$TARGET_DIR" ]]; then
        log_error "目标目录不存在: $TARGET_DIR"
        exit 1
    fi
    log_info "目标目录存在: $TARGET_DIR"
}

# 执行 WAL checkpoint
do_checkpoint() {
    log_info "正在执行 WAL checkpoint..."

    # 使用 PRAGMA wal_checkpoint(TRUNCATE) 将 WAL 内容写入数据库并截断 WAL 文件
    # TRUNCATE 模式会阻塞其他写入操作，确保完整 checkpoint
    result=$(sqlite3 "$DB_PATH" "PRAGMA wal_checkpoint(TRUNCATE);" 2>&1)

    if [[ $? -eq 0 ]]; then
        log_info "WAL checkpoint 完成"

        # 显示 checkpoint 结果 (busy, log, checkpointed)
        if [[ -n "$result" ]]; then
            log_info "Checkpoint 结果: $result"
        fi
    else
        log_error "WAL checkpoint 失败: $result"
        exit 1
    fi

    # 验证 WAL 文件是否已被处理
    if [[ -f "$WAL_PATH" ]]; then
        WAL_SIZE=$(stat -c%s "$WAL_PATH" 2>/dev/null || echo "0")
        if [[ "$WAL_SIZE" -gt 0 ]]; then
            log_warn "WAL 文件仍存在且有内容 (大小: ${WAL_SIZE} bytes)"
            log_warn "可能有其他进程正在使用数据库"
        else
            log_info "WAL 文件已清空"
        fi
    fi
}

# 复制数据库到目标目录
copy_db() {
    log_info "正在复制数据库到 $TARGET_DIR ..."

    # 生成带时间戳的备份文件名（可选）
    TIMESTAMP=$(date +%Y%m%d_%H%M%S)
    BACKUP_NAME="${DB_NAME}"
    TARGET_PATH="${TARGET_DIR}/${BACKUP_NAME}"

    # 复制文件
    cp -v "$DB_PATH" "$TARGET_PATH"

    if [[ $? -eq 0 ]]; then
        log_info "数据库复制成功: $TARGET_PATH"

        # 显示文件大小
        SRC_SIZE=$(du -h "$DB_PATH" | cut -f1)
        DST_SIZE=$(du -h "$TARGET_PATH" | cut -f1)
        log_info "源文件大小: $SRC_SIZE, 目标文件大小: $DST_SIZE"
    else
        log_error "数据库复制失败"
        exit 1
    fi
}

# 显示完成信息
show_summary() {
    echo ""
    log_info "========== 操作完成 =========="
    log_info "数据库路径: $DB_PATH"
    log_info "备份路径: $TARGET_DIR/$DB_NAME"
    echo ""
}

# 主函数
main() {
    log_info "开始执行数据库 checkpoint 和备份..."
    echo ""

    check_sqlite3
    check_db_files
    check_target_dir
    do_checkpoint
    copy_db
    show_summary
}

# 运行主函数
main "$@"
