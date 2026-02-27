#!/usr/bin/env python3
"""
SQL Dump Splitter
将大型 MySQL dump 文件按指定大小拆分，每个分片都是可独立执行的完整 SQL 文件。
支持在 INSERT 语句之间切割，避免单个超大表导致分片失控。
"""

import os
import sys
import argparse
import re

def parse_size(size_str):
    size_str = size_str.strip().upper()
    units = {'GB': 1024**3, 'MB': 1024**2, 'KB': 1024, 'G': 1024**3, 'M': 1024**2, 'K': 1024, 'B': 1}
    for unit, multiplier in units.items():
        if size_str.endswith(unit):
            return int(float(size_str[:-len(unit)]) * multiplier)
    return int(size_str)

def extract_header(input_file):
    """提取文件头：从开头到第一个 DROP TABLE/CREATE TABLE 之前"""
    header_lines = []
    with open(input_file, 'r', encoding='utf-8', errors='replace') as f:
        for line in f:
            if re.match(r'^(DROP TABLE|CREATE TABLE)', line.strip()):
                break
            header_lines.append(line)
    return ''.join(header_lines)

FOOTER = (
    "\n/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;\n"
    "/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;\n"
    "/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;\n"
    "/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;\n"
    "/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;\n"
    "/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;\n"
    "/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;\n"
    "/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;\n"
)

def split_sql(input_file, output_dir, max_size_bytes, prefix='part'):
    os.makedirs(output_dir, exist_ok=True)

    file_size = os.path.getsize(input_file)
    print(f"[*] 输入文件: {input_file}")
    print(f"[*] 文件大小: {file_size / 1024**2:.1f} MB")
    print(f"[*] 分片上限: {max_size_bytes / 1024**2:.1f} MB")
    print(f"[*] 输出目录: {output_dir}")
    print()

    header = extract_header(input_file)
    header_size = len(header.encode('utf-8'))

    part_num = 1
    part_lines = [header]
    part_size = header_size

    # 当前表的 LOCK/UNLOCK 上下文（用于跨分片的 INSERT 块）
    current_lock_line = None   # "LOCK TABLES `xxx` WRITE;\n"
    current_disable_keys = None  # "/*!40000 ALTER TABLE `xxx` DISABLE KEYS */;\n"
    in_data_block = False      # 是否在 LOCK...UNLOCK 块内

    total_parts = 0

    def flush_part():
        nonlocal part_num, part_lines, part_size, total_parts
        filename = os.path.join(output_dir, f"{prefix}_{part_num:04d}.sql")
        with open(filename, 'w', encoding='utf-8') as f:
            f.writelines(part_lines)
            f.write(FOOTER)
        size_mb = os.path.getsize(filename) / 1024**2
        print(f"  -> {os.path.basename(filename)}  ({size_mb:.1f} MB)")
        part_num += 1
        total_parts += 1
        part_lines = [header]
        part_size = header_size

    def maybe_flush(extra_bytes=0):
        """如果加上 extra_bytes 后超限，先 flush"""
        if part_size + extra_bytes > max_size_bytes and part_size > header_size:
            flush_part()

    with open(input_file, 'r', encoding='utf-8', errors='replace') as f:
        # 跳过 header 部分
        for line in f:
            if re.match(r'^(DROP TABLE|CREATE TABLE)', line.strip()):
                # 这行是正文第一行，交给下面处理
                line_bytes = len(line.encode('utf-8'))
                maybe_flush(line_bytes)
                part_lines.append(line)
                part_size += line_bytes
                break

        for line in f:
            stripped = line.strip()
            line_bytes = len(line.encode('utf-8'))

            # ---- LOCK TABLES：数据块开始 ----
            if stripped.startswith('LOCK TABLES'):
                in_data_block = True
                current_lock_line = line
                current_disable_keys = None
                # 不在 INSERT 中间，可以在此切割
                maybe_flush(line_bytes)
                part_lines.append(line)
                part_size += line_bytes
                continue

            # ---- DISABLE KEYS（紧跟 LOCK 之后）----
            if stripped.startswith('/*!40000 ALTER TABLE') and 'DISABLE KEYS' in stripped:
                current_disable_keys = line
                part_lines.append(line)
                part_size += line_bytes
                continue

            # ---- UNLOCK TABLES：数据块结束 ----
            if stripped.startswith('UNLOCK TABLES'):
                in_data_block = False
                current_lock_line = None
                current_disable_keys = None
                part_lines.append(line)
                part_size += line_bytes
                continue

            # ---- INSERT 语句：在语句之间可以切割 ----
            if stripped.startswith('INSERT INTO') and in_data_block:
                if part_size + line_bytes > max_size_bytes and part_size > header_size:
                    # 在写入新 INSERT 前，先补上 ENABLE KEYS + UNLOCK，然后 flush
                    table_match = re.search(r'LOCK TABLES `(.+?)`', current_lock_line) if current_lock_line else None
                    table_name = table_match.group(1) if table_match else 'unknown'
                    part_lines.append(f"/*!40000 ALTER TABLE `{table_name}` ENABLE KEYS */;\n")
                    part_lines.append("UNLOCK TABLES;\n")
                    flush_part()
                    # 新分片开头重新 LOCK 这张表
                    part_lines.append(current_lock_line)
                    part_size += len(current_lock_line.encode('utf-8'))
                    if current_disable_keys:
                        part_lines.append(current_disable_keys)
                        part_size += len(current_disable_keys.encode('utf-8'))

                part_lines.append(line)
                part_size += line_bytes
                continue

            # ---- 其他语句（DDL、注释等）：在语句边界切割 ----
            if not in_data_block:
                maybe_flush(line_bytes)

            part_lines.append(line)
            part_size += line_bytes

    # 写出最后一个分片
    if part_size > header_size:
        flush_part()

    print()
    print(f"[+] 完成！共 {total_parts} 个分片，保存在: {output_dir}")
    return total_parts

def main():
    parser = argparse.ArgumentParser(
        description='SQL Dump Splitter - 将大型 MySQL dump 拆分为多个可独立导入的小文件',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python split_sql.py dump.sql                         # 默认 100MB 拆分
  python split_sql.py dump.sql -s 50MB                 # 按 50MB 拆分
  python split_sql.py dump.sql -s 50MB -o ./output     # 指定输出目录
  python split_sql.py dump.sql -s 50MB -p chunk        # 自定义文件名前缀

导入方式（按顺序逐个导入）:
  mysql -u root -p --default-character-set=utf8mb4 < part_0001.sql
  mysql -u root -p --default-character-set=utf8mb4 < part_0002.sql
  ...

或批量导入（Linux/macOS）:
  for f in split_output/part_*.sql; do mysql -u root -p < "$f"; done
        """
    )
    parser.add_argument('input', help='输入的 SQL dump 文件路径')
    parser.add_argument('-s', '--size', default='100MB',
                        help='每个分片的最大大小，支持 B/KB/MB/GB，默认 100MB')
    parser.add_argument('-o', '--output', default=None,
                        help='输出目录，默认为输入文件同目录下的 split_output/')
    parser.add_argument('-p', '--prefix', default='part',
                        help='输出文件名前缀，默认 part')

    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"[!] 错误：文件不存在: {args.input}")
        sys.exit(1)

    max_size = parse_size(args.size)
    if max_size <= 0:
        print("[!] 错误：分片大小必须大于 0")
        sys.exit(1)

    if args.output is None:
        base_dir = os.path.dirname(os.path.abspath(args.input))
        args.output = os.path.join(base_dir, 'split_output')

    split_sql(args.input, args.output, max_size, args.prefix)

if __name__ == '__main__':
    main()
