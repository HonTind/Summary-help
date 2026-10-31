import os
import sys
from .config import APP_FULL_NAME, FOLDER_NAME, OUTPUT_FILE, OUTPUT_FOLDER_NAME, ERROR_LOG
from .utils import get_base_dir, get_config_dir, get_logs_dir, get_unique_file_path
from .field_mapper import load_settings, load_field_config
from .excel_reader import read_and_aggregate
from .processor import sort_by_amount
from .excel_writer import generate_summary


def main():
    # ========== 1. 路径设置 ==========
    base_dir = get_base_dir()
    config_dir = get_config_dir()
    logs_dir = get_logs_dir()
    load_settings(config_dir)
    load_field_config(config_dir)
    folder = os.path.join(base_dir, FOLDER_NAME)

    # ========== 2. 异常处理：文件夹检查 ==========
    if not os.path.isdir(folder):
        print("=" * 35)
        print('错误：未找到 [对账单] 文件夹')
        print('请确保在程序所在目录创建 [对账单] 文件夹')
        print("=" * 35)
        input("按 Enter 键退出...")
        sys.exit(1)

    # ========== 3. 读取并汇总数据（v1.5.0 动态字段）==========
    result = read_and_aggregate(folder, logs_dir)
    total, file_count, field_names = result

    if file_count == 0:
        print("=" * 35)
        print('错误：[对账单] 文件夹中没有发现 Excel 文件')
        print('请将 .xlsx 文件放入 [对账单] 文件夹后重试')
        print("=" * 35)
        input("按 Enter 键退出...")
        sys.exit(1)

    # ========== 4. 排序 ==========
    sorted_data = sort_by_amount(total, field_names)

    # ========== 5. 生成汇总 Excel（v1.5.1: 默认输出目录 + 防覆盖）==========
    output_dir = os.path.join(base_dir, OUTPUT_FOLDER_NAME)
    output_path = get_unique_file_path(output_dir, OUTPUT_FILE)
    try:
        generate_summary(sorted_data, total, output_path, field_names)
    except PermissionError:
        print("错误：请关闭正在打开的汇总.xlsx 后重新运行")
        input("按 Enter 键退出...")
        sys.exit(1)
    except Exception as e:
        print(f"错误：保存失败 - {e}")
        with open(os.path.join(logs_dir, ERROR_LOG), "a", encoding="utf-8") as logf:
            logf.write(f"保存失败 {output_path}：{repr(e)}\n")
        input("按 Enter 键退出...")
        sys.exit(1)

    # ========== 6. 运行结果提示（v1.5.0 动态字段）==========
    sum_field = field_names[1] if len(field_names) >= 2 else field_names[0]
    sum_total = sum(item.get(sum_field, 0) for item in total.values())

    print("=" * 35)
    print(f"   {APP_FULL_NAME}")
    print("=" * 35)
    print(f"  读取Excel：{file_count} 个")
    print(f"  分组数量：{len(total)} 个")
    print(f"  汇总{sum_field}：{sum_total:,.2f} 元")
    print(f"  生成文件：{output_path}")
    print("=" * 35)

    input("按 Enter 键退出...")


if __name__ == "__main__":
    main()
