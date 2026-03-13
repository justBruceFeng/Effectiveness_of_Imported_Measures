# -*- coding: utf-8 -*-
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import warnings
import zipfile
import re
from reliability.Fitters import (
    Fit_Weibull_2P, Fit_Weibull_3P, Fit_Exponential_2P, 
    Fit_Normal_2P, Fit_Lognormal_2P, Fit_Gamma_2P,
    Fit_Gumbel_2P, Fit_Weibull_Mixture, Fit_Weibull_CR,
    Fit_Loglogistic_2P, Fit_Everything
)
from reliability.Probability_plotting import plot_points
from reliability.Distributions import (
    Weibull_Distribution, Normal_Distribution, Lognormal_Distribution,
    Exponential_Distribution, Gamma_Distribution, Gumbel_Distribution,
    Loglogistic_Distribution, Mixture_Model, Competing_Risks_Model
)

# 尝试导入tqdm用于进度条
try:
    from tqdm import tqdm
    TQDM_AVAILABLE = True
except ImportError:
    TQDM_AVAILABLE = False

warnings.filterwarnings('ignore')

# 默认配置参数
DEFAULT_ANALYSIS_MONTHS = 24  # 默认分析时长（月）
DEFAULT_INTERVAL_MONTHS = 3    # 默认分析间隔（月）
DEFAULT_IMPORT_DATE = '2025-06-27'  # 默认改善措施导入时间
DEFAULT_OUTPUT_DIR = "改善对比分析结果"  # 默认输出目录

def get_user_input():
    """获取用户输入"""
    import sys
    
    # 检查是否有命令行参数 --auto 或 -a
    if len(sys.argv) > 1 and (sys.argv[1] == '--auto' or sys.argv[1] == '-a'):
        print("检测到 --auto 参数，使用默认参数")
        print("=" * 60)
        print("可靠性分析系统 - 自动配置")
        print("=" * 60)
        
        # 使用默认值
        import_date = DEFAULT_IMPORT_DATE
        analysis_months = DEFAULT_ANALYSIS_MONTHS
        interval_months = DEFAULT_INTERVAL_MONTHS
        output_dir = DEFAULT_OUTPUT_DIR
        file_path = 'T920-变形问题-验算.xlsx'
        
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        # 打印配置信息
        print(f"使用默认导入时间: {import_date}")
        print(f"使用默认分析时长: {analysis_months} 个月")
        print(f"使用默认时间间隔: {interval_months} 个月")
        print(f"输出目录设置为: {output_dir}")
        print(f"使用默认文件路径: {file_path}")
        print(f"数据文件: {file_path}")
        print("=" * 60)
        
        return {
            'import_date': import_date,
            'analysis_months': analysis_months,
            'interval_months': interval_months,
            'output_dir': output_dir,
            'file_path': file_path
        }
    
    # 检查是否在非交互式环境中运行
    if not sys.stdin.isatty():
        print("检测到非交互式环境，使用默认参数")
        print("=" * 60)
        print("可靠性分析系统 - 自动配置")
        print("=" * 60)
        
        # 使用默认值
        import_date = DEFAULT_IMPORT_DATE
        analysis_months = DEFAULT_ANALYSIS_MONTHS
        interval_months = DEFAULT_INTERVAL_MONTHS
        output_dir = DEFAULT_OUTPUT_DIR
        file_path = 'T920-变形问题-验算.xlsx'
        
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        # 打印配置信息
        print(f"使用默认导入时间: {import_date}")
        print(f"使用默认分析时长: {analysis_months} 个月")
        print(f"使用默认时间间隔: {interval_months} 个月")
        print(f"输出目录设置为: {output_dir}")
        print(f"使用默认文件路径: {file_path}")
        print(f"数据文件: {file_path}")
        print("=" * 60)
        
        return {
            'import_date': import_date,
            'analysis_months': analysis_months,
            'interval_months': interval_months,
            'output_dir': output_dir,
            'file_path': file_path
        }
    
    # 交互式环境
    print("=" * 60)
    print("可靠性分析系统 - 用户配置")
    print("=" * 60)
    
    # 获取改善措施导入时间
    while True:
        import_date = input(f"请输入改善措施导入时间 (格式: YYYY-MM-DD, 默认: {DEFAULT_IMPORT_DATE}): ").strip()
        if not import_date:
            import_date = DEFAULT_IMPORT_DATE
            print(f"使用默认导入时间: {import_date}")
            break
        
        # 验证日期格式
        if validate_date_format(import_date):
            print(f"导入时间设置为: {import_date}")
            break
        else:
            print("错误: 日期格式不正确，请使用 YYYY-MM-DD 格式")
    
    # 获取分析时长（月）
    while True:
        analysis_months_str = input(f"请输入分析时长 (月, 默认: {DEFAULT_ANALYSIS_MONTHS}): ").strip()
        if not analysis_months_str:
            analysis_months = DEFAULT_ANALYSIS_MONTHS
            print(f"使用默认分析时长: {analysis_months} 个月")
            break
        
        try:
            analysis_months = int(analysis_months_str)
            if 3 <= analysis_months <= 60:  # 3个月到5年
                print(f"分析时长设置为: {analysis_months} 个月")
                break
            else:
                print("错误: 分析时长应在3-60个月之间")
        except ValueError:
            print("错误: 请输入有效的整数")
    
    # 获取时间间隔（月）
    while True:
        interval_months_str = input(f"请输入分析间隔 (月, 默认: {DEFAULT_INTERVAL_MONTHS}): ").strip()
        if not interval_months_str:
            interval_months = DEFAULT_INTERVAL_MONTHS
            print(f"使用默认时间间隔: {interval_months} 个月")
            break
        
        try:
            interval_months = int(interval_months_str)
            if interval_months in [1, 2, 3, 6, 12]:
                print(f"时间间隔设置为: {interval_months} 个月")
                break
            else:
                print("错误: 时间间隔应为1, 2, 3, 6或12个月")
        except ValueError:
            print("错误: 请输入有效的整数")
    
    # 获取输出目录
    while True:
        output_dir = input(f"请输入输出目录 (默认: '{DEFAULT_OUTPUT_DIR}'): ").strip()
        if not output_dir:
            output_dir = DEFAULT_OUTPUT_DIR
        
        # 检查目录是否存在，如果不存在则询问是否创建
        if not os.path.exists(output_dir):
            create_dir = input(f"目录 '{output_dir}' 不存在，是否创建? (y/n, 默认: y): ").strip().lower()
            if not create_dir or create_dir == 'y':
                try:
                    os.makedirs(output_dir, exist_ok=True)
                    print(f"目录 '{output_dir}' 已创建")
                    break
                except Exception as e:
                    print(f"创建目录失败: {e}")
                    print("请重新输入或使用默认目录")
            else:
                print("请重新输入输出目录")
        else:
            print(f"输出目录设置为: {output_dir}")
            break
    
    # 获取数据文件路径
    while True:
        file_path = input("请输入Excel数据文件路径 (例如: T920-变形问题-验算.xlsx): ").strip()
        if not file_path:
            file_path = 'T920-变形问题-验算.xlsx'
            print(f"使用默认文件路径: {file_path}")
        
        if os.path.exists(file_path):
            print(f"数据文件: {file_path}")
            break
        else:
            print(f"错误: 文件 '{file_path}' 不存在")
            retry = input("是否重新输入文件路径? (y/n, 默认: y): ").strip().lower()
            if retry and retry != 'y':
                print("使用默认文件路径")
                break
    
    print("=" * 60)
    
    return {
        'import_date': import_date,
        'analysis_months': analysis_months,
        'interval_months': interval_months,
        'output_dir': output_dir,
        'file_path': file_path
    }

def validate_date_format(date_str):
    """验证日期格式是否为YYYY-MM-DD"""
    pattern = r'^\d{4}-\d{2}-\d{2}$'
    if not re.match(pattern, date_str):
        return False
    
    try:
        datetime.strptime(date_str, '%Y-%m-%d')
        return True
    except ValueError:
        return False

def generate_time_points(analysis_months, interval_months):
    """生成时间点（月）"""
    time_points = []
    current_months = 0  # 从0开始
    
    while current_months <= analysis_months:
        time_points.append(current_months)
        current_months += interval_months
    
    return time_points

def days_to_months(days):
    """将天数转换为月份"""
    return round(days / 30.44)

plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

def is_valid_excel_file(file_path):
    """检查是否为有效的Excel文件"""
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            return False, "文件不存在"
        
        # 检查文件大小
        file_size = os.path.getsize(file_path)
        if file_size == 0:
            return False, "文件大小为0"
        
        # 检查文件扩展名
        if not file_path.lower().endswith(('.xlsx', '.xls')):
            return False, "文件扩展名不是.xlsx或.xls"
        
        # 检查是否为临时文件
        if os.path.basename(file_path).startswith('~$'):
            return False, "文件是Excel临时文件"
        
        # 尝试检查.xlsx文件的ZIP结构
        if file_path.lower().endswith('.xlsx'):
            try:
                with zipfile.ZipFile(file_path, 'r') as zip_file:
                    if 'xl/workbook.xml' not in zip_file.namelist():
                        return False, "不是有效的OOXML格式文件"
            except zipfile.BadZipFile:
                return False, "ZIP文件结构损坏"
        
        return True, "文件有效"
    except Exception as e:
        return False, f"文件检查失败: {str(e)}"

def repair_excel_file(input_path, output_path):
    """尝试修复Excel文件"""
    try:
        print(f"尝试修复文件: {input_path} -> {output_path}")
        
        # 方法1: 使用不同的引擎尝试读取并重新保存
        engines_to_try = ['openpyxl', 'xlrd']
        
        for engine in engines_to_try:
            try:
                df = pd.read_excel(input_path, engine=engine)
                
                # 成功读取，保存修复后的文件
                df.to_excel(output_path, index=False, engine='openpyxl')
                print(f"使用引擎 {engine} 修复成功")
                return True
            except Exception as e:
                print(f"引擎 {engine} 修复失败: {e}")
                continue
        
        # 方法2: 尝试读取所有sheet
        try:
            excel_file = pd.ExcelFile(input_path)
            sheet_names = excel_file.sheet_names
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name in sheet_names:
                    try:
                        df = pd.read_excel(input_path, sheet_name=sheet_name)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                    except Exception as e:
                        print(f"读取sheet {sheet_name} 失败: {e}")
                        # 创建空的数据框作为占位符
                        pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)
            
            print("通过分sheet读取修复成功")
            return True
        except Exception as e:
            print(f"分sheet读取修复失败: {e}")
        
        return False
    except Exception as e:
        print(f"修复过程出错: {e}")
        return False

def load_reliability_data_multisheet_robust(file_path, sheet_names=None, required_columns=None):
    """
    健壮地读取包含多个sheet的可靠性分析数据
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")

    print(f"正在读取文件: {file_path}")
    
    # 首先检查文件有效性
    is_valid, message = is_valid_excel_file(file_path)
    if not is_valid:
        print(f"文件有效性检查失败: {message}")
        print("尝试修复文件...")
        
        # 创建修复后的文件路径
        repaired_path = file_path.replace('.xlsx', '_repaired.xlsx')
        
        if repair_excel_file(file_path, repaired_path):
            print(f"文件修复成功，使用修复后的文件: {repaired_path}")
            file_path = repaired_path
        else:
            raise ValueError(f"文件损坏且修复失败: {message}")
    
    # 读取所有sheet或指定sheet
    if sheet_names is None:
        try:
            # 获取所有sheet名称
            excel_file = pd.ExcelFile(file_path)
            all_sheet_names = excel_file.sheet_names
            print(f"检测到 {len(all_sheet_names)} 个sheet: {', '.join(all_sheet_names)}")
            excel_file.close()
            
            # 使用进度条读取每个sheet
            all_sheets = {}
            if TQDM_AVAILABLE:
                sheet_iterator = tqdm(all_sheet_names, desc="读取Sheet进度")
            else:
                sheet_iterator = all_sheet_names
                print("开始读取各Sheet数据...")
            
            for sheet_name in sheet_iterator:
                try:
                    # 尝试不同的引擎
                    engines = ['openpyxl', 'xlrd']
                    df = None
                    
                    for engine in engines:
                        try:
                            df = pd.read_excel(file_path, sheet_name=sheet_name, engine=engine)
                            print(f"  使用引擎 {engine} 成功读取: {sheet_name}")
                            break
                        except Exception as e:
                            print(f"  引擎 {engine} 读取失败: {e}")
                            continue
                    
                    if df is not None:
                        all_sheets[sheet_name] = df
                        if not TQDM_AVAILABLE:
                            print(f"  已读取: {sheet_name} ({len(df)}行)")
                    else:
                        print(f"  所有引擎都无法读取sheet: {sheet_name}")
                        
                except Exception as e:
                    print(f"读取sheet '{sheet_name}' 失败: {e}")
        except Exception as e:
            print(f"读取Excel文件结构失败: {e}")
            # 尝试直接读取第一个sheet
            try:
                print("尝试直接读取第一个sheet...")
                df = pd.read_excel(file_path, sheet_name=0, engine='openpyxl')
                all_sheets = {'Sheet1': df}
                print(f"成功读取第一个sheet: {len(df)}行")
            except Exception as e2:
                print(f"直接读取也失败: {e2}")
                return None
    else:
        all_sheets = {}
        if TQDM_AVAILABLE:
            sheet_iterator = tqdm(sheet_names, desc="读取指定Sheet进度")
        else:
            sheet_iterator = sheet_names
            print("开始读取指定Sheet数据...")
            
        for sheet_name in sheet_iterator:
            try:
                # 尝试不同的引擎
                engines = ['openpyxl', 'xlrd']
                df = None
                
                for engine in engines:
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, engine=engine)
                        break
                    except:
                        continue
                
                if df is not None:
                    all_sheets[sheet_name] = df
                    if not TQDM_AVAILABLE:
                        print(f"  已读取: {sheet_name} ({len(df)}行)")
                else:
                    print(f"  无法读取sheet: {sheet_name}")
            except Exception as e:
                print(f"读取sheet '{sheet_name}' 失败: {e}")
    
    if not all_sheets:
        raise ValueError("未找到任何可用的sheet数据")
    
    # 验证每个sheet的列结构
    if required_columns:
        valid_sheets = {}
        for sheet_name, df in all_sheets.items():
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                print(f"警告: sheet '{sheet_name}' 缺少必需列: {missing_columns}")
                # 尝试查找相似的列名
                available_columns = df.columns.tolist()
                print(f"  可用列: {available_columns}")
                
                # 检查是否有相似的列名（大小写、空格差异）
                for required_col in required_columns:
                    if required_col not in df.columns:
                        # 查找忽略大小写和空格的匹配
                        lower_required = required_col.lower().replace(' ', '')
                        matches = [col for col in available_columns 
                                  if col.lower().replace(' ', '') == lower_required]
                        if matches:
                            print(f"  找到相似列: {matches[0]} -> {required_col}")
                            df[required_col] = df[matches[0]]
                            missing_columns.remove(required_col)
                
                if not missing_columns:
                    valid_sheets[sheet_name] = df
                    print(f"  sheet '{sheet_name}' 列结构修复后验证通过")
            else:
                valid_sheets[sheet_name] = df
                print(f"sheet '{sheet_name}' 列结构验证通过")
        
        if not valid_sheets:
            raise ValueError("所有sheet都缺少必需的列")
        all_sheets = valid_sheets
    
    # 合并所有sheet的数据
    combined_data = []
    total_rows = 0
    
    print("开始合并Sheet数据...")
    for sheet_name, df in all_sheets.items():
        df['数据来源Sheet'] = sheet_name
        combined_data.append(df)
        total_rows += len(df)
        print(f"  sheet '{sheet_name}': {len(df)} 行数据")
    
    # 垂直合并所有数据
    if combined_data:
        try:
            final_df = pd.concat(combined_data, ignore_index=True)
            print(f"成功合并数据! 总行数: {total_rows:,}")
            print(f"合并后的列名: {', '.join(final_df.columns)}")
        except Exception as e:
            print(f"数据合并失败: {e}")
            # 尝试逐个合并
            print("尝试逐个合并数据...")
            final_df = combined_data[0]
            for i in range(1, len(combined_data)):
                try:
                    final_df = pd.concat([final_df, combined_data[i]], ignore_index=True)
                except Exception as e2:
                    print(f"合并第{i+1}个sheet失败: {e2}")
    else:
        raise ValueError("没有有效数据可合并")
    
    return final_df

def split_data_by_production_date(data, import_date_str):
    """
    根据生产日期将数据分为改善前和改善后两组
    """
    if data is None:
        raise ValueError("输入数据为None，无法进行分组")
    
    print(f"\n开始根据导入时间 {import_date_str} 分组数据...")
    
    # 确保数据包含必要的列
    required_cols = ['生产日期', '出货时间', '退货时间', '是否质量问题']
    missing_cols = [col for col in required_cols if col not in data.columns]
    if missing_cols:
        print(f"警告: 数据中缺少必要的列: {missing_cols}")
        print(f"可用列: {data.columns.tolist()}")
        raise ValueError(f"数据中缺少必要的列: {missing_cols}")
    
    # 转换日期列为datetime类型
    data['生产日期'] = pd.to_datetime(data['生产日期'], errors='coerce')
    data['出货时间'] = pd.to_datetime(data['出货时间'], errors='coerce')
    data['退货时间'] = pd.to_datetime(data['退货时间'], errors='coerce')
    import_date = pd.to_datetime(import_date_str)
    
    print(f"原始数据行数: {len(data)}")
    
    # 根据生产日期分组
    before_improvement_mask = data['生产日期'] < import_date
    after_improvement_mask = data['生产日期'] >= import_date
    
    before_data = data[before_improvement_mask].copy()
    after_data = data[after_improvement_mask].copy()
    
    print(f"改善前数据（生产日期 < {import_date_str}）: {len(before_data)} 行")
    print(f"改善后数据（生产日期 >= {import_date_str}）: {len(after_data)} 行")
    
    if len(before_data) == 0 and len(after_data) == 0:
        raise ValueError("分组后没有剩余数据，请检查日期格式和数据")
    
    return before_data, after_data

def prepare_reliability_data(data, import_date_str, analysis_months=36):
    """
    准备可靠性分析数据 - 根据导入时间处理数据
    """
    if data is None:
        raise ValueError("输入数据为None，无法进行处理")
    
    print(f"\n开始准备可靠性分析数据，导入时间: {import_date_str}")
    
    # 确保数据包含必要的列
    required_cols = ['生产日期', '出货时间', '退货时间', '是否质量问题']
    missing_cols = [col for col in required_cols if col not in data.columns]
    if missing_cols:
        print(f"警告: 数据中缺少必要的列: {missing_cols}")
        print(f"可用列: {data.columns.tolist()}")
        raise ValueError(f"数据中缺少必要的列: {missing_cols}")
    
    # 转换日期列为datetime类型
    data['生产日期'] = pd.to_datetime(data['生产日期'], errors='coerce')
    data['出货时间'] = pd.to_datetime(data['出货时间'], errors='coerce')
    data['退货时间'] = pd.to_datetime(data['退货时间'], errors='coerce')
    import_date = pd.to_datetime(import_date_str)
    
    print(f"原始数据行数: {len(data)}")
    
    # 筛选逻辑：只保留生产日期在分析范围内的数据
    # 假设分析范围为导入时间前后指定月数
    start_date = import_date - pd.DateOffset(months=analysis_months)
    end_date = import_date + pd.DateOffset(months=analysis_months)
    
    time_range_mask = (data['生产日期'] >= start_date) & (data['生产日期'] <= end_date)
    filtered_data = data[time_range_mask].copy()
    print(f"时间范围筛选后: {len(filtered_data)} 行 (生产日期在 {start_date.date()} 到 {end_date.date()} 之间)")
    
    if len(filtered_data) == 0:
        raise ValueError("筛选后没有剩余数据，请检查日期范围")
    
    # 重新标记事件类型
    print("检查'是否质量问题'列的数据类型和内容...")
    print(f"'是否质量问题'列的数据类型: {filtered_data['是否质量问题'].dtype}")
    print(f"'是否质量问题'列的唯一值: {filtered_data['是否质量问题'].unique()}")
    
    # 统一处理"是否质量问题"列为字符串类型，并标准化内容
    filtered_data['是否质量问题'] = filtered_data['是否质量问题'].astype(str)
    filtered_data['是否质量问题'] = filtered_data['是否质量问题'].str.strip().str.replace(' ', '')
    
    # 创建新的事件标记列，基于退货时间和质量问题状态
    # 条件1: 退货时间在截止日期之前
    # 条件2: 是否质量问题为"是"
    failure_condition = (
        (filtered_data['退货时间'].notna()) & 
        (filtered_data['退货时间'] <= end_date) &
        (filtered_data['是否质量问题'].str.upper().str.contains('是|YES|TRUE|1', na=False))
    )
    
    filtered_data['event'] = np.where(failure_condition, 1, 0)
    
    # 计算duration（对于截尾数据，duration应为出货时间到截止日期的时间差）
    mask_censored = filtered_data['event'] == 0
    filtered_data.loc[mask_censored, 'duration'] = (
        end_date - filtered_data.loc[mask_censored, '出货时间']
    ).dt.days
    
    # 对于失效数据，确保duration是正数（从出货时间到退货时间）
    mask_failure = filtered_data['event'] == 1
    filtered_data.loc[mask_failure, 'duration'] = (
        filtered_data.loc[mask_failure, '退货时间'] - filtered_data.loc[mask_failure, '出货时间']
    ).dt.days
    
    # 确保duration为正数且合理
    filtered_data = filtered_data[
        (filtered_data['duration'] > 0) & 
        (filtered_data['duration'] < 10000)  # 合理的最大天数
    ]
    
    print(f"最终有效数据行数: {len(filtered_data)}")
    print(f"失效事件数量: {filtered_data['event'].sum()}")
    print(f"截尾数据数量: {len(filtered_data) - filtered_data['event'].sum()}")
    
    # 调试信息：显示一些样本数据
    print("\n样本数据检查:")
    sample_failures = filtered_data[filtered_data['event'] == 1].head(3)
    if len(sample_failures) > 0:
        print("失效数据样本:")
        for _, row in sample_failures.iterrows():
            print(f"  生产日期: {row['生产日期']}, 出货时间: {row['出货时间']}, 退货时间: {row['退货时间']}, "
                  f"duration: {row['duration']}天, 是否质量问题: {row['是否质量问题']}")
    
    sample_censored = filtered_data[filtered_data['event'] == 0].head(2)
    if len(sample_censored) > 0:
        print("截尾数据样本:")
        for _, row in sample_censored.iterrows():
            print(f"  生产日期: {row['生产日期']}, 出货时间: {row['出货时间']}, 退货时间: {row['退货时间']}, "
                  f"duration: {row['duration']}天, 是否质量问题: {row['是否质量问题']}")
    
    return filtered_data

class AdvancedReliabilityAnalyzer:
    """高级可靠性分析器"""
    
    def __init__(self, analysis_months=36, interval_months=6):
        self.analysis_months = analysis_months
        self.interval_months = interval_months
        self.key_time_points = generate_time_points(analysis_months, interval_months)
        self.analysis_results = {}
        self.comparison_results = {}  # 存储改善前后对比结果
        
        # 定义15个分布的列表，与用户提到的图1一致
        self.distribution_list = [
            ('Weibull_DS', Fit_Weibull_2P, 'Weibull_DS'),
            ('Weibull_Mixture', Fit_Weibull_Mixture, 'Weibull_Mixture'),
            ('Lognormal_3P', Fit_Lognormal_2P, 'Lognormal_3P'),  # 注意：可能需要不同的拟合函数
            ('Lognormal_2P', Fit_Lognormal_2P, 'Lognormal_2P'),
            ('Loglogistic_3P', Fit_Loglogistic_2P, 'Loglogistic_3P'),  # 注意：可能需要不同的拟合函数
            ('Weibull_3P', Fit_Weibull_3P, 'Weibull_3P'),
            ('Gamma_3P', Fit_Gamma_2P, 'Gamma_3P'),  # 注意：可能需要不同的拟合函数
            ('Loglogistic_2P', Fit_Loglogistic_2P, 'Loglogistic_2P'),
            ('Weibull_2P', Fit_Weibull_2P, 'Weibull_2P'),
            ('Weibull_CR', Fit_Weibull_CR, 'Weibull_CR'),
            ('Gamma_2P', Fit_Gamma_2P, 'Gamma_2P'),
            ('Exponential_2P', Fit_Exponential_2P, 'Exponential_2P'),
            ('Exponential_1P', Fit_Exponential_2P, 'Exponential_1P'),  # 注意：可能需要不同的拟合函数
            ('Normal_2P', Fit_Normal_2P, 'Normal_2P'),
            ('Gumbel_2P', Fit_Gumbel_2P, 'Gumbel_2P'),
        ]
        
        self.all_results = []  # 存储所有分布拟合结果
    
    def process_failure_data(self, data):
        """处理失效数据"""
        if data is None:
            raise ValueError("输入数据为None")
            
        if 'duration' not in data.columns or 'event' not in data.columns:
            raise ValueError("数据中必须包含'duration'和'event'列")
        
        durations = data['duration']
        valid_mask = (durations > 0) & (durations.notna()) & (durations < 10000)
        durations = durations[valid_mask]
        
        # 使用新的事件列
        events = data.loc[valid_mask, 'event']
        
        print(f"数据清洗: 原始{len(data)}行 -> 有效{len(durations)}行")
        print(f"失效事件数量: {events.sum()}")
        print(f"截尾数据数量: {len(events) - events.sum()}")
        
        return durations, events

    def _calculate_ad_statistic_corrected(self, durations, events, distribution):
        """计算修正的AD统计量"""
        try:
            if distribution is None:
                return 1000, 1000  # 返回原始AD和修正AD
                
            failure_times = durations[events == 1].values
            failure_times = failure_times[failure_times > 0]
            
            if len(failure_times) < 2:
                return 1000, 1000
                
            sorted_failures = np.sort(failure_times)
            n = len(sorted_failures)
            
            # 计算原始AD统计量
            empirical_quantiles = np.arange(1, n + 1) / (n + 1)
            
            try:
                theoretical_quantiles = [distribution.CDF(float(t)) for t in sorted_failures]
                theoretical_quantiles = np.clip(theoretical_quantiles, 1e-10, 1 - 1e-10)
                
                # 计算AD统计量
                ad_sum = 0
                for i in range(n):
                    term1 = np.log(theoretical_quantiles[i])
                    term2 = np.log(1 - theoretical_quantiles[n - i - 1])
                    ad_sum += (2 * (i + 1) - 1) * (term1 + term2)
                
                ad_value = -n - (1/n) * ad_sum
                
                # 应用Stephens修正公式: A2* = A2 × (1 + 0.75/n + 2.25/n²) 
                if n > 0:
                    correction_factor = 1 + 0.75/n + 2.25/(n**2)
                    ad_corrected = ad_value * correction_factor
                else:
                    ad_corrected = ad_value
                
                return ad_value, ad_corrected
                
            except Exception as e:
                print(f"计算AD统计量时出错: {e}")
                return 1000, 1000
                
        except Exception as e:
            print(f"AD统计量计算失败: {e}")
            return 1000, 1000

    def analyze_improvement_comparison(self, before_data, after_data, group_name="改善前后对比"):
        """分析改善前后对比"""
        print("\n" + "="*60)
        print(f"{group_name}分析")
        print("="*60)
        
        results = {}
        
        # 分析改善前数据
        if len(before_data) > 0:
            print("\n--- 改善前数据分析 ---")
            before_result = self.analyze_complete_data_with_selection(before_data, "改善前")
            if before_result:
                results['改善前'] = before_result
                print(f"改善前最佳模型: {before_result['best_model']}")
                print(f"改善前AD修正值: {before_result['AD_corrected']:.4f}")
                print(f"改善前BIC: {before_result['BIC']:.2f}")
        
        # 分析改善后数据
        if len(after_data) > 0:
            print("\n--- 改善后数据分析 ---")
            after_result = self.analyze_complete_data_with_selection(after_data, "改善后")
            if after_result:
                results['改善后'] = after_result
                print(f"改善后最佳模型: {after_result['best_model']}")
                print(f"改善后AD修正值: {after_result['AD_corrected']:.4f}")
                print(f"改善后BIC: {after_result['BIC']:.2f}")
        
        if not results:
            print("错误: 两组数据均分析失败")
            return None
        
        self.comparison_results[group_name] = results
        return results

    def analyze_complete_data_with_selection(self, data, group_name="数据组"):
        """完整数据分析，支持用户手动选择最佳模型"""
        print("\n" + "="*60)
        print(f"{group_name}完整数据分析")
        print("="*60)
        
        if data is None:
            print("错误: 输入数据为None")
            return None
            
        durations, events = self.process_failure_data(data)
        
        if len(durations) == 0:
            print("错误: 没有有效数据")
            return None
        
        if events.sum() < 2:
            print(f"错误: 失效数量不足: {events.sum()}")
            return None
        
        return self._perform_analysis_with_selection(durations, events, group_name)

    def _perform_analysis_with_selection(self, durations, events, group_name):
        """执行分析，支持用户选择最佳模型"""
        print("执行数据分析...")
        
        # 数据验证
        valid_mask = (durations > 0) & (durations.notna())
        durations = durations[valid_mask]
        events = events[valid_mask]
        
        n_failures = events.sum()
        if n_failures < 2:
            print(f"失效数量不足: {n_failures}")
            return None
        
        print(f"分析数据: {n_failures}个失效, {len(durations)}个总样本")
        
        # 准备失效时间和截尾时间
        failure_times = durations[events == 1].values
        failure_times = failure_times[failure_times > 0]
        right_censored = durations[events == 0].values
        right_censored = right_censored[right_censored > 0]
        
        # 拟合所有分布并生成概率图
        all_results = self._fit_all_distributions_with_probability_plots(failure_times, right_censored, durations, events, group_name)
        
        if not all_results:
            print("错误: 所有分布拟合失败")
            return None
        
        # 显示所有模型比较结果
        self._display_all_models_comparison(all_results)
        
        # 询问用户选择
        selected_result = self._ask_user_for_selection(all_results, group_name)
        
        if selected_result is None:
            print("警告: 用户没有选择，将使用默认最佳模型")
            # 使用默认选择逻辑
            selected_result = self._select_best_model_by_new_criteria(all_results)
        
        if selected_result is None:
            print("错误: 无法选择最佳模型")
            return None
        
        dist_name, ad_value, ad_corrected, bic_value, distribution_obj, fit_result = selected_result
        
        # 计算累计失效率
        cumulative_failure = self._calculate_cumulative_failure_robust(distribution_obj, dist_name)
        
        # 计算失效率函数
        failure_rate = self._calculate_failure_rate_robust(distribution_obj, dist_name)
        
        # 存储结果
        result = {
            'analysis_type': '完整数据分析',
            'event_count': n_failures,
            'total_count': len(durations),
            'best_model': dist_name,
            'AD': ad_value,
            'AD_corrected': ad_corrected,
            'BIC': bic_value,
            'cumulative_failure': cumulative_failure,
            'failure_rate': failure_rate,
            'distribution': distribution_obj,
            'fit_result': fit_result,
            'durations': durations,
            'events': events
        }
        
        self.analysis_results['完整数据分析'] = result
        return result

    def _fit_all_distributions_with_probability_plots(self, failure_times, right_censored, durations, events, group_name):
        """拟合所有分布并生成概率图"""
        print("拟合所有分布模型并生成概率图...")
        print(f"数据: {len(failure_times)}个失效时间, {len(right_censored)}个截尾时间")
        
        all_results = []
        
        # 创建输出目录
        output_dir = "改善对比分析结果"
        os.makedirs(output_dir, exist_ok=True)
        
        # 使用Fit_Everything一次性拟合所有分布
        try:
            print("  使用Fit_Everything一次性拟合所有分布...")
            fit_everything = Fit_Everything(
                failures=failure_times,
                right_censored=right_censored if len(right_censored) > 0 else None,
                show_probability_plot=True,  # 设置为True以生成概率图
                print_results=True  # 设置为True以打印结果
            )
            
            # 只保留图3的解释，其他图片将被删除
            print("\n图3 (Probability plot of best distribution) 详细解释:")
            print("- 显示最佳拟合分布的概率图，使用Weibull概率纸")
            print("- 黑色点表示实际数据点（失效时间）")
            print("- 蓝色线表示理论拟合线")
            print("- 数据点越接近拟合线，说明拟合效果越好")
            print("- 图中还显示了95%置信区间（虚线）")
            print("- 最佳分布是根据AD值自动选择的")
            
            # 输出最佳分布选择逻辑
            print("\n最佳分布选择逻辑:")
            print("1. 计算每个分布的Anderson-Darling (AD) 统计量")
            print("2. 对所有分布按AD值从小到大排序")
            print("3. 选择AD值最小的分布作为最佳模型")
            print("4. 如果多个分布的AD值差异在5%以内，选择BIC值最小的模型")
            
            # 关闭所有Fit_Everything生成的图表，只保留我们自己生成的概率图
            import matplotlib.pyplot as plt
            plt.close('all')
            
            # 提取所有拟合结果
            fit_results = []
            # 尝试不同的属性名来获取拟合结果
            distributions = None
            try:
                # 打印fit_everything的所有属性，以便调试
                print("  调试: fit_everything属性:", [attr for attr in dir(fit_everything) if not attr.startswith('_')])
                
                # 尝试不同的属性名
                possible_attributes = ['distributions', 'fits', 'results', 'model_results', 'model_fits', 'fit_results']
                for attr in possible_attributes:
                    if hasattr(fit_everything, attr):
                        distributions = getattr(fit_everything, attr)
                        print(f"  使用fit_everything.{attr}获取拟合结果")
                        print(f"  类型: {type(distributions).__name__}")
                        break
                
                if distributions is None:
                    # 尝试直接访问属性
                    try:
                        distributions = fit_everything.distributions
                        print("  直接访问fit_everything.distributions获取拟合结果")
                    except:
                        pass
                
                if distributions is None:
                    # 尝试使用fit_everything的其他属性
                    print("  尝试使用其他方法获取拟合结果...")
                    # 检查是否有__dict__属性
                    if hasattr(fit_everything, '__dict__'):
                        print("  fit_everything.__dict__ keys:", [k for k in fit_everything.__dict__ if not k.startswith('_')])
                        # 尝试从__dict__中获取分布信息
                        for key in fit_everything.__dict__:
                            if not key.startswith('_'):
                                value = fit_everything.__dict__[key]
                                if hasattr(value, '__iter__') and not isinstance(value, (str, bytes)):
                                    print(f"  检查属性 {key}: {type(value).__name__}")
                
                if distributions is None:
                    print("  无法获取Fit_Everything的分布结果，使用回退方法")
            except Exception as attr_error:
                print(f"  获取拟合结果时出错: {attr_error}")
                import traceback
                traceback.print_exc()
            
            # 如果成功获取到分布结果，处理它们
            if distributions is not None:
                try:
                    # 检查distributions的类型
                    print(f"  distributions类型: {type(distributions).__name__}")
                    
                    # 处理不同类型的distributions
                    if isinstance(distributions, dict):
                        print(f"  处理字典类型的distributions，包含 {len(distributions)} 个分布")
                        for i, (dist_name, fit_result) in enumerate(distributions.items(), 1):
                            if fit_result is not None:
                                print(f"  {i:2d}. 拟合 {dist_name} 成功")
                                fit_results.append((dist_name, dist_name, fit_result, failure_times, right_censored))
                            else:
                                print(f"  {i:2d}. 拟合 {dist_name} 失败")
                    elif hasattr(distributions, '__iter__'):
                        print(f"  处理可迭代类型的distributions")
                        for i, item in enumerate(distributions, 1):
                            if isinstance(item, (list, tuple)) and len(item) >= 2:
                                dist_name, fit_result = item[0], item[1]
                                if fit_result is not None:
                                    print(f"  {i:2d}. 拟合 {dist_name} 成功")
                                    fit_results.append((dist_name, dist_name, fit_result, failure_times, right_censored))
                                else:
                                    print(f"  {i:2d}. 拟合 {dist_name} 失败")
                    else:
                        print(f"  无法处理distributions类型: {type(distributions).__name__}")
                    
                    # 计算每个分布的AD统计量和BIC
                    print(f"  处理 {len(fit_results)} 个拟合结果")
                    for i, (dist_name, display_name, fit_result, _, _) in enumerate(fit_results, 1):
                        try:
                            distribution_obj = self._create_distribution_object(fit_result, dist_name)
                            if distribution_obj is None:
                                print(f"  跳过 {display_name}，因为无法创建分布对象")
                                continue
                            
                            # 计算AD统计量和BIC
                            ad_value, ad_corrected = self._calculate_ad_statistic_corrected(durations, events, distribution_obj)
                            bic_value = self._calculate_bic_value(fit_result, len(failure_times))
                            
                            all_results.append({
                                'index': i,
                                'dist_name': dist_name,
                                'dist_display_name': display_name,
                                'ad_value': ad_value,
                                'ad_corrected': ad_corrected,
                                'bic_value': bic_value,
                                'distribution': distribution_obj,
                                'fit_result': fit_result
                            })
                            
                            print(f"    {display_name}: AD={ad_value:.4f}, AD修正={ad_corrected:.4f}, BIC={bic_value:.2f}")
                        except Exception as e:
                            print(f"  计算 {display_name} 统计量失败: {e}")
                            import traceback
                            traceback.print_exc()
                            continue
                except Exception as process_error:
                    print(f"  处理拟合结果时出错: {process_error}")
                    import traceback
                    traceback.print_exc()
                    # 继续使用回退方法
            
            # 如果没有获取到任何结果，使用回退方法
            if not all_results:
                print("  没有获取到拟合结果，使用回退方法")
                raise Exception("无法获取Fit_Everything的分布结果")
            
            # 根据AD值选择最佳模型，并重新生成图3
            if all_results:
                # 按AD修正值排序
                sorted_by_ad = sorted(all_results, key=lambda x: x['ad_corrected'])
                best_model = sorted_by_ad[0]
                print(f"\n根据AD值选择的最佳模型: {best_model['dist_display_name']} (AD修正值: {best_model['ad_corrected']:.4f})")
                
                # 为最佳模型生成新的概率图
                try:
                    import matplotlib.pyplot as plt
                    from reliability.Probability_plotting import plot_points
                    
                    # 关闭现有的图3
                    plt.close(3)
                    # 创建新的图3
                    fig3 = plt.figure(3)
                    
                    # 使用reliability.Probability_plotting库的方法生成概率图
                    print("  使用reliability.Probability_plotting.plot_points生成概率图")
                    sorted_failures = np.sort(failure_times)
                    n_failures = len(sorted_failures)
                    if n_failures > 0:
                        # 绘制概率图
                        plot_points(
                            failures=sorted_failures,
                            right_censored=right_censored if right_censored is not None and len(right_censored) > 0 else None
                        )
                        
                        # 如果有分布对象，绘制拟合线
                        if best_model['distribution'] is not None:
                            print(f"  为{best_model['dist_display_name']}绘制拟合线")
                            # 生成x轴范围
                            x_min = min(sorted_failures) * 0.9 if min(sorted_failures) > 0 else 0.1
                            x_max = max(sorted_failures) * 1.1
                            x_range = np.linspace(x_min, x_max, 100)
                            
                            # 计算CDF值
                            y_range = [best_model['distribution'].CDF(x) for x in x_range]
                            
                            # 绘制拟合线
                            plt.plot(x_range, y_range, 'b-', label='拟合线')
                            plt.legend()
                    
                    # 设置标题
                    plt.title(f"Probability plot of best distribution\n{best_model['dist_display_name']} (AD-based)")
                    plt.xlabel('时间')
                    plt.ylabel('累积失效概率')
                    
                    # 保存新的图3
                    plot_path3 = os.path.join(output_dir, f'{group_name}_Figure_3.png')
                    fig3.savefig(plot_path3, dpi=300, bbox_inches='tight')
                    print(f"\n已更新图3为AD值最佳模型: {plot_path3}")
                except Exception as e:
                    print(f"  生成最佳模型概率图失败: {e}")
                    import traceback
                    traceback.print_exc()
        except Exception as e:
            print(f"  使用Fit_Everything失败: {e}")
            # 回退到原来的方法
            print("  回退到逐个拟合分布的方法...")
            fit_results = []
            
            for i, (dist_name, fitter, dist_display_name) in enumerate(self.distribution_list, 1):
                try:
                    print(f"  {i:2d}. 拟合 {dist_display_name}...")
                    
                    fit_result = None
                    
                    # 特殊处理高级分布模型
                    if dist_name in ['Weibull_Mixture', 'Weibull_CR']:
                        if len(failure_times) < 20:
                            print(f"    {dist_display_name}: 数据量不足(需要至少20个失效), 跳过")
                            continue
                        
                        try:
                            if dist_name == 'Weibull_Mixture':
                                fit_result = Fit_Weibull_Mixture(
                                    failures=failure_times,
                                    right_censored=right_censored if len(right_censored) > 0 else None,
                                    show_probability_plot=False,
                                    print_results=False,
                                    optimizer='L-BFGS-B'
                                )
                            else:
                                fit_result = Fit_Weibull_CR(
                                    failures=failure_times,
                                    right_censored=right_censored if len(right_censored) > 0 else None,
                                    show_probability_plot=False,
                                    print_results=False
                                )
                        except Exception as e1:
                            print(f"    {dist_display_name} 第一次拟合失败: {e1}")
                            try:
                                if dist_name == 'Weibull_Mixture':
                                    fit_result = Fit_Weibull_Mixture(
                                        failures=failure_times,
                                        right_censored=right_censored if len(right_censored) > 0 else None,
                                        show_probability_plot=False,
                                        print_results=False
                                    )
                                else:
                                    fit_result = Fit_Weibull_CR(
                                        failures=failure_times,
                                        right_censored=right_censored if len(right_censored) > 0 else None,
                                        show_probability_plot=False,
                                        print_results=False
                                    )
                            except Exception as e2:
                                print(f"    {dist_display_name} 最终拟合失败: {e2}")
                                continue
                    else:
                        # 基础分布使用标准拟合方法
                        try:
                            fit_result = fitter(
                                failures=failure_times,
                                right_censored=right_censored if len(right_censored) > 0 else None,
                                show_probability_plot=False,
                                print_results=False
                            )
                        except Exception as e:
                            print(f"    {dist_display_name} 拟合失败: {e}")
                            continue
                    
                    if fit_result is None:
                        continue
                    
                    # 存储拟合结果，包括失效时间和截尾时间
                    fit_results.append((dist_name, dist_display_name, fit_result, failure_times, right_censored))
                    
                    distribution_obj = self._create_distribution_object(fit_result, dist_name)
                    if distribution_obj is None:
                        continue
                    
                    # 计算AD统计量和BIC
                    ad_value, ad_corrected = self._calculate_ad_statistic_corrected(durations, events, distribution_obj)
                    bic_value = self._calculate_bic_value(fit_result, len(failure_times))
                    
                    all_results.append({
                        'index': i,
                        'dist_name': dist_name,
                        'dist_display_name': dist_display_name,
                        'ad_value': ad_value,
                        'ad_corrected': ad_corrected,
                        'bic_value': bic_value,
                        'distribution': distribution_obj,
                        'fit_result': fit_result
                    })
                    
                    print(f"    {dist_display_name}: AD={ad_value:.4f}, AD修正={ad_corrected:.4f}, BIC={bic_value:.2f}")
                    
                except Exception as e:
                    print(f"  拟合 {dist_display_name} 失败: {e}")
                    continue
        
        return all_results

    def _plot_probability_plots_grid(self, fit_results, group_name, output_dir):
        """生成概率图网格（已取消）"""
        print(f"\n概率图网格绘制已取消，因为用户表示这部分代码总是出错")
        print("  已使用Fit_Everything生成了更详细的拟合效果图表")
        print("  请查看已保存的Figure_1.png、Figure_2.png和Figure_3.png文件")
        return

    def _display_all_models_comparison(self, all_results):
        """显示所有模型比较结果"""
        print("\n" + "="*100)
        print("所有分布模型比较结果")
        print("="*100)
        
        # 按AD修正值排序
        sorted_results = sorted(all_results, key=lambda x: x['ad_corrected'])
        
        print(f"{'编号':<5} {'分布名称':<25} {'AD值':<10} {'AD修正值':<12} {'BIC':<10} {'选择状态':<10}")
        print("-" * 80)
        
        for i, result in enumerate(sorted_results, 1):
            marker = "★" if i == 1 else " "
            print(f"{marker} {result['index']:<3} {result['dist_display_name']:<23} "
                  f"{result['ad_value']:>8.4f} {result['ad_corrected']:>10.4f} "
                  f"{result['bic_value']:>10.2f} {'最佳' if i == 1 else ''}")

    def _ask_user_for_selection(self, all_results, group_name):
        """询问用户选择最佳模型"""
        print("\n" + "="*100)
        print(f"请为'{group_name}'选择最佳拟合分布:")
        print("="*100)
        
        # 显示可选的分布列表
        print("可选分布列表:")
        print("-" * 80)
        print(f"{'编号':<5} {'分布名称':<25} {'AD修正值':<12} {'BIC':<10}")
        print("-" * 80)
        
        for result in sorted(all_results, key=lambda x: x['ad_corrected']):
            print(f"{result['index']:<5} {result['dist_display_name']:<25} "
                  f"{result['ad_corrected']:>10.4f} {result['bic_value']:>10.2f}")
        
        print("\n选择说明:")
        print("1. 输入分布编号（1-15）选择对应的分布作为最佳模型")
        print("2. 输入0或不输入，将自动选择AD修正值最小的分布")
        print("3. 自动选择逻辑: AD修正值最小，差异5%以内时选择BIC最小的模型")
        print("4. 建议查看上面的概率图网格，选择拟合线最接近数据点的分布")
        print("5. AD修正值越小表示拟合越好，BIC值越小表示模型越简洁")
        
        while True:
            try:
                user_input = input(f"\n请输入为'{group_name}'选择的分布编号 (0-{len(all_results)}，默认0): ").strip()
                
                if not user_input:  # 直接回车
                    print("使用自动选择（AD修正值最小的分布）")
                    return self._select_best_model_by_new_criteria(all_results)
                
                choice = int(user_input)
                
                if choice == 0:
                    print("使用自动选择（AD修正值最小的分布）")
                    return self._select_best_model_by_new_criteria(all_results)
                
                # 查找用户选择的分布
                for result in all_results:
                    if result['index'] == choice:
                        print(f"您选择了: {result['dist_display_name']}")
                        return (result['dist_name'], result['ad_value'], result['ad_corrected'], 
                                result['bic_value'], result['distribution'], result['fit_result'])
                
                print(f"错误: 编号 {choice} 不在可选范围内，请重新输入")
                
            except ValueError:
                print("错误: 请输入有效的数字")
            except Exception as e:
                print(f"选择错误: {e}")

    def _select_best_model_by_new_criteria(self, all_results):
        """根据新规则选择最佳模型：先判断AD修正值，差异5%以内再判断BIC"""
        if not all_results:
            return None
        
        # 将结果转换为元组列表
        result_tuples = []
        for result in all_results:
            result_tuples.append((
                result['dist_name'], 
                result['ad_value'], 
                result['ad_corrected'], 
                result['bic_value'], 
                result['distribution'], 
                result['fit_result']
            ))
        
        # 按AD修正值排序
        sorted_by_ad = sorted(result_tuples, key=lambda x: x[2])  # x[2]是AD修正值
        
        # 最佳AD修正值
        best_ad_corrected = sorted_by_ad[0][2]
        
        # 找出AD修正值差异在5%以内的模型
        candidates = []
        for result in sorted_by_ad:
            ad_corrected = result[2]
            # 计算与最佳AD的差异百分比
            if best_ad_corrected > 0:
                diff_percent = (ad_corrected - best_ad_corrected) / best_ad_corrected * 100
            else:
                diff_percent = 0
            
            if diff_percent <= 5:  # 差异在5%以内
                candidates.append(result)
            else:
                break  # 由于已排序，后续模型的AD值会更大
        
        print(f"\nAD修正值差异在5%以内的候选模型 ({len(candidates)}个):")
        for result in candidates:
            print(f"  {result[0]}: AD修正={result[2]:.4f}, BIC={result[3]:.2f}")
        
        # 如果只有一个候选模型，直接选择
        if len(candidates) == 1:
            print(f"只有一个候选模型，选择: {candidates[0][0]}")
            return candidates[0]
        
        # 如果有多个候选模型，选择BIC最小的
        best_by_bic = min(candidates, key=lambda x: x[3])  # x[3]是BIC值
        print(f"多个候选模型，选择BIC最小的: {best_by_bic[0]} (BIC={best_by_bic[3]:.2f})")
        
        return best_by_bic

    def _get_param(self, fit_result, name, default=None):
        """获取参数的辅助函数"""
        if hasattr(fit_result, name):
            value = getattr(fit_result, name)
            print(f"      获取参数 {name}: {value}")
            return value
        elif isinstance(fit_result, (dict, pd.Series)) and name in fit_result:
            value = fit_result[name]
            print(f"      从Series/dict获取参数 {name}: {value}")
            return value
        else:
            print(f"      参数 {name} 不存在")
            return default
    
    def _create_weibull_distribution(self, fit_result, dist_name):
        """创建Weibull分布对象"""
        alpha = self._get_param(fit_result, 'alpha')
        beta = self._get_param(fit_result, 'beta')
        gamma = self._get_param(fit_result, 'gamma', 0)
        if alpha is not None and beta is not None:
            print(f"        Weibull参数: alpha={alpha:.4f}, beta={beta:.4f}, gamma={gamma:.4f}")
            return Weibull_Distribution(alpha=alpha, beta=beta, gamma=gamma)
        return None
    
    def _create_lognormal_distribution(self, fit_result):
        """创建Lognormal分布对象"""
        mu = self._get_param(fit_result, 'mu')
        sigma = self._get_param(fit_result, 'sigma')
        if mu is not None and sigma is not None:
            print(f"        Lognormal参数: mu={mu:.4f}, sigma={sigma:.4f}")
            return Lognormal_Distribution(mu=mu, sigma=sigma)
        return None
    
    def _create_normal_distribution(self, fit_result):
        """创建Normal分布对象"""
        mu = self._get_param(fit_result, 'mu')
        sigma = self._get_param(fit_result, 'sigma')
        if mu is not None and sigma is not None:
            print(f"        Normal参数: mu={mu:.4f}, sigma={sigma:.4f}")
            return Normal_Distribution(mu=mu, sigma=sigma)
        return None
    
    def _create_gamma_distribution(self, fit_result):
        """创建Gamma分布对象"""
        alpha = self._get_param(fit_result, 'alpha')
        beta = self._get_param(fit_result, 'beta')
        if alpha is not None and beta is not None:
            print(f"        Gamma参数: alpha={alpha:.4f}, beta={beta:.4f}")
            return Gamma_Distribution(alpha=alpha, beta=beta)
        return None
    
    def _create_loglogistic_distribution(self, fit_result):
        """创建Loglogistic分布对象"""
        alpha = self._get_param(fit_result, 'alpha')
        beta = self._get_param(fit_result, 'beta')
        if alpha is not None and beta is not None:
            print(f"        Loglogistic参数: alpha={alpha:.4f}, beta={beta:.4f}")
            return Loglogistic_Distribution(alpha=alpha, beta=beta)
        return None
    
    def _create_gumbel_distribution(self, fit_result):
        """创建Gumbel分布对象"""
        mu = self._get_param(fit_result, 'mu')
        sigma = self._get_param(fit_result, 'sigma')
        if mu is not None and sigma is not None:
            print(f"        Gumbel参数: mu={mu:.4f}, sigma={sigma:.4f}")
            return Gumbel_Distribution(mu=mu, sigma=sigma)
        return None
    
    def _create_exponential_distribution(self, fit_result):
        """创建Exponential分布对象"""
        Lambda = self._get_param(fit_result, 'Lambda')
        if Lambda is not None:
            print(f"        Exponential参数: Lambda={Lambda:.6f}")
            return Exponential_Distribution(Lambda=Lambda)
        return None
    
    def _create_mixture_model(self, fit_result):
        """创建混合模型"""
        alpha1 = self._get_param(fit_result, 'alpha_1')
        beta1 = self._get_param(fit_result, 'beta_1')
        alpha2 = self._get_param(fit_result, 'alpha_2')
        beta2 = self._get_param(fit_result, 'beta_2')
        proportion1 = self._get_param(fit_result, 'proportion_1', 0.5)
        if alpha1 is not None and beta1 is not None and alpha2 is not None and beta2 is not None:
            print(f"        Weibull_Mixture参数:")
            print(f"          组分1: alpha={alpha1:.4f}, beta={beta1:.4f}, 比例={proportion1:.4f}")
            print(f"          组分2: alpha={alpha2:.4f}, beta={beta2:.4f}, 比例={1-proportion1:.4f}")
            # 直接创建混合模型
            dist1 = Weibull_Distribution(alpha=alpha1, beta=beta1)
            dist2 = Weibull_Distribution(alpha=alpha2, beta=beta2)
            return Mixture_Model(distributions=[dist1, dist2], proportions=[proportion1, 1-proportion1])
        return None
    
    def _create_competing_risks_model(self, fit_result):
        """创建竞争风险模型"""
        alpha1 = self._get_param(fit_result, 'alpha_1')
        beta1 = self._get_param(fit_result, 'beta_1')
        alpha2 = self._get_param(fit_result, 'alpha_2')
        beta2 = self._get_param(fit_result, 'beta_2')
        if alpha1 is not None and beta1 is not None and alpha2 is not None and beta2 is not None:
            print(f"        Weibull_CR参数:")
            print(f"          风险1: alpha={alpha1:.4f}, beta={beta1:.4f}")
            print(f"          风险2: alpha={alpha2:.4f}, beta={beta2:.4f}")
            # 直接创建竞争风险模型
            dist1 = Weibull_Distribution(alpha=alpha1, beta=beta1)
            dist2 = Weibull_Distribution(alpha=alpha2, beta=beta2)
            return Competing_Risks_Model(distributions=[dist1, dist2])
        return None
    
    def _create_distribution_object(self, fit_result, dist_name):
        """创建分布对象"""
        try:
            print(f"    创建{dist_name}分布对象...")
            
            # 检查fit_result类型
            print(f"    fit_result类型: {type(fit_result).__name__}")
            
            # 首先尝试使用fit_result自带的distribution属性
            if hasattr(fit_result, 'distribution') and fit_result.distribution is not None:
                print(f"      使用fit_result.distribution属性")
                return fit_result.distribution
            
            # 如果fit_result没有distribution属性，则手动创建
            print(f"      手动创建{dist_name}分布对象")
            
            # 根据分布类型选择创建方法
            if 'Weibull' in dist_name or dist_name in ['Weibull_2P', 'Weibull_3P', 'Weibull_DS']:
                return self._create_weibull_distribution(fit_result, dist_name)
            elif 'Lognormal' in dist_name:
                return self._create_lognormal_distribution(fit_result)
            elif 'Normal' in dist_name:
                return self._create_normal_distribution(fit_result)
            elif 'Gamma' in dist_name:
                return self._create_gamma_distribution(fit_result)
            elif 'Loglogistic' in dist_name:
                return self._create_loglogistic_distribution(fit_result)
            elif 'Gumbel' in dist_name:
                return self._create_gumbel_distribution(fit_result)
            elif 'Exponential' in dist_name:
                return self._create_exponential_distribution(fit_result)
            elif 'Mixture' in dist_name:
                return self._create_mixture_model(fit_result)
            elif 'CR' in dist_name:
                return self._create_competing_risks_model(fit_result)
            
            print(f"    无法创建{dist_name}分布对象: 缺少必要参数")
            # 打印fit_result的内容，以便调试
            if isinstance(fit_result, pd.Series):
                print(f"    fit_result内容: {fit_result.to_dict()}")
            else:
                print(f"    fit_result属性: {dir(fit_result)}")
            return None
            
        except Exception as e:
            print(f"    创建{dist_name}分布对象失败: {e}")
            import traceback
            traceback.print_exc()
            return None

    def _create_mixture_from_fit_result(self, fit_result):
        """从拟合结果创建Weibull混合模型"""
        try:
            print("    创建Weibull混合模型...")
            
            # 检查是否有proportion_1属性
            if hasattr(fit_result, 'proportion_1'):
                proportion1 = fit_result.proportion_1
                proportion2 = 1 - proportion1
            else:
                # 如果没有明确的比例属性，尝试从fit_result获取
                proportion1 = 0.5
                proportion2 = 0.5
            
            # 检查是否有alpha_1和beta_1属性
            if hasattr(fit_result, 'alpha_1') and hasattr(fit_result, 'beta_1'):
                alpha1 = fit_result.alpha_1
                beta1 = fit_result.beta_1
                alpha2 = fit_result.alpha_2
                beta2 = fit_result.beta_2
                
                print(f"      组分1: alpha={alpha1:.4f}, beta={beta1:.4f}, 比例={proportion1:.4f}")
                print(f"      组分2: alpha={alpha2:.4f}, beta={beta2:.4f}, 比例={proportion2:.4f}")
                
                # 创建两个Weibull分布
                dist1 = Weibull_Distribution(alpha=alpha1, beta=beta1)
                dist2 = Weibull_Distribution(alpha=alpha2, beta=beta2)
                
                # 创建混合模型
                mixture_model = Mixture_Model(
                    distributions=[dist1, dist2], 
                    proportions=[proportion1, proportion2]
                )
                
                return mixture_model
            else:
                print("    警告: fit_result缺少alpha_1和beta_1属性")
                return None
        except Exception as e:
            print(f"    创建混合模型失败: {e}")
            return None

    def _create_competing_risks_from_fit_result(self, fit_result):
        """从拟合结果创建Weibull竞争风险模型"""
        try:
            print("    创建Weibull竞争风险模型...")
            
            # 检查是否有alpha_1和beta_1属性
            if hasattr(fit_result, 'alpha_1') and hasattr(fit_result, 'beta_1'):
                alpha1 = fit_result.alpha_1
                beta1 = fit_result.beta_1
                alpha2 = fit_result.alpha_2
                beta2 = fit_result.beta_2
                
                print(f"      风险1: alpha={alpha1:.4f}, beta={beta1:.4f}")
                print(f"      风险2: alpha={alpha2:.4f}, beta={beta2:.4f}")
                
                # 创建两个Weibull分布
                dist1 = Weibull_Distribution(alpha=alpha1, beta=beta1)
                dist2 = Weibull_Distribution(alpha=alpha2, beta=beta2)
                
                # 创建竞争风险模型
                competing_risks_model = Competing_Risks_Model(distributions=[dist1, dist2])
                
                return competing_risks_model
            else:
                print("    警告: fit_result缺少alpha_1和beta_1属性")
                return None
        except Exception as e:
            print(f"    创建竞争风险模型失败: {e}")
            return None

    def _calculate_bic_value(self, fit_result, n_failures):
        """计算BIC值"""
        try:
            # 优先使用fit_result自带的BIC属性
            if hasattr(fit_result, 'BIC'):
                bic_value = fit_result.BIC
                print(f"      使用fit_result.BIC: {bic_value:.2f}")
                return bic_value
            
            # 其次使用loglik和k计算
            elif hasattr(fit_result, 'loglik') and hasattr(fit_result, 'k'):
                loglik = fit_result.loglik
                k = fit_result.k
                # BIC = k * ln(n) - 2 * loglik
                bic_value = k * np.log(n_failures) - 2 * loglik
                print(f"      计算BIC: k={k}, n={n_failures}, loglik={loglik:.2f}, BIC={bic_value:.2f}")
                return bic_value
            
            # 如果无法计算BIC，尝试从AICc计算
            elif hasattr(fit_result, 'AICc'):
                aicc = fit_result.AICc
                # 近似转换: BIC ≈ AICc + k*log(n) - 2*k
                if hasattr(fit_result, 'k'):
                    k = fit_result.k
                    bic_value = aicc + k * np.log(n_failures) - 2 * k
                    print(f"      从AICc计算BIC: AICc={aicc:.2f}, BIC={bic_value:.2f}")
                    return bic_value
            
            # 如果以上都失败，返回一个大值
            print(f"      警告: 无法计算BIC，使用默认值10000")
            return 10000
        except Exception as e:
            print(f"      计算BIC失败: {e}")
            return 10000

    def _calculate_cumulative_failure_robust(self, distribution, model_name):
        """鲁棒地计算累计失效率"""
        try:
            if distribution is None:
                print(f"      警告: 分布对象为None，无法计算累计失效率")
                return {}
            
            cumulative_failure = {}
            
            print(f"      计算{model_name}的累计失效率...")
            
            for months in self.key_time_points:
                try:
                    # 将月份转换为天数
                    days = months * 30.44
                    # 尝试计算CDF
                    cdf = 0.0
                    
                    # 对于不同的模型类型，使用不同的计算方法
                    if model_name in ['Weibull_Mixture_2Comp', 'Weibull_CR_2Risks']:
                        # 混合模型和竞争风险模型有特殊的CDF计算方法
                        if hasattr(distribution, 'CDF'):
                            cdf = distribution.CDF(float(days))
                        else:
                            # 如果CDF方法不可用，尝试使用PDF和SF计算
                            if hasattr(distribution, 'PDF') and hasattr(distribution, 'SF'):
                                pdf = distribution.PDF(float(days))
                                sf = distribution.SF(float(days))
                                cdf = 1.0 - sf
                            else:
                                cdf = 0.0
                    else:
                        # 标准分布使用CDF方法
                        if hasattr(distribution, 'CDF'):
                            cdf = distribution.CDF(float(days))
                        else:
                            cdf = 0.0
                    
                    # 确保CDF在合理范围内
                    if cdf < 0:
                        cdf = 0.0
                    elif cdf > 1:
                        cdf = 1.0
                    
                    # 检查CDF是否为0
                    if cdf == 0 and days > 0:
                        # 尝试备选计算方法
                        print(f"        警告: 在{months}个月时CDF=0，尝试备选计算方法...")
                        # 使用数值积分近似计算CDF
                        try:
                            if hasattr(distribution, 'PDF'):
                                # 数值积分计算CDF
                                time_points = np.linspace(0.1, days, 100)
                                pdf_values = []
                                for t in time_points:
                                    try:
                                        pdf_val = distribution.PDF(float(t))
                                        pdf_values.append(pdf_val)
                                    except:
                                        pdf_values.append(0)
                                
                                if len(pdf_values) > 0:
                                    # 梯形数值积分
                                    cdf = np.trapz(pdf_values, time_points)
                                    print(f"          数值积分CDF: {cdf:.6f}")
                        except:
                            cdf = 0.0
                    
                    cumulative_failure[months] = cdf
                    
                except Exception as e:
                    print(f"      计算{months}个月累计失效率失败: {e}")
                    cumulative_failure[months] = 0.0
            
            # 检查累计失效率是否全部为0
            cdf_values = list(cumulative_failure.values())
            if all(v == 0 for v in cdf_values):
                print(f"      警告: {model_name}的所有累计失效率都为0")
                # 输出分布参数用于调试
                self._debug_distribution_params(distribution, model_name)
            
            return cumulative_failure
        except Exception as e:
            print(f"      计算累计失效率失败: {e}")
            return {}

    def _calculate_failure_rate_robust(self, distribution, model_name):
        """鲁棒地计算失效率函数"""
        try:
            if distribution is None:
                print(f"      警告: 分布对象为None，无法计算失效率")
                return {}
            
            failure_rate = {}
            
            print(f"      计算{model_name}的失效率...")
            
            for months in self.key_time_points:
                try:
                    # 将月份转换为天数
                    days = months * 30.44
                    # 计算失效率 = PDF(t) / SF(t)
                    if hasattr(distribution, 'PDF') and hasattr(distribution, 'SF'):
                        pdf = distribution.PDF(float(days))
                        sf = distribution.SF(float(days))
                        
                        # 避免除以零
                        if sf > 0:
                            rate = pdf / sf
                        else:
                            rate = 0.0
                        
                        # 确保失效率为正数
                        if rate < 0:
                            rate = 0.0
                        
                        failure_rate[months] = rate
                    else:
                        # 如果没有PDF或SF方法，使用默认值
                        failure_rate[months] = 0.0
                except Exception as e:
                    print(f"      计算{months}个月失效率失败: {e}")
                    failure_rate[months] = 0.0
            
            # 检查失效率是否全部为0
            rate_values = list(failure_rate.values())
            if all(v == 0 for v in rate_values):
                print(f"      警告: {model_name}的所有失效率都为0")
                # 输出分布参数用于调试
                self._debug_distribution_params(distribution, model_name)
            
            return failure_rate
        except Exception as e:
            print(f"      计算失效率失败: {e}")
            return {}

    def _debug_distribution_params(self, distribution, model_name):
        """调试分布参数"""
        try:
            print(f"      调试{model_name}分布参数:")
            print(f"      分布类型: {type(distribution).__name__}")
            print(f"      分布属性: {dir(distribution)}")
            
            # 尝试打印一些常见参数
            if hasattr(distribution, 'alpha'):
                print(f"      alpha: {distribution.alpha}")
            if hasattr(distribution, 'beta'):
                print(f"      beta: {distribution.beta}")
            if hasattr(distribution, 'mu'):
                print(f"      mu: {distribution.mu}")
            if hasattr(distribution, 'sigma'):
                print(f"      sigma: {distribution.sigma}")
            if hasattr(distribution, 'Lambda'):
                print(f"      Lambda: {distribution.Lambda}")
        except Exception as e:
            print(f"      调试分布参数失败: {e}")

    def _manual_probability_plot(self, ax, fit_result, dist_display_name):
        """手动绘制概率图"""
        try:
            # 提取失效时间和截尾时间
            failure_times = []
            right_censored = []
            
            # 尝试多种方式获取数据
            if hasattr(fit_result, 'failures'):
                failure_times = fit_result.failures
                print(f"  从fit_result.failures获取失效时间: {len(failure_times)}个")
            elif hasattr(fit_result, 'data'):
                # 某些版本的库可能将数据存储在data属性中
                failure_times = fit_result.data
                print(f"  从fit_result.data获取失效时间: {len(failure_times)}个")
            elif hasattr(fit_result, 'times'):
                # 其他可能的数据存储位置
                failure_times = fit_result.times
                print(f"  从fit_result.times获取失效时间: {len(failure_times)}个")
            
            # 尝试获取截尾数据
            if hasattr(fit_result, 'right_censored') and fit_result.right_censored is not None:
                right_censored = fit_result.right_censored
                print(f"  从fit_result.right_censored获取截尾时间: {len(right_censored)}个")
            elif hasattr(fit_result, 'censored'):
                # 某些版本的库可能将截尾数据存储在censored属性中
                right_censored = fit_result.censored
                print(f"  从fit_result.censored获取截尾时间: {len(right_censored)}个")
            
            # 确保数据是列表格式
            if not isinstance(failure_times, (list, np.ndarray)):
                failure_times = []
            if not isinstance(right_censored, (list, np.ndarray)):
                right_censored = []
            
            # 合并所有数据点
            all_times = list(failure_times) + list(right_censored)
            if not all_times:
                # 打印fit_result的属性，以便调试
                print(f"  调试: fit_result属性: {dir(fit_result)}")
                ax.text(0.5, 0.5, f'{dist_display_name}\n无数据点',
                       ha='center', va='center', transform=ax.transAxes)
                return
            
            # 计算经验概率
            sorted_failures = np.sort(failure_times)
            n_failures = len(sorted_failures)
            
            if n_failures == 0:
                ax.text(0.5, 0.5, f'{dist_display_name}\n无失效数据',
                       ha='center', va='center', transform=ax.transAxes)
                return
            
            # 计算经验概率（使用中位秩公式）
            empirical_probs = (np.arange(1, n_failures + 1) - 0.3) / (n_failures + 0.4)
            empirical_probs = np.clip(empirical_probs, 1e-10, 1 - 1e-10)
            
            # 根据分布类型计算理论概率
            theoretical_probs = []
            if hasattr(fit_result, 'distribution') and fit_result.distribution is not None:
                distribution = fit_result.distribution
                for t in sorted_failures:
                    try:
                        cdf = distribution.CDF(float(t))
                        theoretical_probs.append(cdf)
                    except:
                        theoretical_probs.append(0)
            else:
                # 如果没有distribution对象，尝试使用参数计算
                theoretical_probs = empirical_probs
            
            # 绘制数据点
            ax.scatter(sorted_failures, empirical_probs, color='black', label='失效数据')
            
            # 绘制截尾数据
            if right_censored:
                censored_probs = [0] * len(right_censored)
                ax.scatter(right_censored, censored_probs, color='red', marker='+', label='截尾数据')
            
            # 绘制拟合线
            if theoretical_probs:
                ax.plot(sorted_failures, theoretical_probs, 'b-', label='拟合线')
            
            # 设置坐标轴
            ax.set_xlabel('时间')
            ax.set_ylabel('累积概率')
            ax.set_title(f'{dist_display_name}概率图')
            ax.grid(True, linestyle='--', alpha=0.7)
            ax.legend()
            
        except Exception as e:
            print(f"  手动绘制{dist_display_name}概率图失败: {e}")
            ax.text(0.5, 0.5, f'{dist_display_name}\n手动绘图失败',
                   ha='center', va='center', transform=ax.transAxes)

    def plot_improvement_comparison(self, output_dir):
        """绘制改善前后对比图"""
        if not self.comparison_results:
            print("错误: 没有对比结果可绘制")
            return
        
        print("\n生成改善前后对比图...")
        
        # 获取改善前后的结果
        comparison_results = list(self.comparison_results.values())[0]
        before_result = comparison_results.get('改善前')
        after_result = comparison_results.get('改善后')
        
        if not before_result and not after_result:
            print("错误: 改善前后数据均不存在")
            return
        
        # 绘制失效率对比图
        self._plot_failure_rate_comparison(before_result, after_result, output_dir)
        
        # 绘制累计失效率对比图
        self._plot_cumulative_failure_comparison(before_result, after_result, output_dir)
        
        # 绘制改善效果直方图
        self._plot_improvement_effect_histogram(before_result, after_result, output_dir)

    def _plot_failure_rate_comparison(self, before_result, after_result, output_dir):
        """绘制失效率对比图"""
        try:
            plt.figure(figsize=(12, 6))
            
            # 确定所有时间点
            all_times = []
            if before_result and before_result['failure_rate']:
                all_times.extend(list(before_result['failure_rate'].keys()))
            if after_result and after_result['failure_rate']:
                all_times.extend(list(after_result['failure_rate'].keys()))
            all_times = sorted(list(set(all_times)))
            
            # 绘制改善前失效率
            if before_result:
                failure_rate_before = before_result['failure_rate']
                if failure_rate_before:
                    times = list(failure_rate_before.keys())
                    rates = list(failure_rate_before.values())
                    plt.plot(times, rates, 'r-', linewidth=2, label=f"改善前 - {before_result['best_model']}")
                    
                    # 标记间隔点
                    for time in times:
                        rate = failure_rate_before[time]
                        plt.annotate(f'{rate:.6f}', (time, rate), textcoords="offset points", xytext=(0,10), ha='center')
            
            # 绘制改善后失效率
            if after_result:
                failure_rate_after = after_result['failure_rate']
                if failure_rate_after:
                    times = list(failure_rate_after.keys())
                    rates = list(failure_rate_after.values())
                    plt.plot(times, rates, 'b-', linewidth=2, label=f"改善后 - {after_result['best_model']}")
                    
                    # 标记间隔点
                    for time in times:
                        rate = failure_rate_after[time]
                        plt.annotate(f'{rate:.6f}', (time, rate), textcoords="offset points", xytext=(0,-10), ha='center')
            
            plt.title('改善前后失效率对比', fontsize=16, fontweight='bold')
            plt.xlabel('时间（月）', fontsize=12)
            plt.ylabel('失效率', fontsize=12)
            # 设置x轴刻度与时间点一致
            if all_times:
                plt.xticks(all_times)
            plt.grid(True, linestyle='--', alpha=0.7)
            plt.legend(loc='best', fontsize=10)
            plt.tight_layout()
            
            # 保存图表
            plot_path = os.path.join(output_dir, '改善前后失效率对比.png')
            plt.savefig(plot_path, dpi=300, bbox_inches='tight')
            plt.show()
            print(f"失效率对比图已保存: {plot_path}")
        except Exception as e:
            print(f"绘制失效率对比图失败: {e}")

    def _plot_cumulative_failure_comparison(self, before_result, after_result, output_dir):
        """绘制累计失效率对比图"""
        try:
            plt.figure(figsize=(12, 6))
            
            # 确定所有时间点
            all_months = []
            if before_result and before_result['cumulative_failure']:
                all_months.extend(list(before_result['cumulative_failure'].keys()))
            if after_result and after_result['cumulative_failure']:
                all_months.extend(list(after_result['cumulative_failure'].keys()))
            all_months = sorted(list(set(all_months)))
            
            # 绘制改善前累计失效率
            if before_result:
                cumulative_before = before_result['cumulative_failure']
                if cumulative_before:
                    months = list(cumulative_before.keys())
                    cdfs = list(cumulative_before.values())
                    plt.plot(months, cdfs, 'r-', linewidth=2, label=f"改善前 - {before_result['best_model']}")
                    
                    # 标记间隔点
                    for month, cdf in zip(months, cdfs):
                        plt.annotate(f'{cdf*100:.2f}%', (month, cdf), textcoords="offset points", xytext=(0,10), ha='center')
            
            # 绘制改善后累计失效率
            if after_result:
                cumulative_after = after_result['cumulative_failure']
                if cumulative_after:
                    months = list(cumulative_after.keys())
                    cdfs = list(cumulative_after.values())
                    plt.plot(months, cdfs, 'b-', linewidth=2, label=f"改善后 - {after_result['best_model']}")
                    
                    # 标记间隔点
                    for month, cdf in zip(months, cdfs):
                        plt.annotate(f'{cdf*100:.2f}%', (month, cdf), textcoords="offset points", xytext=(0,-10), ha='center')
            
            plt.title('改善前后累计失效率对比', fontsize=16, fontweight='bold')
            plt.xlabel('时间（月）', fontsize=12)
            plt.ylabel('累计失效率', fontsize=12)
            # 设置y轴为百分比格式
            plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.2f}%'.format(y * 100)))
            # 设置x轴刻度与时间点一致
            if all_months:
                plt.xticks(all_months)
            plt.grid(True, linestyle='--', alpha=0.7)
            plt.legend(loc='best', fontsize=10)
            plt.tight_layout()
            
            # 保存图表
            plot_path = os.path.join(output_dir, '改善前后累计失效率对比.png')
            plt.savefig(plot_path, dpi=300, bbox_inches='tight')
            plt.show()
            print(f"累计失效率对比图已保存: {plot_path}")
        except Exception as e:
            print(f"绘制累计失效率对比图失败: {e}")

    def _plot_improvement_effect_histogram(self, before_result, after_result, output_dir):
        """绘制改善效果直方图"""
        try:
            if not before_result or not after_result:
                print("警告: 缺少改善前后数据，无法绘制改善效果直方图")
                return
            
            cumulative_before = before_result['cumulative_failure']
            cumulative_after = after_result['cumulative_failure']
            
            if not cumulative_before or not cumulative_after:
                print("警告: 缺少累计失效率数据，无法绘制改善效果直方图")
                return
            
            # 计算改善效果
            improvement_effects = {}
            for month in cumulative_before:
                if month in cumulative_after:
                    cdf_before = cumulative_before[month]
                    cdf_after = cumulative_after[month]
                    if cdf_before > 0:
                        improvement = (cdf_before - cdf_after) / cdf_before * 100
                    else:
                        improvement = 0
                    improvement_effects[month] = improvement
            
            if not improvement_effects:
                print("警告: 无法计算改善效果，无法绘制改善效果直方图")
                return
            
            # 绘制直方图
            plt.figure(figsize=(12, 6))
            
            months = list(improvement_effects.keys())
            months.sort()  # 确保月份按顺序排列
            effects = [improvement_effects[month] for month in months]
            
            plt.bar(months, effects, color='green', alpha=0.7)
            plt.plot(months, effects, 'r-', linewidth=2, marker='o')
            
            # 标记间隔点
            for month, effect in zip(months, effects):
                plt.annotate(f'{effect:.2f}%', (month, effect), textcoords="offset points", xytext=(0,10), ha='center')
            
            plt.title('改善效果直方图', fontsize=16, fontweight='bold')
            plt.xlabel('时间（月）', fontsize=12)
            plt.ylabel('改善效果 (%)', fontsize=12)
            # 设置x轴刻度与时间点一致
            plt.xticks(months)
            plt.grid(True, linestyle='--', alpha=0.7)
            plt.tight_layout()
            
            # 保存图表
            plot_path = os.path.join(output_dir, '改善效果直方图.png')
            plt.savefig(plot_path, dpi=300, bbox_inches='tight')
            plt.show()
            print(f"改善效果直方图已保存: {plot_path}")
        except Exception as e:
            print(f"绘制改善效果直方图失败: {e}")

    def save_analysis_results(self, output_dir):
        """保存分析结果到文件"""
        try:
            # 保存改善前后对比结果
            if self.comparison_results:
                comparison_results = list(self.comparison_results.values())[0]
                before_result = comparison_results.get('改善前')
                after_result = comparison_results.get('改善后')
                
                # 保存详细结果
                result_file = os.path.join(output_dir, '分析结果详细报告.txt')
                with open(result_file, 'w', encoding='utf-8') as f:
                    f.write("可靠性分析详细报告\n")
                    f.write("=" * 80 + "\n")
                    
                    # 改善前结果
                    if before_result:
                        f.write("\n改善前分析结果:\n")
                        f.write("-" * 60 + "\n")
                        f.write(f"最佳模型: {before_result['best_model']}\n")
                        f.write(f"失效事件数量: {before_result['event_count']}\n")
                        f.write(f"总样本数量: {before_result['total_count']}\n")
                        f.write(f"AD值: {before_result['AD']:.4f}\n")
                        f.write(f"AD修正值: {before_result['AD_corrected']:.4f}\n")
                        f.write(f"BIC值: {before_result['BIC']:.2f}\n")
                        f.write("\n累计失效率:\n")
                        for time, cdf in before_result['cumulative_failure'].items():
                            f.write(f"  {time}天: {cdf:.4f}\n")
                        f.write("\n失效率:\n")
                        for time, rate in before_result['failure_rate'].items():
                            f.write(f"  {time}天: {rate:.6f}\n")
                    
                    # 改善后结果
                    if after_result:
                        f.write("\n改善后分析结果:\n")
                        f.write("-" * 60 + "\n")
                        f.write(f"最佳模型: {after_result['best_model']}\n")
                        f.write(f"失效事件数量: {after_result['event_count']}\n")
                        f.write(f"总样本数量: {after_result['total_count']}\n")
                        f.write(f"AD值: {after_result['AD']:.4f}\n")
                        f.write(f"AD修正值: {after_result['AD_corrected']:.4f}\n")
                        f.write(f"BIC值: {after_result['BIC']:.2f}\n")
                        f.write("\n累计失效率:\n")
                        for time, cdf in after_result['cumulative_failure'].items():
                            f.write(f"  {time}天: {cdf:.4f}\n")
                        f.write("\n失效率:\n")
                        for time, rate in after_result['failure_rate'].items():
                            f.write(f"  {time}天: {rate:.6f}\n")
                    
                    # 计算改善效果
                    if before_result and after_result:
                        f.write("\n改善效果分析:\n")
                        f.write("-" * 60 + "\n")
                        cumulative_before = before_result['cumulative_failure']
                        cumulative_after = after_result['cumulative_failure']
                        
                        for time in cumulative_before:
                            if time in cumulative_after:
                                cdf_before = cumulative_before[time]
                                cdf_after = cumulative_after[time]
                                if cdf_before > 0:
                                    improvement = (cdf_before - cdf_after) / cdf_before * 100
                                else:
                                    improvement = 0
                                f.write(f"  {time}天: 改善效果 = {improvement:.2f}%\n")
                
                print(f"分析结果详细报告已保存: {result_file}")
        except Exception as e:
            print(f"保存分析结果失败: {e}")


def main():
    """主函数"""
    try:
        # 获取用户输入
        user_input = get_user_input()
        import_date = user_input['import_date']
        analysis_months = user_input['analysis_months']
        interval_months = user_input['interval_months']
        output_dir = user_input['output_dir']
        file_path = user_input['file_path']
        
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 读取数据
        required_columns = ['生产日期', '出货时间', '退货时间', '是否质量问题']
        data = load_reliability_data_multisheet_robust(file_path, required_columns=required_columns)
        
        if data is None:
            print("错误: 无法读取数据")
            return
        
        # 分割数据为改善前和改善后
        before_data, after_data = split_data_by_production_date(data, import_date)
        
        # 准备可靠性分析数据
        before_reliability_data = prepare_reliability_data(before_data, import_date, analysis_months)
        after_reliability_data = prepare_reliability_data(after_data, import_date, analysis_months)
        
        # 创建分析器
        analyzer = AdvancedReliabilityAnalyzer(analysis_months, interval_months)
        
        # 分析改善前后对比
        comparison_results = analyzer.analyze_improvement_comparison(before_reliability_data, after_reliability_data)
        
        if comparison_results:
            # 绘制对比图
            analyzer.plot_improvement_comparison(output_dir)
            
            # 保存分析结果
            analyzer.save_analysis_results(output_dir)
            
            print("\n" + "=" * 80)
            print("分析完成!")
            print(f"所有结果已保存到: {output_dir}")
            print("=" * 80)
        else:
            print("错误: 分析失败")
            
    except Exception as e:
        print(f"执行错误: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

