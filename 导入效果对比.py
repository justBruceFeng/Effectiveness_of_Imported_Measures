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
    Fit_Loglogistic_2P
)
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
DEFAULT_ANALYSIS_YEARS = 3
DEFAULT_INTERVAL_MONTHS = 6
DEFAULT_IMPORT_DATE = '2027-02-11'  # 默认改善措施导入时间
DEFAULT_OUTPUT_DIR = "改善对比分析结果"  # 默认输出目录

def get_user_input():
    """获取用户输入"""
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
    
    # 获取分析年限
    while True:
        analysis_years_str = input(f"请输入分析年限 (年, 默认: {DEFAULT_ANALYSIS_YEARS}): ").strip()
        if not analysis_years_str:
            analysis_years = DEFAULT_ANALYSIS_YEARS
            print(f"使用默认分析年限: {analysis_years} 年")
            break
        
        try:
            analysis_years = int(analysis_years_str)
            if 1 <= analysis_years <= 10:
                print(f"分析年限设置为: {analysis_years} 年")
                break
            else:
                print("错误: 分析年限应在1-10年之间")
        except ValueError:
            print("错误: 请输入有效的整数")
    
    # 获取时间间隔
    while True:
        interval_months_str = input(f"请输入时间间隔 (月, 默认: {DEFAULT_INTERVAL_MONTHS}): ").strip()
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
        file_path = input("请输入Excel数据文件路径 (例如: E310-260205-韶音.xlsx): ").strip()
        if not file_path:
            file_path = 'E310-260205-韶音.xlsx'
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
        'analysis_years': analysis_years,
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

def generate_time_points(years=3, interval_months=6):
    """生成时间点（天）"""
    time_points = []
    total_months = years * 12
    current_months = interval_months
    
    while current_months <= total_months:
        days = int(current_months * 30.44)
        time_points.append(days)
        current_months += interval_months
    
    return time_points

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
        engines_to_try = ['openpyxl', 'xlrd', 'calamine']
        
        for engine in engines_to_try:
            try:
                if engine == 'calamine':
                    # calamine需要特殊处理
                    try:
                        import calamine
                        df = pd.read_excel(input_path, engine=engine)
                    except ImportError:
                        continue
                else:
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

def split_data_by_import_date(data, import_date_str):
    """
    根据导入时间将数据分为改善前和改善后两组
    """
    if data is None:
        raise ValueError("输入数据为None，无法进行分组")
    
    print(f"\n开始根据导入时间 {import_date_str} 分组数据...")
    
    # 确保数据包含必要的列
    required_cols = ['出货时间', '退货时间', '是否质量问题']
    missing_cols = [col for col in required_cols if col not in data.columns]
    if missing_cols:
        print(f"警告: 数据中缺少必要的列: {missing_cols}")
        print(f"可用列: {data.columns.tolist()}")
        raise ValueError(f"数据中缺少必要的列: {missing_cols}")
    
    # 转换日期列为datetime类型
    data['出货时间'] = pd.to_datetime(data['出货时间'], errors='coerce')
    data['退货时间'] = pd.to_datetime(data['退货时间'], errors='coerce')
    import_date = pd.to_datetime(import_date_str)
    
    print(f"原始数据行数: {len(data)}")
    
    # 根据出货时间分组
    before_improvement_mask = data['出货时间'] < import_date
    after_improvement_mask = data['出货时间'] >= import_date
    
    before_data = data[before_improvement_mask].copy()
    after_data = data[after_improvement_mask].copy()
    
    print(f"改善前数据（出货时间 < {import_date_str}）: {len(before_data)} 行")
    print(f"改善后数据（出货时间 >= {import_date_str}）: {len(after_data)} 行")
    
    if len(before_data) == 0 and len(after_data) == 0:
        raise ValueError("分组后没有剩余数据，请检查日期格式和数据")
    
    return before_data, after_data

def prepare_reliability_data(data, import_date_str, analysis_years=3):
    """
    准备可靠性分析数据 - 根据导入时间处理数据
    """
    if data is None:
        raise ValueError("输入数据为None，无法进行处理")
    
    print(f"\n开始准备可靠性分析数据，导入时间: {import_date_str}")
    
    # 确保数据包含必要的列
    required_cols = ['出货时间', '退货时间', '是否质量问题']
    missing_cols = [col for col in required_cols if col not in data.columns]
    if missing_cols:
        print(f"警告: 数据中缺少必要的列: {missing_cols}")
        print(f"可用列: {data.columns.tolist()}")
        raise ValueError(f"数据中缺少必要的列: {missing_cols}")
    
    # 转换日期列为datetime类型
    data['出货时间'] = pd.to_datetime(data['出货时间'], errors='coerce')
    data['退货时间'] = pd.to_datetime(data['退货时间'], errors='coerce')
    import_date = pd.to_datetime(import_date_str)
    
    print(f"原始数据行数: {len(data)}")
    
    # 筛选逻辑：只保留出货时间在分析范围内的数据
    # 假设分析范围为导入时间前后指定年限
    start_date = import_date - pd.DateOffset(years=analysis_years)
    end_date = import_date + pd.DateOffset(years=analysis_years)
    
    time_range_mask = (data['出货时间'] >= start_date) & (data['出货时间'] <= end_date)
    filtered_data = data[time_range_mask].copy()
    print(f"时间范围筛选后: {len(filtered_data)} 行 (出货时间在 {start_date.date()} 到 {end_date.date()} 之间)")
    
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
            print(f"  出货时间: {row['出货时间']}, 退货时间: {row['退货时间']}, "
                  f"duration: {row['duration']}天, 是否质量问题: {row['是否质量问题']}")
    
    sample_censored = filtered_data[filtered_data['event'] == 0].head(2)
    if len(sample_censored) > 0:
        print("截尾数据样本:")
        for _, row in sample_censored.iterrows():
            print(f"  出货时间: {row['出货时间']}, 退货时间: {row['退货时间']}, "
                  f"duration: {row['duration']}天, 是否质量问题: {row['是否质量问题']}")
    
    return filtered_data

class AdvancedReliabilityAnalyzer:
    """高级可靠性分析器 - 简化版，只输出标准概率图"""
    
    def __init__(self, analysis_years=3, interval_months=6):
        self.analysis_years = analysis_years
        self.interval_months = interval_months
        self.key_time_points = generate_time_points(analysis_years, interval_months)
        self.analysis_results = {}
        self.comparison_results = {}  # 存储改善前后对比结果
    
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
            before_result = self.analyze_complete_data(before_data)
            if before_result:
                results['改善前'] = before_result
                print(f"改善前最佳模型: {before_result['best_model']}")
                print(f"改善前AD修正值: {before_result['AD_corrected']:.4f}")
                print(f"改善前BIC: {before_result['BIC']:.2f}")
        
        # 分析改善后数据
        if len(after_data) > 0:
            print("\n--- 改善后数据分析 ---")
            after_result = self.analyze_complete_data(after_data)
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

    def analyze_complete_data(self, data):
        """只进行完整数据分析"""
        print("\n" + "="*60)
        print("完整数据分析")
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
            # 显示一些调试信息
            print("调试信息 - 事件统计:")
            print(f"总事件数: {len(events)}")
            print(f"失效事件数: {events.sum()}")
            if hasattr(events, 'value_counts'):
                print(f"事件值分布: {events.value_counts().to_dict()}")
            return None
        
        return self._perform_advanced_reliability_analysis(durations, events)

    def _perform_advanced_reliability_analysis(self, durations, events):
        """执行高级可靠性分析"""
        print("执行完整数据分析...")
        
        # 数据验证
        valid_mask = (durations > 0) & (durations.notna())
        durations = durations[valid_mask]
        events = events[valid_mask]
        
        n_failures = events.sum()
        if n_failures < 2:
            print(f"失效数量不足: {n_failures}")
            return None
        
        print(f"分析数据: {n_failures}个失效, {len(durations)}个总样本")
        
        # 使用扩展的分布拟合
        fit_result = self._fit_comprehensive_distributions(durations, events)
        if fit_result is None:
            return None
        
        distribution, best_model_name, ad_value, ad_corrected, bic_value, fit_result_obj = fit_result
        
        # 计算累计失效率
        cumulative_failure = self._calculate_cumulative_failure_robust(distribution, best_model_name)
        
        # 计算失效率函数
        failure_rate = self._calculate_failure_rate_robust(distribution, best_model_name)
        
        # 输出模型参数用于调试
        if best_model_name == 'Weibull_Mixture_2Comp':
            self._print_mixture_model_params(distribution)
        
        # 存储结果
        result = {
            'analysis_type': '完整数据分析',
            'event_count': n_failures,
            'total_count': len(durations),
            'best_model': best_model_name,
            'AD': ad_value,
            'AD_corrected': ad_corrected,
            'BIC': bic_value,
            'cumulative_failure': cumulative_failure,
            'failure_rate': failure_rate,
            'distribution': distribution,
            'fit_result': fit_result_obj,
            'durations': durations,
            'events': events
        }
        
        self.analysis_results['完整数据分析'] = result
        return result

    def _print_mixture_model_params(self, mixture_dist):
        """打印混合模型参数"""
        print("\n混合模型参数:")
        try:
            if hasattr(mixture_dist, 'distributions') and hasattr(mixture_dist, 'proportions'):
                for i, (dist, prop) in enumerate(zip(mixture_dist.distributions, mixture_dist.proportions)):
                    print(f"  组分{i+1} (比例={prop:.4f}):")
                    if hasattr(dist, 'alpha') and hasattr(dist, 'beta'):
                        print(f"    α={dist.alpha:.2f}, β={dist.beta:.2f}")
                    
                    # 计算该组分的特征寿命
                    if hasattr(dist, 'alpha'):
                        characteristic_life = dist.alpha
                        mean_life = dist.mean
                        print(f"    特征寿命={characteristic_life:.2f}天, 平均寿命={mean_life:.2f}天")
                    
                    # 计算该组分在不同时间点的累计失效率
                    print("    累计失效率预测:")
                    for months in [3, 6, 12, 24]:
                        days = int(months * 30.44)
                        try:
                            cdf = dist.CDF(days)
                            print(f"      {months}个月: {cdf*100:.4f}%")
                        except:
                            print(f"      {months}个月: 计算失败")
        except Exception as e:
            print(f"输出混合模型参数失败: {e}")

    def _calculate_cumulative_failure_robust(self, distribution, model_name):
        """鲁棒地计算累计失效率"""
        try:
            if distribution is None:
                return {}
            
            cumulative_failure = {}
            
            for days in self.key_time_points:
                try:
                    # 对于混合模型，需要特殊处理
                    if model_name == 'Weibull_Mixture_2Comp':
                        # 混合模型的CDF计算
                        cdf = distribution.CDF(days)
                    else:
                        # 其他模型的CDF计算
                        cdf = distribution.CDF(days)
                    
                    # 确保CDF在0-1范围内
                    cdf = max(0.0, min(1.0, cdf))
                    cumulative_failure[days] = cdf
                    
                except Exception as e:
                    print(f"计算{days}天累计失效率失败: {e}")
                    cumulative_failure[days] = None
            
            return cumulative_failure
        except Exception as e:
            print(f"计算累计失效率失败: {e}")
            return {}

    def _calculate_failure_rate_robust(self, distribution, model_name):
        """鲁棒地计算失效率函数"""
        try:
            if distribution is None:
                return {}
            
            failure_rate = {}
            
            for days in self.key_time_points:
                try:
                    # 尝试使用PDF和SF计算
                    if hasattr(distribution, 'PDF') and hasattr(distribution, 'SF'):
                        pdf = distribution.PDF(days)
                        sf = distribution.SF(days)
                        if sf > 1e-10:  # 避免除以极小的数
                            failure_rate[days] = pdf / sf
                        else:
                            failure_rate[days] = 0.0
                    else:
                        # 备选计算方法：数值微分
                        epsilon = 1e-5
                        cdf_plus = distribution.CDF(days + epsilon)
                        cdf_minus = distribution.CDF(days - epsilon)
                        failure_rate[days] = (cdf_plus - cdf_minus) / (2 * epsilon)
                    
                except Exception as e:
                    # 尝试备选计算方法
                    try:
                        # 使用数值微分近似计算失效率
                        h = 0.1
                        cdf_t = distribution.CDF(days)
                        cdf_tplus = distribution.CDF(days + h)
                        hazard_rate = (cdf_tplus - cdf_t) / (h * (1 - cdf_t + 1e-10))
                        failure_rate[days] = hazard_rate
                    except:
                        failure_rate[days] = None
            
            return failure_rate
        except Exception as e:
            print(f"计算失效率失败: {e}")
            return {}

    def _fit_comprehensive_distributions(self, durations, events):
        """拟合完整分布模型 - 更新AD值计算和模型选择逻辑"""
        print("拟合完整分布模型...")
        
        failure_times = durations[events == 1].values
        failure_times = failure_times[failure_times > 0]
        
        if len(failure_times) < 2:
            print("错误: 有效失效数据不足")
            return None
        
        right_censored = durations[events == 0].values
        right_censored = right_censored[right_censored > 0]
        
        print(f"拟合数据: {len(failure_times)}个失效时间, {len(right_censored)}个截尾时间")
        
        # 完整的分布列表
        distributions_to_fit = [
            ('Weibull_2P', Fit_Weibull_2P),
            ('Weibull_3P', Fit_Weibull_3P),
            ('Lognormal_2P', Fit_Lognormal_2P),
            ('Normal_2P', Fit_Normal_2P),
            ('Gamma_2P', Fit_Gamma_2P),
            ('Loglogistic_2P', Fit_Loglogistic_2P),
            ('Gumbel_2P', Fit_Gumbel_2P),
            ('Exponential_2P', Fit_Exponential_2P),
        ]
        
        # 添加高级分布（如果数据量足够）
        if len(failure_times) >= 20:
            distributions_to_fit.extend([
                ('Weibull_Mixture_2Comp', Fit_Weibull_Mixture),
                ('Weibull_CR_2Risks', Fit_Weibull_CR)
            ])
        
        all_results = []
        
        for dist_name, fitter in distributions_to_fit:
            try:
                print(f"  拟合 {dist_name}...")
                
                fit_result = None
                
                # 特殊处理高级分布模型
                if dist_name in ['Weibull_Mixture_2Comp', 'Weibull_CR_2Risks']:
                    try:
                        if dist_name == 'Weibull_Mixture_2Comp':
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
                        print(f"    {dist_name} 第一次拟合失败: {e1}")
                        try:
                            if dist_name == 'Weibull_Mixture_2Comp':
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
                            print(f"    {dist_name} 最终拟合失败: {e2}")
                            continue
                else:
                    # 基础分布使用标准拟合方法
                    fit_result = fitter(
                        failures=failure_times,
                        right_censored=right_censored if len(right_censored) > 0 else None,
                        show_probability_plot=False,
                        print_results=False
                    )
                
                if fit_result is None:
                    continue
                
                distribution_obj = self._create_distribution_object(fit_result, dist_name)
                if distribution_obj is None:
                    continue
                
                # 计算AD统计量和BIC
                ad_value, ad_corrected = self._calculate_ad_statistic_corrected(durations, events, distribution_obj)
                bic_value = self._calculate_bic_value(fit_result, len(failure_times))
                
                all_results.append((dist_name, ad_value, ad_corrected, bic_value, distribution_obj, fit_result))
                
                print(f"    {dist_name}: AD={ad_value:.4f}, AD修正={ad_corrected:.4f}, BIC={bic_value:.2f}")
                
            except Exception as e:
                print(f"  拟合 {dist_name} 失败: {e}")
                continue
        
        if not all_results:
            print("错误: 所有分布拟合失败")
            return None
        
        # 根据新规则选择最佳模型：先判断AD修正值，差异5%以内再判断BIC
        best_model = self._select_best_model_by_new_criteria(all_results)
        
        if best_model is None:
            print("错误: 无法选择最佳模型")
            return None
        
        dist_name, ad_value, ad_corrected, bic_value, distribution_obj, fit_result = best_model
        
        # 输出所有模型比较结果
        print(f"\n所有模型比较结果:")
        print("-" * 80)
        print(f"{'模型':<25} {'AD值':<10} {'AD修正值':<12} {'BIC':<10} {'选择状态':<10}")
        print("-" * 80)
        for result in sorted(all_results, key=lambda x: x[2]):  # 按AD修正值排序
            marker = "★" if result[0] == dist_name else " "
            print(f"{marker} {result[0]:<23} {result[1]:>8.4f} {result[2]:>10.4f} {result[3]:>10.2f} {'最佳' if result[0] == dist_name else ''}")
        
        return distribution_obj, dist_name, ad_value, ad_corrected, bic_value, fit_result

    def _select_best_model_by_new_criteria(self, all_results):
        """根据新规则选择最佳模型：先判断AD修正值，差异5%以内再判断BIC"""
        if not all_results:
            return None
        
        # 按AD修正值排序
        sorted_by_ad = sorted(all_results, key=lambda x: x[2])  # x[2]是AD修正值
        
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

    def _create_distribution_object(self, fit_result, dist_name):
        """创建分布对象"""
        try:
            if dist_name == 'Weibull_2P':
                return Weibull_Distribution(alpha=fit_result.alpha, beta=fit_result.beta)
            elif dist_name == 'Weibull_3P':
                return Weibull_Distribution(alpha=fit_result.alpha, beta=fit_result.beta, gamma=fit_result.gamma)
            elif dist_name == 'Lognormal_2P':
                return Lognormal_Distribution(mu=fit_result.mu, sigma=fit_result.sigma)
            elif dist_name == 'Normal_2P':
                return Normal_Distribution(mu=fit_result.mu, sigma=fit_result.sigma)
            elif dist_name == 'Gamma_2P':
                return Gamma_Distribution(alpha=fit_result.alpha, beta=fit_result.beta)
            elif dist_name == 'Loglogistic_2P':
                return Loglogistic_Distribution(alpha=fit_result.alpha, beta=fit_result.beta)
            elif dist_name == 'Gumbel_2P':
                return Gumbel_Distribution(mu=fit_result.mu, sigma=fit_result.sigma)
            elif dist_name == 'Exponential_2P':
                return Exponential_Distribution(Lambda=fit_result.Lambda)
            elif dist_name == 'Weibull_Mixture_2Comp':
                if hasattr(fit_result, 'distribution'):
                    return fit_result.distribution
                else:
                    return self._create_mixture_from_fit_result(fit_result)
            elif dist_name == 'Weibull_CR_2Risks':
                if hasattr(fit_result, 'distribution'):
                    return fit_result.distribution
                else:
                    return self._create_competing_risks_from_fit_result(fit_result)
            else:
                return None
        except Exception as e:
            print(f"创建{dist_name}分布对象失败: {e}")
            return None

    def _create_mixture_from_fit_result(self, fit_result):
        """从拟合结果创建Weibull混合模型"""
        try:
            if hasattr(fit_result, 'alpha_1') and hasattr(fit_result, 'beta_1'):
                dist1 = Weibull_Distribution(alpha=fit_result.alpha_1, beta=fit_result.beta_1)
                dist2 = Weibull_Distribution(alpha=fit_result.alpha_2, beta=fit_result.beta_2)
                proportions = [fit_result.proportion_1, fit_result.proportion_2]
                return Mixture_Model(distributions=[dist1, dist2], proportions=proportions)
            else:
                return None
        except Exception as e:
            print(f"创建混合模型失败: {e}")
            return None

    def _create_competing_risks_from_fit_result(self, fit_result):
        """从拟合结果创建Weibull竞争风险模型"""
        try:
            if hasattr(fit_result, 'alpha_1') and hasattr(fit_result, 'beta_1'):
                dist1 = Weibull_Distribution(alpha=fit_result.alpha_1, beta=fit_result.beta_1)
                dist2 = Weibull_Distribution(alpha=fit_result.alpha_2, beta=fit_result.beta_2)
                return Competing_Risks_Model(distributions=[dist1, dist2])
            else:
                return None
        except Exception as e:
            print(f"创建竞争风险模型失败: {e}")
            return None

    def _calculate_bic_value(self, fit_result, n_failures):
        """计算BIC值"""
        try:
            if hasattr(fit_result, 'BIC'):
                return fit_result.BIC
            elif hasattr(fit_result, 'loglik') and hasattr(fit_result, 'k'):
                # BIC = k * ln(n) - 2 * loglik
                return fit_result.k * np.log(n_failures) - 2 * fit_result.loglik
            else:
                # 如果无法计算BIC，返回一个大值
                return 10000
        except Exception as e:
            print(f"计算BIC失败: {e}")
            return 10000

    def plot_standard_probability_plots(self, comparison_results, output_dir="output"):
        """绘制标准概率图 - 使用reliability库的Fitters现有方法"""
        if not comparison_results:
            print("没有对比结果可绘制")
            return
        
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        for group_name, result in comparison_results.items():
            print(f"\n绘制{group_name}的标准概率图...")
            
            durations = result.get('durations')
            events = result.get('events')
            
            if durations is None or events is None:
                print(f"警告: {group_name}缺少持续时间或事件数据")
                continue
            
            # 准备失效时间和截尾时间
            failure_times = durations[events == 1].values
            right_censored = durations[events == 0].values
            
            if len(failure_times) < 2:
                print(f"警告: {group_name}失效数据不足，无法绘制概率图")
                continue
            
            # 获取最佳模型
            best_model = result.get('best_model', 'Weibull_2P')
            
            # 根据最佳模型类型选择合适的Fitter
            try:
                fit_result = None
                
                if best_model in ['Weibull_2P', 'Weibull_3P', 'Weibull_Mixture_2Comp']:
                    # 使用Weibull拟合
                    fit_result = Fit_Weibull_2P(
                        failures=failure_times,
                        right_censored=right_censored if len(right_censored) > 0 else None,
                        show_probability_plot=True,  # 显示概率图
                        print_results=True
                    )
                    
                elif best_model == 'Lognormal_2P':
                    fit_result = Fit_Lognormal_2P(
                        failures=failure_times,
                        right_censored=right_censored if len(right_censored) > 0 else None,
                        show_probability_plot=True,
                        print_results=True
                    )
                    
                elif best_model == 'Normal_2P':
                    fit_result = Fit_Normal_2P(
                        failures=failure_times,
                        right_censored=right_censored if len(right_censored) > 0 else None,
                        show_probability_plot=True,
                        print_results=True
                    )
                    
                elif best_model == 'Gamma_2P':
                    fit_result = Fit_Gamma_2P(
                        failures=failure_times,
                        right_censored=right_censored if len(right_censored) > 0 else None,
                        show_probability_plot=True,
                        print_results=True
                    )
                    
                elif best_model == 'Loglogistic_2P':
                    fit_result = Fit_Loglogistic_2P(
                        failures=failure_times,
                        right_censored=right_censored if len(right_censored) > 0 else None,
                        show_probability_plot=True,
                        print_results=True
                    )
                    
                elif best_model == 'Gumbel_2P':
                    fit_result = Fit_Gumbel_2P(
                        failures=failure_times,
                        right_censored=right_censored if len(right_censored) > 0 else None,
                        show_probability_plot=True,
                        print_results=True
                    )
                    
                elif best_model == 'Exponential_2P':
                    fit_result = Fit_Exponential_2P(
                        failures=failure_times,
                        right_censored=right_censored if len(right_censored) > 0 else None,
                        show_probability_plot=True,
                        print_results=True
                    )
                else:
                    print(f"警告: 未知模型类型 {best_model}，使用Weibull分布")
                    fit_result = Fit_Weibull_2P(
                        failures=failure_times,
                        right_censored=right_censored if len(right_censored) > 0 else None,
                        show_probability_plot=True,
                        print_results=True
                    )
                
                # 获取当前图形并保存
                plt.figure(plt.gcf().number)  # 获取当前图形
                plt.title(f'{group_name} - {best_model}概率图', fontsize=14, fontweight='bold')
                plt.tight_layout()
                
                # 保存图表
                plot_path = os.path.join(output_dir, f'{group_name}_标准概率图.png')
                plt.savefig(plot_path, dpi=300, bbox_inches='tight')
                plt.show()
                print(f"{group_name}标准概率图已保存: {plot_path}")
                
            except Exception as e:
                print(f"绘制{group_name}标准概率图失败: {e}")
                import traceback
                traceback.print_exc()
    
    def plot_comparison_charts_only(self, comparison_results, output_dir="output"):
        """只绘制改善前后对比图表"""
        if not comparison_results:
            print("没有对比结果可绘制")
            return
        
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        # 准备数据
        time_points_days = self.key_time_points
        time_points_months = [days / 30.44 for days in time_points_days]
        
        # 关键时间点（3, 6, 9, 12个月，直到分析年限）
        key_months = [3, 6, 9, 12] + [i for i in range(15, self.analysis_years*12 + 1, 3)]
        key_days = [int(month * 30.44) for month in key_months]
        
        # 1. 失效率曲线对比
        plt.figure(figsize=(12, 8))
        
        colors = {'改善前': 'red', '改善后': 'blue'}
        markers = {'改善前': 'o', '改善后': 's'}
        
        for group_name, result in comparison_results.items():
            if 'failure_rate' in result and result['failure_rate']:
                failure_rates = []
                valid_times = []
                
                for days in time_points_days:
                    if days in result['failure_rate'] and result['failure_rate'][days] is not None:
                        failure_rates.append(result['failure_rate'][days])
                        valid_times.append(days / 30.44)  # 转换为月
                
                if failure_rates:
                    plt.plot(valid_times, failure_rates, 
                            label=f"{group_name} ({result['best_model']})",
                            color=colors.get(group_name, 'red' if group_name == '改善前' else 'blue'),
                            linewidth=2, marker=markers.get(group_name, 'o'))
        
        # 标记关键时间点
        for month in key_months:
            if month <= self.analysis_years * 12:  # 只标记在分析年限内的时间点
                plt.axvline(x=month, color='gray', linestyle='--', alpha=0.3, linewidth=0.5)
        
        plt.xlabel('时间 (月)', fontsize=12)
        plt.ylabel('失效率 (次/单位时间)', fontsize=12)
        plt.title('改善前后失效率曲线对比', fontsize=14, fontweight='bold')
        plt.legend(fontsize=10)
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        
        # 保存图表
        failure_rate_path = os.path.join(output_dir, '失效率曲线对比.png')
        plt.savefig(failure_rate_path, dpi=300)
        plt.show()
        print(f"失效率曲线对比图已保存: {failure_rate_path}")
        
        # 2. 累计失效率曲线对比
        plt.figure(figsize=(12, 8))
        
        for group_name, result in comparison_results.items():
            if 'cumulative_failure' in result and result['cumulative_failure']:
                cumulative_rates = []
                valid_times = []
                
                for days in time_points_days:
                    if days in result['cumulative_failure'] and result['cumulative_failure'][days] is not None:
                        cumulative_rates.append(result['cumulative_failure'][days] * 100)  # 转换为百分比
                        valid_times.append(days / 30.44)  # 转换为月
                
                if cumulative_rates:
                    plt.plot(valid_times, cumulative_rates, 
                            label=f"{group_name} ({result['best_model']})",
                            color=colors.get(group_name, 'red' if group_name == '改善前' else 'blue'),
                            linewidth=2, marker=markers.get(group_name, 'o'))
        
        # 标记关键时间点
        for month in key_months:
            if month <= self.analysis_years * 12:  # 只标记在分析年限内的时间点
                plt.axvline(x=month, color='gray', linestyle='--', alpha=0.3, linewidth=0.5)
        
        plt.xlabel('时间 (月)', fontsize=12)
        plt.ylabel('累计失效率 (%)', fontsize=12)
        plt.title('改善前后累计失效率曲线对比', fontsize=14, fontweight='bold')
        plt.legend(fontsize=10)
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        
        # 保存图表
        cumulative_path = os.path.join(output_dir, '累计失效率曲线对比.png')
        plt.savefig(cumulative_path, dpi=300)
        plt.show()
        print(f"累计失效率曲线对比图已保存: {cumulative_path}")
        
        # 3. 创建详细对比表格
        self._create_comparison_table(comparison_results, key_days, output_dir)
        
        return failure_rate_path, cumulative_path

    def _create_comparison_table(self, comparison_results, key_days, output_dir):
        """创建详细对比表格 - 修复版本"""
        print("\n" + "="*80)
        print("改善前后详细对比数据表")
        print("="*80)
        
        # 准备表格数据
        table_data = []
        
        for group_name, result in comparison_results.items():
            if 'cumulative_failure' in result and 'failure_rate' in result:
                row = {'组别': group_name}
                row['最佳模型'] = result['best_model']
                row['失效数/总数'] = f"{result['event_count']}/{result['total_count']}"
                row['AD修正值'] = f"{result['AD_corrected']:.4f}"
                row['BIC值'] = f"{result['BIC']:.2f}"
                
                # 使用固定的月份列表，而不是从天数转换
                fixed_months = [3, 6, 9, 12] + [i for i in range(15, self.analysis_years*12 + 1, 3)]
                
                # 只添加在分析年限内的固定月份
                for month in fixed_months:
                    if month <= self.analysis_years * 12:
                        # 计算对应的天数
                        days = int(month * 30.44)
                        
                        # 累计失效率
                        cum_fail_key = f"{month}个月累计失效率(%)"
                        if days in result['cumulative_failure'] and result['cumulative_failure'][days] is not None:
                            # 显示更多小数位
                            cum_value = result['cumulative_failure'][days] * 100
                            row[cum_fail_key] = f"{cum_value:.6f}"
                        else:
                            row[cum_fail_key] = "N/A"
                        
                        # 失效率
                        fail_rate_key = f"{month}个月失效率"
                        if days in result['failure_rate'] and result['failure_rate'][days] is not None:
                            row[fail_rate_key] = f"{result['failure_rate'][days]:.6e}"
                        else:
                            row[fail_rate_key] = "N/A"
                
                table_data.append(row)
        
        # 创建DataFrame并显示
        if table_data:
            df_comparison = pd.DataFrame(table_data)
            
            # 重新排序列顺序
            columns_order = ['组别', '最佳模型', '失效数/总数', 'AD修正值', 'BIC值']
            
            # 添加时间点列 - 使用固定月份
            fixed_months = [3, 6, 9, 12] + [i for i in range(15, self.analysis_years*12 + 1, 3)]
            for month in fixed_months:
                if month <= self.analysis_years * 12:
                    columns_order.extend([f"{month}个月累计失效率(%)", f"{month}个月失效率"])
            
            # 只保留实际存在的列
            existing_columns = [col for col in columns_order if col in df_comparison.columns]
            df_comparison = df_comparison[existing_columns]
            
            # 输出到控制台
            pd.set_option('display.max_columns', None)
            pd.set_option('display.width', None)
            print("\n详细对比表格:")
            print(df_comparison.to_string(index=False))
            
            # 保存到Excel
            excel_path = os.path.join(output_dir, '改善前后对比分析结果.xlsx')
            df_comparison.to_excel(excel_path, index=False)
            print(f"\n详细对比表格已保存: {excel_path}")
            
            # 计算改善效果
            self._calculate_improvement_effect(df_comparison, fixed_months, output_dir)
            
            return df_comparison
        else:
            print("错误: 无法创建对比表格")
            return None

    def _calculate_improvement_effect(self, df_comparison, fixed_months, output_dir):
        """计算改善效果"""
        print("\n" + "="*80)
        print("改善效果分析")
        print("="*80)
        
        if len(df_comparison) != 2:
            print("警告: 需要改善前和改善后两组数据进行比较")
            return
        
        # 提取数据
        before_data = df_comparison[df_comparison['组别'] == '改善前'].iloc[0]
        after_data = df_comparison[df_comparison['组别'] == '改善后'].iloc[0]
        
        improvement_results = []
        
        # 只使用固定月份
        for month in fixed_months:
            if month <= self.analysis_years * 12:
                cum_before_key = f"{month}个月累计失效率(%)"
                cum_after_key = f"{month}个月累计失效率(%)"
                
                if cum_before_key in before_data and cum_after_key in after_data:
                    try:
                        cum_before = float(before_data[cum_before_key])
                        cum_after = float(after_data[cum_after_key])
                        cum_improvement = ((cum_before - cum_after) / cum_before * 100) if cum_before > 0 else 0
                        
                        improvement_results.append({
                            '时间点(月)': month,
                            '改善前累计失效率(%)': cum_before,
                            '改善后累计失效率(%)': cum_after,
                            '绝对改善(%)': cum_before - cum_after,
                            '相对改善(%)': cum_improvement
                        })
                    except (ValueError, TypeError):
                        pass
        
        if improvement_results:
            df_improvement = pd.DataFrame(improvement_results)
            
            print("\n累计失效率改善效果:")
            print(df_improvement.to_string(index=False))
            
            # 保存改善效果
            improvement_path = os.path.join(output_dir, '改善效果分析.xlsx')
            df_improvement.to_excel(improvement_path, index=False)
            print(f"改善效果分析已保存: {improvement_path}")
            
            # 绘制改善效果图
            self._plot_improvement_chart(df_improvement, output_dir)

    def _plot_improvement_chart(self, df_improvement, output_dir):
        """绘制改善效果图"""
        plt.figure(figsize=(10, 6))
        
        months = df_improvement['时间点(月)']
        relative_improvement = df_improvement['相对改善(%)']
        
        bars = plt.bar(months, relative_improvement, color='green', alpha=0.7)
        
        # 在柱状图上添加数值标签
        for bar, value in zip(bars, relative_improvement):
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                    f'{value:.1f}%', ha='center', va='bottom', fontsize=9)
        
        plt.xlabel('时间 (月)', fontsize=12)
        plt.ylabel('相对改善 (%)', fontsize=12)
        plt.title('改善措施效果 - 累计失效率相对改善百分比', fontsize=14, fontweight='bold')
        plt.grid(True, alpha=0.3, axis='y')
        plt.xticks(months)
        plt.tight_layout()
        
        # 保存图表
        improvement_chart_path = os.path.join(output_dir, '改善效果柱状图.png')
        plt.savefig(improvement_chart_path, dpi=300)
        plt.show()
        print(f"改善效果柱状图已保存: {improvement_chart_path}")

    def print_comprehensive_comparison(self, comparison_results):
        """输出综合对比结果报告 - 根据分析年限动态调整"""
        if not comparison_results:
            print("没有可用的对比分析结果")
            return
        
        print("\n" + "="*100)
        print("改善前后可靠性分析综合对比报告")
        print("="*100)
        
        for group_name, result in comparison_results.items():
            print(f"\n{'-'*50}")
            print(f"{group_name}数据分析结果")
            print(f"{'-'*50}")
            
            print(f"最佳模型: {result['best_model']}")
            print(f"失效数/总数: {result['event_count']}/{result['total_count']}")
            print(f"AD值: {result['AD']:.4f}")
            print(f"AD修正值: {result['AD_corrected']:.4f}")
            print(f"BIC值: {result['BIC']:.2f}")
            
            # 根据分析年限生成关键时间点
            print(f"\n关键时间点累计失效率预测:")
            print(f"{'时间点':<10} {'累计失效率(%)':<15}")
            print(f"{'-'*25}")
            
            # 生成3,6,9,12个月直到分析年限结束
            for months in range(3, self.analysis_years*12 + 1, 3):
                days = int(months * 30.44)
                if days in result['cumulative_failure'] and result['cumulative_failure'][days] is not None:
                    failure_rate = result['cumulative_failure'][days] * 100
                    if failure_rate < 0.0001:  # 极小的值
                        print(f"{months:>4}个月    {'<0.0001%':>12}")
                    else:
                        print(f"{months:>4}个月    {failure_rate:>12.4f}%")
                else:
                    print(f"{months:>4}个月    {'N/A':>12}")
        
        print(f"\n{'='*100}")
        print("模型选择说明:")
        print("1. 首先比较AD修正值，选择AD修正值最小的模型")
        print("2. 如果多个模型的AD修正值差异在5%以内，再比较BIC值，选择BIC值最小的模型")
        print("3. AD修正值计算公式: A2* = A2 × (1 + 0.75/n + 2.25/n²)")
        print("4. 累计失效率显示说明: 当值小于0.0001%时显示为<0.0001%")
        print(f"{'='*100}")

# 主程序
def main():
    try:
        # 获取用户输入
        user_config = get_user_input()
        
        # 定义必需的列名
        REQUIRED_COLUMNS = ['duration', '是否质量问题', '出货时间', '退货时间']
        
        # 加载数据
        print("\n开始加载数据...")
        data = load_reliability_data_multisheet_robust(
            file_path=user_config['file_path'],
            sheet_names=None,
            required_columns=REQUIRED_COLUMNS
        )
        
        if data is None:
            print("数据加载失败，无法继续分析")
            return
        
        # 准备可靠性分析数据
        data = prepare_reliability_data(
            data, 
            user_config['import_date'], 
            user_config['analysis_years']
        )
        
        # 根据导入时间分割数据
        before_data, after_data = split_data_by_import_date(data, user_config['import_date'])
        
        if len(before_data) == 0 and len(after_data) == 0:
            print("错误: 改善前后都没有数据")
            return
        
        # 创建高级分析器
        analyzer = AdvancedReliabilityAnalyzer(
            analysis_years=user_config['analysis_years'],
            interval_months=user_config['interval_months']
        )
        
        # 分析改善前后对比
        print("\n开始分析改善前后对比...")
        comparison_results = analyzer.analyze_improvement_comparison(
            before_data, after_data, "改善措施效果对比"
        )
        
        if comparison_results is None:
            print("对比分析失败，请检查数据")
            return
        
        # 输出综合结果报告
        analyzer.print_comprehensive_comparison(comparison_results)
        
        # 绘制标准概率图
        print("\n生成标准概率图...")
        analyzer.plot_standard_probability_plots(
            comparison_results, 
            output_dir=user_config['output_dir']
        )
        
        # 绘制对比图表
        print("\n生成对比图表...")
        analyzer.plot_comparison_charts_only(
            comparison_results, 
            output_dir=user_config['output_dir']
        )
        
        print("\n" + "="*80)
        print("分析完成！")
        print("="*80)
        print(f"分析配置:")
        print(f"1. 改善措施导入时间: {user_config['import_date']}")
        print(f"2. 分析年限: {user_config['analysis_years']} 年")
        print(f"3. 时间间隔: {user_config['interval_months']} 个月")
        print(f"4. 数据文件: {user_config['file_path']}")
        print(f"5. 输出目录: {user_config['output_dir']}")
        print(f"\n结果说明:")
        print(f"1. 改善前: 生产日期早于{user_config['import_date']}的产品")
        print(f"2. 改善后: 生产日期晚于或等于{user_config['import_date']}的产品")
        print(f"3. 所有图表和详细数据已保存到'{user_config['output_dir']}'文件夹")
        print("="*80)
            
    except Exception as e:
        print(f"程序执行错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()