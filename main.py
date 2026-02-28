import pandas as pd
import os
import matplotlib.pyplot as plt

# 尝试导入可选库
try:
    import seaborn as sns
except ImportError:
    print("seaborn库未安装，将使用matplotlib默认样式")

# 检查pyecharts库是否安装
pyecharts_available = False
try:
    from pyecharts.charts import Map, Line
    from pyecharts import options as opts
    pyecharts_available = True
except ImportError:
    print("pyecharts库未安装，数据地图功能将不可用")

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

# 创建可视化子文件夹
def create_visualization_folders():
    """创建可视化结果的子文件夹"""
    data_files = [f for f in os.listdir('data') if f.endswith('.xls') or f.endswith('.csv')]
    for file in data_files:
        folder_path = os.path.join('visualization', file)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
    print("可视化文件夹结构创建完成")

# 数据清洗和处理函数
def clean_data(file_path):
    """清洗和处理Excel数据"""
    try:
        # 尝试读取Excel文件，不将第一行作为列名
        try:
            df = pd.read_excel(file_path, engine='openpyxl', header=None)
        except Exception:
            try:
                df = pd.read_excel(file_path, engine='xlrd', header=None)
            except Exception as e:
                print(f"无法读取Excel文件 {file_path}: {e}")
                return None
        
        # 基本清洗：去除空行和空列
        df = df.dropna(how='all')
        df = df.dropna(axis=1, how='all')
        
        # 重置索引
        df = df.reset_index(drop=True)
        
        # 针对特殊文件的处理
        file_name = os.path.basename(file_path)
        print(f"处理文件: {file_name}")
        
        # 检测表格类型并提取年份列名和数据
        def process_table(df):
            # 打印前10行数据，了解文件结构
            print("前10行数据:")
            print(df.head(10))
            
            # 寻找包含年份的行
            year_row = -1
            for i, row in df.iterrows():
                # 检查该行是否包含多个年份
                year_count = 0
                for j in range(len(df.columns)):
                    cell = row.iloc[j]
                    if pd.notna(cell):
                        cell_value = str(cell)
                        # 检查是否是年份格式（2000-2026）
                        if cell_value.startswith('20') and len(cell_value) in [4, 6]:  # 处理"2024年"这样的格式
                            year_count += 1
                        # 处理单独的年份数字，如"2025"、"2024"等
                        elif cell_value.isdigit() and len(cell_value) == 4 and cell_value.startswith('20'):
                            year_count += 1
                        # 处理包含"年"字的年份，如"2025年"、"2024年"等
                        elif '年' in cell_value and any(char.isdigit() for char in cell_value):
                            import re
                            year_match = re.search(r'20\d{2}', cell_value)
                            if year_match:
                                year_count += 1
                if year_count >= 3:  # 至少有3个年份
                    year_row = i
                    break
            print(f"找到年份行: {year_row}")
            
            # 提取年份列名
            year_columns = []
            if year_row >= 0:
                for j in range(len(df.columns)):
                    cell = df.iloc[year_row, j]
                    if pd.notna(cell):
                        cell_value = str(cell)
                        # 提取年份数字
                        import re
                        year_match = re.search(r'20\d{2}', cell_value)
                        if year_match:
                            year_columns.append(year_match.group())
            print(f"提取的年份列名: {year_columns}")
            
            # 如果没有找到年份行，尝试直接从第一行提取
            if year_row == -1 and not df.empty:
                # 检查第一行是否包含年份
                first_row = df.iloc[0]
                year_count = 0
                temp_year_columns = []
                for j in range(len(df.columns)):
                    cell = first_row.iloc[j]
                    if pd.notna(cell):
                        cell_value = str(cell)
                        # 检查是否是年份格式
                        import re
                        year_match = re.search(r'20\d{2}', cell_value)
                        if year_match:
                            temp_year_columns.append(year_match.group())
                            year_count += 1
                if year_count >= 3:
                    year_row = 0
                    year_columns = temp_year_columns
                    print(f"从第一行提取年份列名: {year_columns}")
            
            # 寻找数据开始的行
            data_start_row = -1
            if year_row >= 0:
                # 数据行通常在年份行下方
                for i in range(year_row + 1, min(len(df), year_row + 10)):
                    # 检查该行是否包含数据
                    has_data = False
                    for j in range(len(df.columns)):
                        cell = df.iloc[i, j]
                        if pd.notna(cell):
                            # 检查是否是数值或指标名称
                            cell_value = str(cell)
                            if cell_value.replace('.', '').replace(',', '').isdigit() or '游客' in cell_value or '收入' in cell_value or '指标' in cell_value or '地区' in cell_value:
                                has_data = True
                                break
                    if has_data:
                        data_start_row = i
                        break
            print(f"找到数据开始行: {data_start_row}")
            
            # 提取数据
            data_rows = []
            if year_row >= 0 and data_start_row >= 0:
                for i in range(data_start_row, len(df)):
                    # 检查是否是注释行
                    first_cell = df.iloc[i, 0]
                    if pd.notna(first_cell):
                        first_cell_str = str(first_cell)
                        if '注：' in first_cell_str or '数据来源' in first_cell_str or '数据库：' in first_cell_str or '时间：' in first_cell_str:
                            continue  # 跳过注释行
                    
                    # 提取指标名称和数据值
                    row_data = []
                    indicator_name = None
                    values = []
                    
                    for j in range(len(df.columns)):
                        cell = df.iloc[i, j]
                        if j == 0:
                            # 第一列是指标名称
                            if pd.notna(cell):
                                indicator_name = str(cell).strip()
                        else:
                            # 其他列是数据值
                            numeric_value = pd.to_numeric(cell, errors='coerce')
                            if not pd.isna(numeric_value):
                                values.append(numeric_value)
                            else:
                                values.append(0)
                    
                    # 确保数据值数量与年份列名数量一致
                    if indicator_name and values and indicator_name != '指标' and indicator_name != '地区':
                        # 截断或填充数据值，使其与年份列名数量一致
                        if len(values) > len(year_columns):
                            values = values[:len(year_columns)]
                        elif len(values) < len(year_columns):
                            values.extend([0] * (len(year_columns) - len(values)))
                        
                        # 添加到数据行
                        row_data.append(indicator_name)
                        row_data.extend(values)
                        data_rows.append(row_data)
            print(f"提取的数据行数: {len(data_rows)}")
            
            # 创建新的DataFrame
            if data_rows and year_columns:
                # 创建列名
                columns = ['指标'] + year_columns
                # 创建DataFrame
                new_df = pd.DataFrame(data_rows, columns=columns)
                print(f"新DataFrame形状: {new_df.shape}")
                print("新DataFrame前5行:")
                print(new_df.head())
                return new_df
            else:
                # 如果没有找到年份列，尝试直接处理
                if not df.empty:
                    # 清理列名
                    new_columns = []
                    for j, col in enumerate(df.columns):
                        if j == 0:
                            new_columns.append('指标')
                        else:
                            # 尝试从列名中提取年份
                            col_str = str(col)
                            import re
                            year_match = re.search(r'20\d{2}', col_str)
                            if year_match:
                                new_columns.append(year_match.group())
                            else:
                                # 尝试从单元格中提取年份
                                for i in range(min(5, len(df))):
                                    cell = df.iloc[i, j]
                                    if pd.notna(cell):
                                        cell_str = str(cell)
                                        year_match = re.search(r'20\d{2}', cell_str)
                                        if year_match:
                                            new_columns.append(year_match.group())
                                            break
                                else:
                                    new_columns.append(f'列{j+1}')
                    
                    # 清理数据
                    cleaned_data = []
                    for i, row in df.iterrows():
                        # 检查是否是注释行
                        first_cell = row.iloc[0]
                        if pd.notna(first_cell):
                            first_cell_str = str(first_cell)
                            if '注：' in first_cell_str or '数据来源' in first_cell_str or '数据库：' in first_cell_str or '时间：' in first_cell_str:
                                continue  # 跳过注释行
                        
                        # 提取数据
                        row_data = []
                        indicator_name = None
                        values = []
                        
                        for j in range(len(df.columns)):
                            cell = row.iloc[j]
                            if j == 0:
                                if pd.notna(cell):
                                    indicator_name = str(cell).strip()
                            else:
                                numeric_value = pd.to_numeric(cell, errors='coerce')
                                if not pd.isna(numeric_value):
                                    values.append(numeric_value)
                                else:
                                    values.append(0)
                        
                        if indicator_name and values and indicator_name != '指标' and indicator_name != '地区':
                            row_data.append(indicator_name)
                            row_data.extend(values)
                            cleaned_data.append(row_data)
                    
                    if cleaned_data:
                        new_df = pd.DataFrame(cleaned_data, columns=new_columns)
                        print(f"直接处理后的DataFrame形状: {new_df.shape}")
                        print("直接处理后的DataFrame前5行:")
                        print(new_df.head())
                        return new_df
                
                return df
        
        # 处理所有文件
        df = process_table(df)
        
        # 尝试将数据列转换为数值类型
        for col in df.columns[1:]:
            try:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            except Exception:
                pass
        
        # 处理空白内容：用0填充数值列的NaN值
        numeric_cols = df.select_dtypes(include=['number']).columns
        df[numeric_cols] = df[numeric_cols].fillna(0)
        
        # 去除仍然包含NaN的行（非数值列）
        df = df.dropna()
        
        return df
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")
        return None

# 保存处理后的数据
def save_cleaned_data(df, output_file):
    """保存清洗后的数据到data_final文件夹"""
    output_path = os.path.join('data_final', output_file)
    df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"数据已保存到: {output_path}")

# 生成柱线图
def generate_line_bar_chart(df, file_name):
    """生成柱线图"""
    try:
        # 检查数据是否为空
        if df.empty or df.shape[1] < 2:
            print(f"数据为空或列数不足，跳过柱线图生成")
            return
        
        # 绘制柱形图
        fig, ax = plt.subplots(figsize=(15, 8))
        
        # 获取指标名称和年份列
        indicator_name = df.columns[0]
        year_columns = df.columns[1:]
        
        # 定义颜色映射，确保每个年份使用固定的颜色
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', 
                  '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
                  '#aec7e8', '#ffbb78', '#98df8a', '#ff9896', '#c5b0d5',
                  '#c49c94', '#f7b6d2', '#c7c7c7', '#dbdb8d', '#9edae5']
        
        # 创建年份到颜色的映射字典
        year_color_map = {year: colors[i % len(colors)] for i, year in enumerate(year_columns)}
        
        # 使用堆叠柱状图
        # 初始化底部位置
        bottom = [0] * len(df)
        
        # 绘制数据
        for year in year_columns:
            ax.bar(df[indicator_name], df[year], label=year, color=year_color_map[year], bottom=bottom)
            # 更新底部位置
            bottom = [bottom[i] + df[year].iloc[i] for i in range(len(df))]
        
        ax.set_title(f'{file_name} 柱线图')
        ax.set_xlabel(indicator_name)
        ax.set_ylabel('数值')
        
        # 调整x轴标签
        plt.xticks(rotation=60, ha='right', fontsize=8)
        
        # 调整图例
        ax.legend(title='年份', bbox_to_anchor=(1.05, 1), loc='upper left')
        
        plt.tight_layout()
        
        # 保存图表
        output_path = os.path.join('visualization', file_name, 'line_bar_chart.png')
        plt.savefig(output_path, bbox_inches='tight')
        plt.close(fig)  # 关闭图表以释放内存
        print(f"柱线图已保存到: {output_path}")
    except Exception as e:
        print(f"生成柱线图时出错: {e}")
        plt.close('all')  # 发生错误时关闭所有图表

# 生成条形图
def generate_bar_chart(df, file_name):
    """生成条形图"""
    try:
        # 检查数据是否为空
        if df.empty or df.shape[1] < 2:
            print(f"数据为空或列数不足，跳过条形图生成")
            return
        
        # 绘制水平条形图
        fig, ax = plt.subplots(figsize=(15, 10))
        
        # 获取指标名称和年份列
        indicator_name = df.columns[0]
        year_columns = df.columns[1:]
        
        # 定义颜色映射，确保每个年份使用固定的颜色
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', 
                  '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
                  '#aec7e8', '#ffbb78', '#98df8a', '#ff9896', '#c5b0d5',
                  '#c49c94', '#f7b6d2', '#c7c7c7', '#dbdb8d', '#9edae5']
        
        # 创建年份到颜色的映射字典
        year_color_map = {year: colors[i % len(colors)] for i, year in enumerate(year_columns)}
        
        # 使用堆叠水平条形图
        # 初始化左侧位置
        left = [0] * len(df)
        
        # 绘制数据
        for year in year_columns:
            ax.barh(df[indicator_name], df[year], label=year, color=year_color_map[year], left=left, alpha=0.8)
            # 更新左侧位置
            left = [left[i] + df[year].iloc[i] for i in range(len(df))]
        
        ax.set_title(f'{file_name} 条形图')
        ax.set_xlabel('数值')
        ax.set_ylabel(indicator_name)
        
        # 调整y轴标签字体大小
        plt.yticks(fontsize=8)
        
        # 调整图例
        ax.legend(title='年份', bbox_to_anchor=(1.05, 1), loc='upper left')
        
        plt.tight_layout()
        
        # 保存图表
        output_path = os.path.join('visualization', file_name, 'bar_chart.png')
        plt.savefig(output_path, bbox_inches='tight')
        plt.close(fig)  # 关闭图表以释放内存
        print(f"条形图已保存到: {output_path}")
    except Exception as e:
        print(f"生成条形图时出错: {e}")
        plt.close('all')  # 发生错误时关闭所有图表

# 生成饼图
def generate_pie_chart(df, file_name):
    """生成饼图"""
    try:
        # 检查数据是否为空
        if df.empty or df.shape[1] < 2:
            print(f"数据为空或列数不足，跳过饼图生成")
            return
        
        # 取最后一年或最后一行的数据
        last_row = df.iloc[-1]
        data = last_row[1:]
        labels = df.columns[1:]
        
        # 过滤掉NaN值
        data = data.dropna()
        if data.empty:
            print(f"数据中没有有效值，跳过饼图生成")
            return
        
        # 绘制饼图
        fig, ax = plt.subplots(figsize=(10, 10))
        ax.pie(data, labels=labels, autopct='%1.1f%%', startangle=90)
        ax.set_title(f'{file_name} 饼图')
        ax.axis('equal')  # 确保饼图是圆的
        plt.tight_layout()
        
        # 保存图表
        output_path = os.path.join('visualization', file_name, 'pie_chart.png')
        plt.savefig(output_path)
        plt.close(fig)  # 关闭图表以释放内存
        print(f"饼图已保存到: {output_path}")
    except Exception as e:
        print(f"生成饼图时出错: {e}")
        plt.close('all')  # 发生错误时关闭所有图表

# 生成数据地图
def generate_data_map(df, file_name):
    """使用pyecharts生成数据地图"""
    try:
        # 检查pyecharts库是否安装
        if pyecharts_available:
            # 只对全国数据进行地图可视化，山西数据不需要做地图
            if "山西" in file_name:
                print("山西数据不需要生成地图")
                return
            
            # 检查是否是全国性数据（如居民人均收入情况、按国别分游客等），这些数据不需要生成地图
            if "居民人均" in file_name or "全国居民" in file_name or "按国别分" in file_name or "按性别、年龄和事由分" in file_name or "旅游业发展情况" in file_name or "国际旅游收入及构成" in file_name:
                print("全国性数据，跳过地图生成")
                return
            
            # 检查数据是否为空
            if df.empty:
                print("数据为空，跳过地图生成")
                return
            
            # 准备数据
            data = []
            
            # 查找包含省份数据的行
            for i, row in df.iterrows():
                if len(row) > 0:
                    try:
                        first_cell = row.iloc[0]
                        if isinstance(first_cell, str) and first_cell not in ["数据库：分省年度数据", "指标：", "时间：", "地区", "数据来源："]:
                            # 提取省份名称
                            province = first_cell.strip()
                            # 提取非零数据
                            for j in range(1, len(row)):
                                try:
                                    cell_value = row.iloc[j]
                                    if pd.notna(cell_value) and cell_value != 0:
                                        value = float(cell_value)
                                        data.append([province, value])
                                        break
                                except Exception:
                                    continue
                    except Exception:
                        continue
            
            # 检查是否有数据
            if not data:
                print("没有有关省份数据，跳过地图生成")
                return
            
            # 获取指标名称
            try:
                indicator_name = str(df.columns[0])
            except Exception:
                indicator_name = "未知指标"
            
            # 创建地图
            map_chart = Map()
            map_chart.set_global_opts(
                title_opts=opts.TitleOpts(title=f"{file_name} - {indicator_name}"),
                visualmap_opts=opts.VisualMapOpts(
                    min_=min([item[1] for item in data]),
                    max_=max([item[1] for item in data]),
                    is_piecewise=False
                )
            )
            map_chart.add("", data, maptype="china")
            
            # 保存地图
            output_path = os.path.join('visualization', file_name, 'data_map.html')
            map_chart.render(output_path)
            print(f"数据地图已保存到: {output_path}")
        else:
            print("pyecharts库未安装，跳过数据地图生成")
    except Exception as e:
        print(f"生成数据地图时出错: {e}")

# 生成折线图
def generate_line_chart(df, file_name):
    """使用pyecharts生成折线图"""
    try:
        # 检查pyecharts库是否安装
        if pyecharts_available:
            # 检查数据是否为空
            if df.empty or len(df.columns) < 2:
                print("数据为空或列数不足，跳过折线图生成")
                return
            
            # 准备数据
            years = []
            values = []
            
            # 提取年份和对应的值
            if len(df.columns) > 1:
                # 使用列名作为年份
                years = [str(col) for col in df.columns[1:]]
                
                # 提取第一行的数据作为值
                if not df.empty:
                    for col in df.columns[1:]:
                        try:
                            value = df.iloc[0][col]
                            if pd.notna(value):
                                values.append(float(value))
                            else:
                                values.append(0)
                        except Exception:
                            values.append(0)
            
            # 检查是否有数据
            if not years or not values:
                print("没有找到年份和数据，跳过折线图生成")
                return
            
            # 获取指标名称
            try:
                indicator_name = str(df.columns[0])
            except Exception:
                indicator_name = "未知指标"
            
            # 创建折线图
            line_chart = Line()
            line_chart.add_xaxis(years)
            line_chart.add_yaxis(indicator_name, values, is_smooth=True)
            line_chart.set_global_opts(
                title_opts=opts.TitleOpts(title=f"{file_name} - {indicator_name}趋势"),
                xaxis_opts=opts.AxisOpts(name="年份"),
                yaxis_opts=opts.AxisOpts(name=indicator_name)
            )
            
            # 保存折线图
            output_path = os.path.join('visualization', file_name, 'line_chart.html')
            line_chart.render(output_path)
            print(f"折线图已保存到: {output_path}")
        else:
            print("pyecharts库未安装，跳过折线图生成")
    except Exception as e:
        print(f"生成折线图时出错: {e}")

# 主函数
def main():
    """主函数"""
    # 创建可视化文件夹结构
    create_visualization_folders()
    
    # 处理所有Excel文件
    data_files = [f for f in os.listdir('data') if f.endswith('.xls')]
    
    for file in data_files:
        print(f"\n处理文件: {file}")
        
        # 构建文件路径
        input_path = os.path.join('data', file)
        output_file = f"{os.path.splitext(file)[0]}.csv"
        
        # 清洗数据
        df = clean_data(input_path)
        if df is not None:
            # 保存清洗后的数据
            save_cleaned_data(df, output_file)
            
            # 生成可视化结果
            generate_line_bar_chart(df, file)
            generate_bar_chart(df, file)
            generate_pie_chart(df, file)
            generate_data_map(df, file)
            generate_line_chart(df, file)
    
    # 处理CSV文件
    csv_files = [f for f in os.listdir('data') if f.endswith('.csv')]
    
    for file in csv_files:
        print(f"\n处理文件: {file}")
        
        # 构建文件路径
        input_path = os.path.join('data', file)
        output_file = f"{os.path.splitext(file)[0]}_cleaned.csv"
        
        # 读取CSV文件
        try:
            df = pd.read_csv(input_path, encoding='utf-8')
            
            # 基本清洗：去除空行和空列
            df = df.dropna(how='all')
            df = df.dropna(axis=1, how='all')
            
            # 重置索引
            df = df.reset_index(drop=True)
            
            # 改进列名处理
            if not df.empty:
                # 清理列名
                df.columns = [str(col).strip() if pd.notna(col) else f'列{i+1}' for i, col in enumerate(df.columns)]
            
            # 尝试将数据列转换为数值类型
            for col in df.columns[1:]:  # 假设第一列是年份或类别
                try:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                except Exception:
                    pass
            
            # 处理空白内容：用0填充数值列的NaN值
            numeric_cols = df.select_dtypes(include=['number']).columns
            df[numeric_cols] = df[numeric_cols].fillna(0)
            
            # 去除仍然包含NaN的行（非数值列）
            df = df.dropna()
            
            # 保存清洗后的数据
            save_cleaned_data(df, output_file)
            
            # 生成可视化结果
            generate_line_bar_chart(df, file)
            generate_bar_chart(df, file)
            generate_pie_chart(df, file)
            generate_data_map(df, file)
        except Exception as e:
            print(f"处理文件 {file} 时出错: {e}")

if __name__ == "__main__":
    main()