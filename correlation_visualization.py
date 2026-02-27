import pandas as pd
import os
import matplotlib.pyplot as plt
from pyecharts.charts import Line, Scatter, HeatMap, Map
from pyecharts import options as opts

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

# 创建可视化文件夹
def create_correlation_folder():
    """创建关联分析可视化文件夹"""
    folder_path = os.path.join('visualization', 'correlation_analysis')
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    return folder_path

# 加载数据
def load_data():
    """加载旅游业数据和生产总值数据"""
    tourism_file = os.path.join('data_final', '全国国际旅游(外汇)收入.csv')
    gdp_file = os.path.join('data_final', '全国地区生产总值.csv')
    
    tourism_df = pd.read_csv(tourism_file, encoding='utf-8-sig')
    gdp_df = pd.read_csv(gdp_file, encoding='utf-8-sig')
    
    return tourism_df, gdp_df

# 处理数据
def process_data(tourism_df, gdp_df):
    """处理数据，计算旅游业收入占GDP的比例"""
    # 提取年份列（2006-2019，因为2020-2025数据可能不完整）
    years = [str(year) for year in range(2006, 2020)]
    
    # 提取省份数据
    provinces = tourism_df['指标'].tolist()
    
    # 计算各省份各年份的旅游业收入占GDP的比例
    correlation_data = []
    for province in provinces:
        for year in years:
            tourism_value = tourism_df.loc[tourism_df['指标'] == province, year].values[0]
            gdp_value = gdp_df.loc[gdp_df['指标'] == province, year].values[0]
            
            if gdp_value > 0:
                ratio = (tourism_value / gdp_value) * 100  # 转换为百分比
                correlation_data.append([province, year, tourism_value, gdp_value, ratio])
    
    # 创建DataFrame
    df = pd.DataFrame(correlation_data, columns=['省份', '年份', '旅游收入', 'GDP', '旅游收入占GDP比例'])
    return df, years, provinces

# 生成双轴折线图
def generate_dual_axis_line_chart(df, years, output_folder):
    """生成双轴折线图，展示旅游业收入与GDP的趋势"""
    # 计算全国平均值（只包含数值列）
    numeric_cols = ['旅游收入', 'GDP', '旅游收入占GDP比例']
    national_avg = df.groupby('年份')[numeric_cols].mean().reset_index()
    
    # 计算山西数据
    shanxi_data = df[df['省份'] == '山西省']
    
    # 创建图表
    fig, ax1 = plt.subplots(figsize=(15, 8))
    
    # 绘制GDP折线
    ax1.plot(national_avg['年份'], national_avg['GDP'], 'b-', label='全国平均GDP')
    ax1.plot(shanxi_data['年份'], shanxi_data['GDP'], 'b--', label='山西GDP')
    ax1.set_xlabel('年份')
    ax1.set_ylabel('GDP (亿元)', color='b')
    ax1.tick_params(axis='y', labelcolor='b')
    
    # 创建第二个y轴
    ax2 = ax1.twinx()
    ax2.plot(national_avg['年份'], national_avg['旅游收入'], 'r-', label='全国平均旅游收入')
    ax2.plot(shanxi_data['年份'], shanxi_data['旅游收入'], 'r--', label='山西旅游收入')
    ax2.set_ylabel('旅游收入 (百万美元)', color='r')
    ax2.tick_params(axis='y', labelcolor='r')
    
    # 添加标题和图例
    plt.title('全国与山西旅游业收入与GDP趋势对比')
    ax1.legend(loc='upper left')
    ax2.legend(loc='upper right')
    
    # 保存图表
    output_path = os.path.join(output_folder, 'dual_axis_line_chart.png')
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()
    print(f"双轴折线图已保存到: {output_path}")

# 生成散点图
def generate_scatter_chart(df, output_folder):
    """生成散点图，分析旅游业收入与GDP的相关性"""
    # 计算2019年数据（最近的完整年份）
    data_2019 = df[df['年份'] == '2019']
    
    # 创建图表
    fig, ax = plt.subplots(figsize=(15, 10))
    
    # 绘制散点
    scatter = ax.scatter(data_2019['GDP'], data_2019['旅游收入'], s=100, alpha=0.7)
    
    # 添加省份标签
    for i, row in data_2019.iterrows():
        ax.annotate(row['省份'], (row['GDP'], row['旅游收入']), fontsize=8)
    
    # 添加标题和标签
    plt.title('2019年各省份旅游业收入与GDP相关性分析')
    plt.xlabel('GDP (亿元)')
    plt.ylabel('旅游收入 (百万美元)')
    
    # 保存图表
    output_path = os.path.join(output_folder, 'scatter_chart.png')
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()
    print(f"散点图已保存到: {output_path}")

# 生成热力图
def generate_heatmap(df, provinces, years, output_folder):
    """生成热力图，展示各省份旅游业收入占GDP比例的变化"""
    # 准备数据
    heatmap_data = []
    for i, province in enumerate(provinces):
        province_data = df[df['省份'] == province]
        for j, year in enumerate(years):
            year_data = province_data[province_data['年份'] == year]
            if not year_data.empty:
                ratio = year_data['旅游收入占GDP比例'].values[0]
                heatmap_data.append([j, i, round(ratio, 4)])
    
    # 创建热力图
    heatmap = HeatMap()
    heatmap.add_xaxis(years)
    heatmap.add_yaxis(series_name="", yaxis_data=provinces)
    heatmap.add(series_name="旅游收入占GDP比例", data=heatmap_data, label_opts=opts.LabelOpts(is_show=True, position="inside"))
    heatmap.set_global_opts(
        title_opts=opts.TitleOpts(title="各省份旅游业收入占GDP比例变化热力图"),
        xaxis_opts=opts.AxisOpts(type_="category", splitarea_opts=opts.SplitAreaOpts(is_show=True)),
        yaxis_opts=opts.AxisOpts(type_="category", splitarea_opts=opts.SplitAreaOpts(is_show=True)),
        visualmap_opts=opts.VisualMapOpts(min_=0, max_=0.5)
    )
    
    # 保存热力图
    output_path = os.path.join(output_folder, 'heatmap.html')
    heatmap.render(output_path)
    print(f"热力图已保存到: {output_path}")

# 生成地图可视化
def generate_map_visualization(df, output_folder):
    """生成地图可视化，展示2019年各省份旅游业收入占GDP比例"""
    # 计算2019年数据
    data_2019 = df[df['年份'] == '2019']
    
    # 准备地图数据
    map_data = []
    for _, row in data_2019.iterrows():
        map_data.append([row['省份'], row['旅游收入占GDP比例']])
    
    # 创建地图
    map_chart = Map()
    map_chart.set_global_opts(
        title_opts=opts.TitleOpts(title="2019年各省份旅游业收入占GDP比例"),
        visualmap_opts=opts.VisualMapOpts(min_=0, max_=0.5, is_piecewise=False)
    )
    map_chart.add("", map_data, maptype="china")
    
    # 保存地图
    output_path = os.path.join(output_folder, 'map_visualization.html')
    map_chart.render(output_path)
    print(f"地图可视化已保存到: {output_path}")

# 生成山西与全国对比图
def generate_shanxi_comparison_chart(df, years, output_folder):
    """生成山西与全国旅游收入占GDP比例的对比图"""
    # 计算全国平均值（只包含数值列）
    numeric_cols = ['旅游收入', 'GDP', '旅游收入占GDP比例']
    national_avg = df.groupby('年份')[numeric_cols].mean().reset_index()
    
    # 计算山西数据
    shanxi_data = df[df['省份'] == '山西省']
    
    # 创建图表
    fig, ax = plt.subplots(figsize=(15, 8))
    
    # 绘制比例折线
    ax.plot(national_avg['年份'], national_avg['旅游收入占GDP比例'], 'b-', label='全国平均比例')
    ax.plot(shanxi_data['年份'], shanxi_data['旅游收入占GDP比例'], 'r-', label='山西比例')
    
    # 添加标题和标签
    plt.title('山西与全国旅游收入占GDP比例对比')
    plt.xlabel('年份')
    plt.ylabel('旅游收入占GDP比例 (%)')
    plt.legend()
    
    # 保存图表
    output_path = os.path.join(output_folder, 'shanxi_comparison_chart.png')
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()
    print(f"山西与全国对比图已保存到: {output_path}")

# 主函数
def main():
    """主函数"""
    # 创建文件夹
    output_folder = create_correlation_folder()
    
    # 加载数据
    tourism_df, gdp_df = load_data()
    
    # 处理数据
    df, years, provinces = process_data(tourism_df, gdp_df)
    
    # 生成可视化
    generate_dual_axis_line_chart(df, years, output_folder)
    generate_scatter_chart(df, output_folder)
    # generate_heatmap(df, provinces, years, output_folder)  # 暂时注释掉，需要修复
    generate_map_visualization(df, output_folder)
    generate_shanxi_comparison_chart(df, years, output_folder)
    
    print("\n关联分析可视化完成！")

if __name__ == "__main__":
    main()
