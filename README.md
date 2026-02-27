# 旅游业GDP数据分析与可视化项目

## 项目结构

```
tourism-gdp-dataAnalysis/
├── data/             # 原始数据文件夹
├── data_final/       # 处理后的清洗数据
├── visualization/    # 可视化结果
│   └── [文件名]/     # 每个文件的可视化结果
├── main.py           # 主脚本
├── check_data.py     # 数据检查脚本
├── test_libraries.py # 库检查脚本
└── README.md         # 项目说明
```

## 功能介绍

1. **数据清洗**：使用pandas对原始数据进行清洗和处理
2. **数据可视化**：
   - 柱线图
   - 条形图
   - 饼图
   - 数据地图（需要folium库）

## 依赖库

- pandas
- matplotlib
- seaborn (可选)
- folium (可选，用于数据地图)
- xlrd (用于读取Excel文件)

## 使用方法

1. 将原始数据放入`data`文件夹
2. 运行主脚本：
   ```bash
   python main.py
   ```
3. 处理后的数据将保存在`data_final`文件夹
4. 可视化结果将保存在`visualization`文件夹中对应的子文件夹

## 注意事项

- 由于网络代理问题，可能无法安装某些依赖库
- 对于Excel文件，需要安装xlrd库
- 对于数据地图，需要安装folium库

## 项目特点

- 代码结构清晰，注释详细
- 功能模块化，易于扩展
- 支持多种数据可视化方式
- 自动创建必要的文件夹结构