import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LinearRegression
from statsmodels.tsa.seasonal import seasonal_decompose
from statsmodels.tsa.holtwinters import ExponentialSmoothing

# 尝试导入Prophet，如果不可用则设置标志
try:
    from prophet import Prophet
    PROPHET_AVAILABLE = True
except ImportError:
    PROPHET_AVAILABLE = False
    print("警告: Prophet库未安装，将无法使用Prophet进行预测。")
    print("请使用 pip install prophet 安装Prophet库。")

class SalesAnalyzer:
    """
    电商销售数据分析类，提供各种数据分析和可视化功能
    """
    def __init__(self, data_path=None):
        """
        初始化分析器，可选择性加载数据
        """
        self.df = None
        if data_path:
            self.load_data(data_path)
            
    def load_data(self, data_path):
        """
        加载销售数据
        
        Args:
            data_path: 数据文件路径，支持CSV或Excel格式
        """
        file_path = Path(data_path)
        if not file_path.exists():
            raise FileNotFoundError(f"文件不存在: {data_path}")
        
        # 根据文件扩展名选择加载方法
        if file_path.suffix.lower() == '.csv':
            self.df = pd.read_csv(file_path)
        elif file_path.suffix.lower() in ['.xlsx', '.xls']:
            self.df = pd.read_excel(file_path)
        else:
            raise ValueError(f"不支持的文件格式: {file_path.suffix}")
            
        # 将日期列转换为日期类型
        if '日期' in self.df.columns:
            self.df['日期'] = pd.to_datetime(self.df['日期'])
    
    def get_data_summary(self):
        """
        获取数据基本统计摘要
        """
        if self.df is None:
            return None
        
        summary = {
            '总记录数': len(self.df),
            '时间范围': [self.df['日期'].min().strftime('%Y-%m-%d'), 
                       self.df['日期'].max().strftime('%Y-%m-%d')],
            '商品类别数量': self.df['商品类别'].nunique(),
            '商品类别列表': self.df['商品类别'].unique().tolist(),
            '平均订单金额': round(self.df['总价'].mean(), 2),
            '最高订单金额': round(self.df['总价'].max(), 2),
            '顾客数量': self.df['顾客ID'].nunique(),
            '店铺数量': self.df['店铺'].nunique()
        }
        return summary
        
    def get_sales_by_time(self, time_unit='月', category=None):
        """
        按时间维度统计销售额
        
        Args:
            time_unit: 时间单位，可选 '日', '月', '季度', '年'
            category: 商品类别，可选，如果指定则只分析特定类别
        """
        if self.df is None:
            raise ValueError("数据未加载，请先加载数据")
            
        # 数据验证
        if len(self.df) == 0:
            raise ValueError("数据集为空")
            
        # 确保数据中有必要的列
        required_columns = {'日期', '总价', '商品类别'}
        missing_columns = required_columns - set(self.df.columns)
        if missing_columns:
            raise ValueError(f"数据中缺少必要的列: {missing_columns}")
            
        data = self.df.copy()
        
        # 确保日期列为datetime类型
        try:
            data['日期'] = pd.to_datetime(data['日期'])
        except Exception as e:
            raise ValueError(f"日期列转换失败: {str(e)}")
            
        # 添加年、月、季度列
        data['年'] = data['日期'].dt.year
        data['月'] = data['日期'].dt.month
        data['季度'] = (data['月'] - 1) // 3 + 1
        
        if category:
            if category not in data['商品类别'].unique():
                raise ValueError(f"指定的商品类别 '{category}' 不存在")
            data = data[data['商品类别'] == category]
            
        if len(data) == 0:
            raise ValueError("筛选后的数据集为空")
        
        try:
            if time_unit == '日':
                grouped = data.groupby('日期')['总价'].sum().reset_index()
                grouped['时间'] = grouped['日期'].dt.strftime('%Y-%m-%d')
            elif time_unit == '月':
                grouped = data.groupby([data['年'], data['月']])['总价'].sum().reset_index()
                grouped['时间'] = grouped.apply(lambda x: f"{int(x['年'])}-{int(x['月']):02d}", axis=1)
            elif time_unit == '季度':
                grouped = data.groupby([data['年'], data['季度']])['总价'].sum().reset_index()
                grouped['时间'] = grouped.apply(lambda x: f"{int(x['年'])}-Q{int(x['季度'])}", axis=1)
            elif time_unit == '年':
                grouped = data.groupby('年')['总价'].sum().reset_index()
                grouped['时间'] = grouped['年'].astype(int).astype(str)
            else:
                raise ValueError(f"不支持的时间单位: {time_unit}")
                
            if len(grouped) == 0:
                raise ValueError("分组后的数据集为空")
                
            return grouped
            
        except Exception as e:
            raise ValueError(f"数据处理过程中出错: {str(e)}")
    
    def get_sales_by_category(self, subcategory=False):
        """
        按商品类别统计销售额
        
        Args:
            subcategory: 是否包含子类别
        """
        if self.df is None:
            return None
        
        if subcategory:
            grouped = self.df.groupby(['商品类别', '子类别'])['总价'].sum().reset_index()
        else:
            grouped = self.df.groupby('商品类别')['总价'].sum().reset_index()
        
        return grouped
    
    def get_sales_by_region(self, region_level='省份'):
        """
        按地区统计销售额
        
        Args:
            region_level: 地区级别，可选 '省份' 或 '城市'
        """
        if self.df is None:
            return None
            
        if region_level not in ['省份', '城市']:
            raise ValueError(f"不支持的地区级别: {region_level}")
            
        grouped = self.df.groupby(region_level)['总价'].sum().reset_index()
        return grouped
    
    def get_top_products(self, n=10, measure='销售额', category=None):
        """
        获取热销产品排行
        
        Args:
            n: 返回前N个产品
            measure: 衡量标准，可选 '销售额' 或 '销售量'
            category: 可选，指定类别
        """
        if self.df is None:
            return None
            
        data = self.df.copy()
        if category:
            data = data[data['商品类别'] == category]
            
        if measure == '销售额':
            grouped = data.groupby('商品名称')['总价'].sum().reset_index()
            grouped = grouped.sort_values('总价', ascending=False).head(n)
            grouped = grouped.rename(columns={'总价': '销售额'})
        elif measure == '销售量':
            grouped = data.groupby('商品名称')['数量'].sum().reset_index()
            grouped = grouped.sort_values('数量', ascending=False).head(n)
            grouped = grouped.rename(columns={'数量': '销售量'})
        else:
            raise ValueError(f"不支持的衡量标准: {measure}")
            
        return grouped
    
    def get_customer_segments(self, n_clusters=4):
        """
        进行客户细分分析
        
        Args:
            n_clusters: 客户群体数量
        """
        if self.df is None:
            return None
            
        # 按客户ID聚合
        customer_data = self.df.groupby('顾客ID').agg({
            '订单ID': 'count',  # 订单数量
            '总价': 'sum',      # 总消费额
            '商品名称': 'nunique',  # 购买的不同商品数
            '日期': ['min', 'max']  # 首次和最近购买日期
        }).reset_index()
        
        # 重命名列
        customer_data.columns = ['顾客ID', '订单数量', '总消费额', '不同商品数', '首次购买日期', '最近购买日期']
        
        # 计算用户消费频率和最近一次购买距今天数
        last_date = self.df['日期'].max()
        customer_data['消费间隔天数'] = (last_date - customer_data['最近购买日期']).dt.days
        customer_data['活跃天数'] = (customer_data['最近购买日期'] - customer_data['首次购买日期']).dt.days
        customer_data['活跃天数'] = customer_data['活跃天数'].clip(lower=1)  # 避免除以零
        customer_data['平均每月消费次数'] = customer_data['订单数量'] / (customer_data['活跃天数'] / 30)
        
        # 选择聚类特征并进行标准化
        features = ['订单数量', '总消费额', '不同商品数', '消费间隔天数', '平均每月消费次数']
        X = customer_data[features]
        
        # 标准化数据
        scaler = StandardScaler()
        X_scaled = scaler.fit_transform(X)
        
        # K-means聚类
        kmeans = KMeans(n_clusters=n_clusters, random_state=42)
        clusters = kmeans.fit_predict(X_scaled)
        
        # 添加聚类结果
        customer_data['客户群体'] = clusters
        
        # 计算各群体特征均值
        cluster_features = customer_data.groupby('客户群体')[features].mean().reset_index()
        
        # 对客户群体进行标记
        # 基于总消费额和订单数量特征，为每个群体添加标签
        segment_labels = {}
        for cluster in range(n_clusters):
            row = cluster_features[cluster_features['客户群体'] == cluster].iloc[0]
            
            # 根据消费金额和频率判断客户类型
            if row['总消费额'] > cluster_features['总消费额'].median() and row['订单数量'] > cluster_features['订单数量'].median():
                label = "高价值忠诚客户"
            elif row['总消费额'] > cluster_features['总消费额'].median() and row['订单数量'] <= cluster_features['订单数量'].median():
                label = "高价值低频客户"
            elif row['总消费额'] <= cluster_features['总消费额'].median() and row['订单数量'] > cluster_features['订单数量'].median():
                label = "低价值高频客户"
            else:
                label = "低价值低频客户"
                
            # 考虑最近消费时间
            if row['消费间隔天数'] < cluster_features['消费间隔天数'].median():
                label = f"活跃{label}"
            else:
                label = f"流失风险{label}"
                
            segment_labels[cluster] = label
            
        # 添加客户群体标签
        customer_data['客户群体标签'] = customer_data['客户群体'].map(segment_labels)
        
        return customer_data, cluster_features
        
    def get_promotion_effect(self):
        """
        分析促销效果
        """
        if self.df is None:
            return None
            
        # 添加是否有折扣标记
        self.df['有折扣'] = self.df['折扣率'] < 1.0
        
        # 计算有折扣和无折扣的销售情况
        discount_effect = self.df.groupby(['有折扣']).agg({
            '订单ID': 'count',  # 订单数量
            '总价': 'sum',      # 总销售额
            '数量': 'sum'       # 总销售量
        }).reset_index()
        
        # 按商品类别计算折扣效果
        category_discount = self.df.groupby(['商品类别', '有折扣']).agg({
            '订单ID': 'count',  # 订单数量
            '总价': 'sum',      # 总销售额
        }).reset_index()
        
        return discount_effect, category_discount
    
    def get_seasonal_trends(self):
        """
        分析季节性趋势
        """
        if self.df is None:
            return None
            
        # 按月份统计销售额
        monthly_sales = self.df.groupby(['年', '月'])['总价'].sum().reset_index()
        monthly_sales['月份'] = monthly_sales['月'].astype(int)
        
        # 按季度统计销售额
        quarterly_sales = self.df.groupby(['年', '季度'])['总价'].sum().reset_index()
        quarterly_sales['季度'] = quarterly_sales['季度'].astype(int)
        
        # 计算特殊购物节的销售情况
        
        # 春节 (假设2022年春节在2月1日，2023年春节在1月22日)
        spring_festival_2022 = (self.df['年'] == 2022) & (self.df['月'] == 2) & (self.df['日'] <= 7)
        spring_festival_2023 = (self.df['年'] == 2023) & (self.df['月'] == 1) & (self.df['日'] >= 20) & (self.df['日'] <= 27)
        
        # 618购物节
        festival_618 = (self.df['月'] == 6) & (self.df['日'] >= 10) & (self.df['日'] <= 20)
        
        # 双11购物节
        festival_1111 = (self.df['月'] == 11) & (self.df['日'] >= 9) & (self.df['日'] <= 12)
        
        # 创建标记特殊日期的列
        self.df['购物节'] = '普通日期'
        self.df.loc[spring_festival_2022 | spring_festival_2023, '购物节'] = '春节'
        self.df.loc[festival_618, '购物节'] = '618购物节'
        self.df.loc[festival_1111, '购物节'] = '双11购物节'
        
        # 特殊购物节的销售统计
        festival_sales = self.df.groupby(['购物节']).agg({
            '订单ID': 'count',
            '总价': 'sum',
            '折扣率': 'mean'
        }).reset_index()
        
        return monthly_sales, quarterly_sales, festival_sales

    def create_sales_trend_plot(self, time_unit='月'):
        """
        创建销售趋势图
        """
        if self.df is None:
            return None
            
        sales_data = self.get_sales_by_time(time_unit)
        
        plt.figure(figsize=(12, 6))
        plt.plot(sales_data['时间'], sales_data['总价'], marker='o', linewidth=2)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.title(f'按{time_unit}销售趋势', fontsize=14)
        plt.xlabel('时间', fontsize=12)
        plt.ylabel('销售额（元）', fontsize=12)
        
        # 如果时间点较多，旋转标签
        if len(sales_data) > 10:
            plt.xticks(rotation=45)
            
        plt.tight_layout()
        
        return plt.gcf()
        
    def create_category_sales_plot(self):
        """
        创建类别销售额对比图
        """
        if self.df is None:
            return None
            
        category_sales = self.get_sales_by_category()
        
        plt.figure(figsize=(10, 8))
        bars = plt.barh(category_sales['商品类别'], category_sales['总价'])
        plt.grid(True, linestyle='--', alpha=0.7, axis='x')
        plt.title('各商品类别销售额', fontsize=14)
        plt.xlabel('销售额（元）', fontsize=12)
        plt.ylabel('商品类别', fontsize=12)
        
        # 添加数据标签
        for bar in bars:
            width = bar.get_width()
            plt.text(width + width*0.02, bar.get_y() + bar.get_height()/2, 
                    f'{width:,.0f}', ha='left', va='center')
        
        plt.tight_layout()
        
        return plt.gcf()
        
    def create_region_sales_plot(self, region_level='省份'):
        """
        创建区域销售对比图
        """
        if self.df is None:
            return None
            
        region_sales = self.get_sales_by_region(region_level)
        
        plt.figure(figsize=(12, 8))
        sns.barplot(x='总价', y=region_level, data=region_sales.sort_values('总价', ascending=False))
        plt.grid(True, linestyle='--', alpha=0.7, axis='x')
        plt.title(f'各{region_level}销售额对比', fontsize=14)
        plt.xlabel('销售额（元）', fontsize=12)
        plt.ylabel(region_level, fontsize=12)
        plt.tight_layout()
        
        return plt.gcf()
        
    def create_festival_impact_plot(self):
        """
        创建购物节影响分析图
        """
        if self.df is None:
            return None
            
        _, _, festival_sales = self.get_seasonal_trends()
        
        plt.figure(figsize=(10, 6))
        sns.barplot(x='购物节', y='总价', data=festival_sales)
        plt.grid(True, linestyle='--', alpha=0.7, axis='y')
        plt.title('各购物节销售额对比', fontsize=14)
        plt.xlabel('购物节', fontsize=12)
        plt.ylabel('销售额（元）', fontsize=12)
        
        # 添加折扣率标签
        for i, row in enumerate(festival_sales.itertuples()):
            plt.text(i, row.总价 * 0.5, f'平均折扣: {row.折扣率:.2f}', 
                    ha='center', va='center', color='white', fontweight='bold')
        
        plt.tight_layout()
        
        return plt.gcf()
        
    def create_heatmap(self, category=None):
        """
        创建热图，展示月度和类别的销售关系
        """
        if self.df is None:
            return None
            
        data = self.df.copy()
        if category:
            data = data[data['商品类别'] == category]
            
        # 创建月份-类别交叉表
        pivot = pd.pivot_table(
            data, 
            values='总价', 
            index='商品类别', 
            columns='月',
            aggfunc='sum'
        ).fillna(0)
        
        plt.figure(figsize=(14, 8))
        sns.heatmap(pivot, annot=True, fmt='.0f', cmap='YlGnBu', linewidths=.5)
        plt.title('月份-商品类别销售热图', fontsize=14)
        plt.xlabel('月份', fontsize=12)
        plt.ylabel('商品类别', fontsize=12)
        plt.tight_layout()
        
        return plt.gcf()
        
    def predict_sales(self, time_unit='月', category=None, periods=6, method='prophet'):
        """
        预测未来销售趋势
        
        Args:
            time_unit: 时间单位，可选 '日', '月', '季度'
            category: 商品类别，可选，如果指定则只分析特定类别
            periods: 预测未来的时间段数量
            method: 预测方法，可选 'prophet', 'exponential_smoothing', 'linear'
            
        Returns:
            包含历史和预测数据的DataFrame
        """
        if self.df is None:
            raise ValueError("数据未加载，请先加载数据")
            
        # 获取销售数据
        sales_data = self.get_sales_by_time(time_unit, category)
        if sales_data is None or len(sales_data) == 0:
            raise ValueError("没有足够的数据进行预测")
        
        # 根据时间单位设置预测参数
        freq_map = {
            '日': 'D', 
            '月': 'M',
            '季度': 'Q'
        }
        
        if time_unit not in freq_map:
            raise ValueError(f"不支持的时间单位: {time_unit}，仅支持日、月、季度")
        
        freq = freq_map[time_unit]
        
        # 根据预测方法执行不同的预测
        if method == 'prophet':
            if not PROPHET_AVAILABLE:
                raise ValueError("Prophet库未安装，无法使用Prophet进行预测。请使用 'exponential_smoothing' 或 'linear' 方法。")
            return self._predict_with_prophet(sales_data, periods, freq)
        elif method == 'exponential_smoothing':
            return self._predict_with_exp_smoothing(sales_data, periods)
        elif method == 'linear':
            return self._predict_with_linear_regression(sales_data, periods)
        else:
            raise ValueError(f"不支持的预测方法: {method}")
    
    def _predict_with_prophet(self, sales_data, periods, freq):
        """使用Prophet模型进行预测"""
        # 准备Prophet所需的数据格式
        df_prophet = sales_data[['时间', '总价']].rename(columns={'时间': 'ds', '总价': 'y'})
        
        # 将时间列转换为日期类型
        if freq == 'M':
            df_prophet['ds'] = pd.to_datetime(df_prophet['ds'] + '-01')
        elif freq == 'Q':
            # 将 "YYYY-QN" 转换为季度末日期
            def convert_quarter(q_str):
                year, quarter = q_str.split('-Q')
                month = int(quarter) * 3
                return f"{year}-{month:02d}-01"
            
            df_prophet['ds'] = pd.to_datetime(df_prophet['ds'].apply(convert_quarter))
        else:
            df_prophet['ds'] = pd.to_datetime(df_prophet['ds'])
        
        # 训练Prophet模型
        model = Prophet(yearly_seasonality=True, weekly_seasonality=False, daily_seasonality=False)
        if freq == 'D':
            model.add_seasonality(name='weekly', period=7, fourier_order=3)
        
        model.fit(df_prophet)
        
        # 创建未来日期的数据框
        future = model.make_future_dataframe(periods=periods, freq=freq)
        
        # 预测
        forecast = model.predict(future)
        
        # 将预测结果与历史数据合并
        result = forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].rename(
            columns={
                'ds': '日期',
                'yhat': '预测销售额',
                'yhat_lower': '预测下限',
                'yhat_upper': '预测上限'
            }
        )
        
        # 添加标记，区分历史数据和预测数据
        result['数据类型'] = '预测'
        result.loc[:len(df_prophet)-1, '数据类型'] = '历史'
        
        # 添加实际数据
        result.loc[:len(df_prophet)-1, '实际销售额'] = df_prophet['y'].values
        
        return result
    
    def _predict_with_exp_smoothing(self, sales_data, periods):
        """使用指数平滑法进行预测"""
        # 创建时间序列
        y = sales_data['总价'].values
        
        # 训练模型（三次指数平滑，也称为Holt-Winters方法）
        model = ExponentialSmoothing(
            y,
            seasonal_periods=12,  # 假设数据有年度季节性
            trend='add',
            seasonal='add',
        )
        
        fit = model.fit(optimized=True)
        
        # 预测
        forecast = fit.forecast(periods)
        
        # 创建结果数据框
        last_date = pd.to_datetime(sales_data['时间'].iloc[-1])
        date_range = pd.date_range(start=last_date, periods=periods+1, freq='M')[1:]
        
        # 创建包含历史和预测的结果
        result = pd.DataFrame({
            '日期': pd.concat([pd.Series(sales_data['时间']), pd.Series(date_range.strftime('%Y-%m'))]),
            '预测销售额': np.concatenate([fit.fittedvalues, forecast]),
            '数据类型': ['历史'] * len(y) + ['预测'] * periods
        })
        
        # 添加实际数据
        result.loc[:len(y)-1, '实际销售额'] = y
        
        return result
    
    def _predict_with_linear_regression(self, sales_data, periods):
        """使用线性回归进行预测"""
        # 创建特征
        X = np.array(range(len(sales_data))).reshape(-1, 1)
        y = sales_data['总价'].values
        
        # 训练模型
        model = LinearRegression()
        model.fit(X, y)
        
        # 预测
        X_future = np.array(range(len(sales_data), len(sales_data) + periods)).reshape(-1, 1)
        forecast = model.predict(X_future)
        
        # 创建结果数据框
        last_date = pd.to_datetime(sales_data['时间'].iloc[-1])
        date_range = pd.date_range(start=last_date, periods=periods+1, freq='M')[1:]
        
        # 创建包含历史和预测的结果
        result = pd.DataFrame({
            '日期': pd.concat([pd.Series(sales_data['时间']), pd.Series(date_range.strftime('%Y-%m'))]),
            '预测销售额': np.concatenate([model.predict(X), forecast]),
            '数据类型': ['历史'] * len(y) + ['预测'] * periods
        })
        
        # 添加实际数据
        result.loc[:len(y)-1, '实际销售额'] = y
        
        return result
    
    def generate_decision_suggestions(self, category=None):
        """
        根据销售数据生成决策建议
        
        Args:
            category: 商品类别，可选，如果指定则只分析特定类别
            
        Returns:
            决策建议列表
        """
        if self.df is None:
            raise ValueError("数据未加载，请先加载数据")
        
        suggestions = []
        
        try:
            # 1. 销售趋势建议
            recent_trend = self._analyze_recent_trend(category)
            suggestions.append({
                "类型": "销售趋势建议",
                "建议": recent_trend
            })
            
            # 2. 库存管理建议
            inventory_suggestion = self._generate_inventory_suggestions(category)
            suggestions.append({
                "类型": "库存管理建议",
                "建议": inventory_suggestion
            })
            
            # 3. 促销策略建议
            promotion_suggestion = self._generate_promotion_suggestions(category)
            suggestions.append({
                "类型": "促销策略建议",
                "建议": promotion_suggestion
            })
            
            # 4. 客户营销建议
            customer_suggestion = self._generate_customer_suggestions()
            suggestions.append({
                "类型": "客户营销建议",
                "建议": customer_suggestion
            })
            
            # 5. 产品组合建议
            product_suggestion = self._generate_product_suggestions(category)
            suggestions.append({
                "类型": "产品组合建议",
                "建议": product_suggestion
            })
            
        except Exception as e:
            suggestions.append({
                "类型": "错误",
                "建议": f"生成建议时发生错误: {str(e)}"
            })
            
        return suggestions
    
    def _analyze_recent_trend(self, category=None):
        """分析最近的销售趋势并给出建议"""
        try:
            # 获取月度销售数据
            sales_data = self.get_sales_by_time('月', category)
            if sales_data is None or len(sales_data) < 3:
                return "数据不足，无法分析最近趋势"
            
            # 获取最近3个月的数据
            recent_data = sales_data.iloc[-3:].copy()
            
            # 计算增长率
            recent_data['增长率'] = recent_data['总价'].pct_change()
            
            # 分析最近的趋势
            last_growth = recent_data['增长率'].iloc[-1] if len(recent_data) > 1 else 0
            
            if last_growth > 0.1:
                return f"销售呈显著上升趋势（增长率{last_growth:.1%}），建议增加库存并加大营销力度以满足增长需求"
            elif last_growth > 0:
                return f"销售呈小幅上升趋势（增长率{last_growth:.1%}），建议维持当前策略，关注客户反馈"
            elif last_growth > -0.05:
                return f"销售基本持平略有下降（变化率{last_growth:.1%}），建议开展促销活动提升销量"
            else:
                return f"销售呈下降趋势（降低率{-last_growth:.1%}），建议深入分析原因，调整产品策略或定价"
                
        except Exception:
            return "无法分析销售趋势"
    
    def _generate_inventory_suggestions(self, category=None):
        """生成库存管理建议"""
        try:
            # 获取类别数据
            if category:
                category_data = self.df[self.df['商品类别'] == category]
            else:
                category_data = self.df
                
            # 获取热销商品
            top_products = self.get_top_products(n=5, measure='销售量', category=category)
            
            # 检查是否有折扣对销量的影响
            has_discount = '折扣率' in self.df.columns
            
            if has_discount:
                # 分析折扣对销量的影响
                discount_effect = category_data.groupby(['商品名称', '折扣率']).agg({'数量': 'sum'}).reset_index()
                
                # 根据分析结果生成建议
                if not top_products.empty:
                    suggestions = [
                        f"热销商品 '{top_products.iloc[0]['商品名称']}' 应保持充足库存，建议采用安全库存管理，设置自动补货点",
                        "定期监控库存水平，对销量预测较高的商品提前备货",
                        "对季节性商品实施季节性库存策略，在销售旺季前增加库存"
                    ]
                    
                    return "；".join(suggestions)
                else:
                    return "数据不足，无法提供具体库存建议"
            else:
                if not top_products.empty:
                    return f"热销商品 '{top_products.iloc[0]['商品名称']}' 应保持充足库存；建议实施数据驱动的库存管理系统"
                else:
                    return "建议实施数据驱动的库存管理系统，根据历史销售数据进行动态调整"
                    
        except Exception:
            return "无法生成库存管理建议"
    
    def _generate_promotion_suggestions(self, category=None):
        """生成促销策略建议"""
        try:
            # 检查是否有折扣率数据
            has_discount = '折扣率' in self.df.columns
            
            if not has_discount:
                return "数据中缺少折扣信息，无法提供促销策略建议"
                
            # 分析折扣效果
            discount_effect, category_discount = self.get_promotion_effect()
            
            # 分析不同折扣力度的效果
            if '折扣率' in self.df.columns:
                # 创建折扣区间
                self.df['折扣区间'] = pd.cut(
                    self.df['折扣率'], 
                    bins=[0, 0.7, 0.8, 0.9, 1.0],
                    labels=['大幅折扣(>30%)', '中度折扣(20-30%)', '小幅折扣(10-20%)', '无折扣/微折扣(<10%)']
                )
                
                # 分析各折扣区间的销售情况
                discount_analysis = self.df.groupby('折扣区间').agg({
                    '数量': 'sum',
                    '总价': 'sum',
                    '订单ID': 'count'
                }).reset_index()
                
                # 计算平均订单金额
                discount_analysis['平均订单金额'] = discount_analysis['总价'] / discount_analysis['订单ID']
                
                # 找出最有效的折扣区间
                if not discount_analysis.empty:
                    best_discount = discount_analysis.sort_values('数量', ascending=False).iloc[0]
                    best_revenue = discount_analysis.sort_values('总价', ascending=False).iloc[0]
                    
                    suggestions = []
                    
                    if best_discount['折扣区间'] == best_revenue['折扣区间']:
                        suggestions.append(f"'{best_discount['折扣区间']}'策略在销量和收入上都表现最佳，建议主要采用此折扣策略")
                    else:
                        suggestions.append(f"'{best_discount['折扣区间']}'策略在销量上表现最佳")
                        suggestions.append(f"'{best_revenue['折扣区间']}'策略在总收入上表现最佳")
                        suggestions.append(f"建议根据销售目标灵活采用不同折扣策略：提升销量使用'{best_discount['折扣区间']}'，提升收入使用'{best_revenue['折扣区间']}'")
                    
                    # 季节性建议
                    _, _, festival_sales = self.get_seasonal_trends()
                    if not festival_sales.empty:
                        best_festival = festival_sales.sort_values('总价', ascending=False).iloc[0]
                        suggestions.append(f"'{best_festival['购物节']}'期间销售表现最佳，建议在该时期加大促销力度")
                    
                    return "；".join(suggestions)
                
            return "建议在热销商品上采取小幅折扣策略，在滞销商品上采取大幅折扣策略"
                
        except Exception:
            return "无法生成促销策略建议"
    
    def _generate_customer_suggestions(self):
        """生成客户营销建议"""
        try:
            # 获取客户细分数据
            customer_data, _ = self.get_customer_segments()
            
            if customer_data is None or customer_data.empty:
                return "数据不足，无法提供客户营销建议"
                
            # 分析各客户群体
            segment_counts = customer_data.groupby('客户群体标签').size().reset_index(name='数量')
            segment_values = customer_data.groupby('客户群体标签')['总消费额'].mean().reset_index()
            
            # 合并信息
            segments = segment_counts.merge(segment_values, on='客户群体标签')
            
            if segments.empty:
                return "无法识别明确的客户群体，建议加强数据收集"
                
            # 生成针对各群体的建议
            suggestions = []
            
            for _, segment in segments.iterrows():
                label = segment['客户群体标签']
                
                if '高价值' in label and '活跃' in label:
                    suggestions.append(f"针对'{label}'群体({segment['数量']}人)：提供VIP服务和专属优惠，增强客户忠诚度")
                elif '高价值' in label and '流失风险' in label:
                    suggestions.append(f"针对'{label}'群体({segment['数量']}人)：实施挽留计划，提供个性化优惠和回馈")
                elif '低价值' in label and '活跃' in label:
                    suggestions.append(f"针对'{label}'群体({segment['数量']}人)：实施上销和交叉销售策略，提高客单价")
                elif '低价值' in label and '流失风险' in label:
                    suggestions.append(f"针对'{label}'群体({segment['数量']}人)：评估营销成本，考虑低成本的自动化营销或放弃")
            
            # 通用建议
            suggestions.append("建立客户生命周期管理系统，针对不同阶段客户采取差异化营销策略")
            
            return "；".join(suggestions)
                
        except Exception:
            return "无法生成客户营销建议"
    
    def _generate_product_suggestions(self, category=None):
        """生成产品组合建议"""
        try:
            # 获取类别销售数据
            category_sales = self.get_sales_by_category()
            
            if category_sales is None or category_sales.empty:
                return "数据不足，无法提供产品组合建议"
                
            # 排序类别销售数据
            category_sales = category_sales.sort_values('总价', ascending=False)
            
            # 计算类别贡献占比
            total_sales = category_sales['总价'].sum()
            category_sales['占比'] = category_sales['总价'] / total_sales
            
            # 找出销售占比最高和最低的类别
            top_category = category_sales.iloc[0]
            bottom_categories = category_sales.iloc[-3:] if len(category_sales) >= 3 else category_sales.iloc[-1:]
            
            suggestions = []
            
            # 针对热销类别的建议
            suggestions.append(f"重点发展'{top_category['商品类别']}'类别，其销售占比达{top_category['占比']:.1%}，可考虑扩充产品线或提高利润率")
            
            # 针对滞销类别的建议
            bottom_names = "、".join(bottom_categories['商品类别'].tolist())
            suggestions.append(f"评估'{bottom_names}'等类别的产品策略，考虑产品创新、调整定价或淘汰部分产品")
            
            # 产品组合多样化建议
            if top_category['占比'] > 0.4:
                suggestions.append(f"当前产品结构过于依赖'{top_category['商品类别']}'类别，建议适当多元化产品组合，分散风险")
            else:
                suggestions.append("当前产品结构相对均衡，建议保持多元化策略，并根据市场反馈适时调整")
            
            # 获取热销产品
            top_products = self.get_top_products(n=3, measure='销售额', category=category)
            if top_products is not None and not top_products.empty:
                top_names = "、".join(top_products['商品名称'].iloc[:3].tolist())
                suggestions.append(f"热销产品'{top_names}'表现突出，建议加强供应链管理，确保库存充足")
            
            return "；".join(suggestions)
                
        except Exception:
            return "无法生成产品组合建议" 