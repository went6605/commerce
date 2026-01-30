import pandas as pd
import numpy as np
import datetime
import random
from faker import Faker
from pathlib import Path

# 设置随机种子以确保可重复性
np.random.seed(42)
random.seed(42)
fake = Faker(['zh_CN'])
Faker.seed(42)

# 设置参数
num_records = 16985  # 总记录数
start_date = datetime.datetime(2022, 1, 1)
end_date = datetime.datetime(2023, 12, 31)
date_range = (end_date - start_date).days

# 商品类别
categories = ['电子产品', '服装', '家居', '食品', '美妆', '图书', '运动', '母婴', '玩具', '珠宝']

# 电子产品子类别和产品
electronics_subcategories = ['手机', '电脑', '相机', '音响', '电视', '耳机']
electronics_products = {
    '手机': ['iPhone 13', 'iPhone 14', 'iPhone 15', '华为P50', '华为P60', '小米13', '小米14', 'OPPO Find X5', '三星S22', 'vivo X90'],
    '电脑': ['MacBook Pro', 'MacBook Air', '联想ThinkPad', '华为MateBook', '惠普Spectre', '戴尔XPS', '华硕ZenBook', '微软Surface'],
    '相机': ['佳能EOS R5', '尼康Z9', '索尼A7IV', '富士X-T4', '松下GH6'],
    '音响': ['Bose SoundLink', 'JBL Flip 6', '索尼WH-1000XM5', '哈曼卡顿Aura Studio 3'],
    '电视': ['小米电视6', '三星Neo QLED', 'LG OLED C2', '索尼A80J', 'TCL Mini LED'],
    '耳机': ['AirPods Pro', 'AirPods Max', '索尼WF-1000XM4', '华为FreeBuds Pro 2', '小米Buds 4 Pro']
}

# 服装子类别和产品
clothing_subcategories = ['上衣', '裤子', '鞋子', '连衣裙', '外套', '内衣']
clothing_products = {
    '上衣': ['纯棉T恤', '亚麻衬衫', '针织衫', '卫衣', '夏季短袖'],
    '裤子': ['牛仔裤', '休闲裤', '运动裤', '短裤', '阔腿裤'],
    '鞋子': ['运动鞋', '休闲鞋', '皮鞋', '高跟鞋', '凉鞋', '靴子'],
    '连衣裙': ['雪纺连衣裙', '针织连衣裙', '吊带裙', '夏季碎花裙', '长裙'],
    '外套': ['羽绒服', '夹克', '风衣', '西装', '毛衣'],
    '内衣': ['文胸', '内裤', '睡衣', '袜子', '保暖内衣']
}

# 家居子类别和产品
home_subcategories = ['家具', '厨具', '床上用品', '装饰品', '灯具']
home_products = {
    '家具': ['沙发', '餐桌', '床', '衣柜', '书架', '电视柜'],
    '厨具': ['锅具套装', '刀具', '厨房电器', '餐具', '保温杯'],
    '床上用品': ['四件套', '被子', '枕头', '床垫', '毛毯'],
    '装饰品': ['挂画', '花瓶', '摆件', '地毯', '抱枕'],
    '灯具': ['吊灯', '台灯', '落地灯', '壁灯', '智能灯']
}

# 生成商店列表
stores = [f"{fake.company()}电商店铺" for _ in range(50)]

# 生成省份和城市
provinces = ['北京市', '上海市', '广东省', '江苏省', '浙江省', '四川省', '河南省', '山东省', '湖北省', '湖南省']
cities = {
    '北京市': ['朝阳区', '海淀区', '东城区', '西城区', '丰台区'],
    '上海市': ['浦东新区', '静安区', '徐汇区', '黄浦区', '虹口区'],
    '广东省': ['广州市', '深圳市', '佛山市', '东莞市', '珠海市'],
    '江苏省': ['南京市', '苏州市', '无锡市', '常州市', '扬州市'],
    '浙江省': ['杭州市', '宁波市', '温州市', '嘉兴市', '湖州市'],
    '四川省': ['成都市', '绵阳市', '德阳市', '南充市', '宜宾市'],
    '河南省': ['郑州市', '洛阳市', '开封市', '新乡市', '许昌市'],
    '山东省': ['济南市', '青岛市', '烟台市', '潍坊市', '临沂市'],
    '湖北省': ['武汉市', '宜昌市', '襄阳市', '荆州市', '黄石市'],
    '湖南省': ['长沙市', '株洲市', '湘潭市', '衡阳市', '岳阳市']
}

# 生成支付方式
payment_methods = ['支付宝', '微信支付', '银联', '信用卡', '货到付款']

# 准备数据列表
data = []

for i in range(num_records):
    # 生成日期，确保日期分布不均匀，模拟季节性和节假日效应
    random_day = random.randint(0, date_range)
    date = start_date + datetime.timedelta(days=random_day)
    
    # 在重要购物节添加更多订单
    is_special_day = False
    # 春节 (假设2022年春节在2月1日，2023年春节在1月22日)
    if (date.month == 2 and date.day <= 7 and date.year == 2022) or \
       (date.month == 1 and date.day >= 20 and date.day <= 27 and date.year == 2023):
        if random.random() < 0.8:  # 80%的概率在春节期间
            is_special_day = True
    # 618购物节
    elif (date.month == 6 and date.day >= 10 and date.day <= 20):
        if random.random() < 0.7:  # 70%的概率在618期间
            is_special_day = True
    # 双11购物节
    elif (date.month == 11 and date.day >= 9 and date.day <= 12):
        if random.random() < 0.9:  # 90%的概率在双11期间
            is_special_day = True
    
    # 如果是特殊日期但未被选中，则随机选择其他日期
    if not is_special_day and random_day % 20 == 0:  # 每20天的样本重新选日期
        random_day = random.randint(0, date_range)
        date = start_date + datetime.timedelta(days=random_day)
    
    # 生成订单信息
    order_id = f"ORD-{date.year}{date.month:02d}{date.day:02d}-{i:05d}"
    
    # 选择类别
    category = random.choice(categories)
    
    # 根据类别选择子类别和产品
    if category == '电子产品':
        subcategory = random.choice(electronics_subcategories)
        product_name = random.choice(electronics_products[subcategory])
        # 电子产品价格范围较大
        price = round(random.uniform(500, 10000), 2)
    elif category == '服装':
        subcategory = random.choice(clothing_subcategories)
        product_name = random.choice(clothing_products[subcategory])
        # 服装价格中等
        price = round(random.uniform(50, 2000), 2)
    elif category == '家居':
        subcategory = random.choice(home_subcategories)
        product_name = random.choice(home_products[subcategory])
        # 家居产品价格区间广
        price = round(random.uniform(30, 5000), 2)
    else:
        subcategory = "常规"
        product_name = f"{category}商品{random.randint(1, 100)}"
        # 其他类别产品价格
        price = round(random.uniform(20, 2000), 2)
    
    # 数量，大多数订单为小数量
    quantity = np.random.geometric(p=0.5)
    
    # 随机生成折扣
    discount_rate = 1.0  # 默认无折扣
    # 特殊日期有更大折扣
    if is_special_day:
        discount_rate = round(random.uniform(0.5, 0.9), 2)  # 特殊日期折扣50%-90%
    elif random.random() < 0.3:  # 30%的普通订单有折扣
        discount_rate = round(random.uniform(0.7, 0.95), 2)  # 普通折扣70%-95%
    
    # 计算总价
    total_price = round(price * quantity * discount_rate, 2)
    
    # 生成顾客信息
    customer_id = f"CUST-{random.randint(10000, 99999)}"
    customer_name = fake.name()
    
    # 生成位置信息
    province = random.choice(provinces)
    city = random.choice(cities[province])
    address = fake.address()
    
    # 生成店铺信息
    store = random.choice(stores)
    
    # 生成评分
    # 高价产品和特殊日期购买的产品评分稍微高一些
    if price > 1000 or is_special_day:
        rating = max(1, min(5, round(random.normalvariate(4.2, 0.7))))
    else:
        rating = max(1, min(5, round(random.normalvariate(3.8, 0.9))))
    
    # 生成支付信息
    payment_method = random.choice(payment_methods)
    
    # 配送时间（天）
    delivery_time = random.randint(1, 10)
    
    # 添加到数据列表
    data.append({
        '订单ID': order_id,
        '日期': date.strftime('%Y-%m-%d'),
        '年': date.year,
        '月': date.month,
        '日': date.day,
        '季度': (date.month - 1) // 3 + 1,
        '商品类别': category,
        '子类别': subcategory,
        '商品名称': product_name,
        '单价': price,
        '数量': quantity,
        '折扣率': discount_rate,
        '总价': total_price,
        '顾客ID': customer_id,
        '顾客姓名': customer_name,
        '省份': province,
        '城市': city,
        '地址': address,
        '店铺': store,
        '评分': rating,
        '支付方式': payment_method,
        '配送时间(天)': delivery_time
    })

# 创建DataFrame
df = pd.DataFrame(data)

# 保存为CSV和Excel
output_dir = Path('data')
output_dir.mkdir(exist_ok=True)

df.to_csv(output_dir / 'ecommerce_sales_data.csv', index=False, encoding='utf-8-sig')
df.to_excel(output_dir / 'ecommerce_sales_data.xlsx', index=False)

print(f"成功生成{len(df)}条电商销售数据记录，已保存到data目录下")

# 打印数据统计信息
print("\n数据统计信息:")
print(f"时间范围: {df['日期'].min()} 至 {df['日期'].max()}")
print(f"商品类别数量: {df['商品类别'].nunique()}")
print(f"平均订单价格: {df['总价'].mean():.2f}")
print(f"最高订单价格: {df['总价'].max():.2f}")
print(f"顾客数量: {df['顾客ID'].nunique()}")
print(f"店铺数量: {df['店铺'].nunique()}") 