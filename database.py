import sqlite3
from datetime import datetime

DATABASE = 'stock.db'


def get_db():
    """获取数据库连接"""
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """初始化数据库表"""
    conn = get_db()
    cursor = conn.cursor()

    # 库存表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT NOT NULL,
            product_code TEXT NOT NULL,
            size TEXT NOT NULL,
            purchase_price REAL NOT NULL,
            quantity INTEGER NOT NULL DEFAULT 0,
            UNIQUE(product_code, size, purchase_price)
        )
    ''')

    # 入库记录表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS stock_in (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT NOT NULL,
            product_code TEXT NOT NULL,
            size TEXT NOT NULL,
            purchase_price REAL NOT NULL,
            quantity INTEGER NOT NULL,
            created_at DATETIME NOT NULL
        )
    ''')

    # 出库记录表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS stock_out (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT NOT NULL,
            product_code TEXT NOT NULL,
            size TEXT NOT NULL,
            purchase_price REAL NOT NULL,
            sell_price REAL NOT NULL,
            quantity INTEGER NOT NULL,
            profit REAL NOT NULL,
            created_at DATETIME NOT NULL
        )
    ''')

    conn.commit()
    conn.close()


def add_stock(category, product_code, size, purchase_price, quantity):
    """入库操作"""
    conn = get_db()
    cursor = conn.cursor()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # 记录入库
    cursor.execute('''
        INSERT INTO stock_in (category, product_code, size, purchase_price, quantity, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', (category, product_code, size, purchase_price, quantity, now))

    # 更新库存（如果存在则增加数量，否则插入新记录）
    cursor.execute('''
        INSERT INTO inventory (category, product_code, size, purchase_price, quantity)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(product_code, size, purchase_price)
        DO UPDATE SET quantity = quantity + ?, category = ?
    ''', (category, product_code, size, purchase_price, quantity, quantity, category))

    conn.commit()
    conn.close()


def get_inventory(category=None):
    """获取库存列表"""
    conn = get_db()
    cursor = conn.cursor()

    if category and category != 'all':
        cursor.execute('''
            SELECT * FROM inventory WHERE quantity > 0 AND category = ?
            ORDER BY product_code, size
        ''', (category,))
    else:
        cursor.execute('''
            SELECT * FROM inventory WHERE quantity > 0
            ORDER BY product_code, size
        ''')

    rows = cursor.fetchall()
    conn.close()
    return rows


def get_total_value():
    """计算总货值"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT SUM(purchase_price * quantity) as total FROM inventory WHERE quantity > 0')
    result = cursor.fetchone()
    conn.close()
    return result['total'] if result['total'] else 0


def search_by_product_code(product_code):
    """根据货号搜索库存"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT * FROM inventory
        WHERE product_code LIKE ? AND quantity > 0
        ORDER BY size
    ''', (f'%{product_code}%',))
    rows = cursor.fetchall()
    conn.close()
    return rows


def get_inventory_item(item_id):
    """获取单个库存项"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM inventory WHERE id = ?', (item_id,))
    row = cursor.fetchone()
    conn.close()
    return row


def remove_stock(item_id, sell_price, quantity):
    """出库操作"""
    conn = get_db()
    cursor = conn.cursor()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # 获取库存信息
    cursor.execute('SELECT * FROM inventory WHERE id = ?', (item_id,))
    item = cursor.fetchone()

    if not item or item['quantity'] < quantity:
        conn.close()
        return False, '库存不足'

    # 计算利润
    profit = (sell_price - item['purchase_price']) * quantity

    # 记录出库
    cursor.execute('''
        INSERT INTO stock_out (category, product_code, size, purchase_price, sell_price, quantity, profit, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (item['category'], item['product_code'], item['size'],
          item['purchase_price'], sell_price, quantity, profit, now))

    # 更新库存
    cursor.execute('''
        UPDATE inventory SET quantity = quantity - ? WHERE id = ?
    ''', (quantity, item_id))

    conn.commit()
    conn.close()
    return True, '出库成功'


def get_stock_in_records():
    """获取入库记录"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM stock_in ORDER BY created_at DESC')
    rows = cursor.fetchall()
    conn.close()
    return rows


def get_stock_out_records():
    """获取出库记录"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM stock_out ORDER BY created_at DESC')
    rows = cursor.fetchall()
    conn.close()
    return rows


def get_monthly_summary():
    """获取月度汇总"""
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute('''
        SELECT
            strftime('%Y-%m', created_at) as month,
            SUM(sell_price * quantity) as revenue,
            SUM(purchase_price * quantity) as cost,
            SUM(profit) as total_profit,
            SUM(quantity) as total_quantity
        FROM stock_out
        GROUP BY strftime('%Y-%m', created_at)
        ORDER BY month DESC
    ''')

    rows = cursor.fetchall()
    conn.close()
    return rows
