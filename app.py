from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
import database

app = Flask(__name__)
app.secret_key = 'stock_manager_secret_key'

CATEGORIES = ['衣服', '鞋子', '配件']


@app.route('/')
def index():
    """首页 - 库存展示"""
    category = request.args.get('category', 'all')
    inventory = database.get_inventory(category)
    total_value = database.get_total_value()
    return render_template('index.html',
                           inventory=inventory,
                           total_value=total_value,
                           categories=CATEGORIES,
                           selected_category=category)


@app.route('/stock_in', methods=['GET', 'POST'])
def stock_in():
    """入库页面"""
    if request.method == 'POST':
        category = request.form.get('category')
        product_code = request.form.get('product_code')
        size = request.form.get('size')
        purchase_price = float(request.form.get('purchase_price'))
        quantity = int(request.form.get('quantity'))

        database.add_stock(category, product_code, size, purchase_price, quantity)
        flash('入库成功！', 'success')
        return redirect(url_for('stock_in'))

    return render_template('stock_in.html', categories=CATEGORIES)


@app.route('/stock_out', methods=['GET', 'POST'])
def stock_out():
    """出库页面"""
    search_results = []
    search_code = request.args.get('search', '')

    if search_code:
        search_results = database.search_by_product_code(search_code)

    return render_template('stock_out.html',
                           search_results=search_results,
                           search_code=search_code)


@app.route('/do_stock_out', methods=['POST'])
def do_stock_out():
    """执行出库操作"""
    item_id = int(request.form.get('item_id'))
    sell_price = float(request.form.get('sell_price'))
    quantity = int(request.form.get('quantity'))

    success, message = database.remove_stock(item_id, sell_price, quantity)

    if success:
        flash(message, 'success')
    else:
        flash(message, 'error')

    return redirect(url_for('stock_out'))


@app.route('/records')
def records():
    """出入库记录"""
    stock_in_records = database.get_stock_in_records()
    stock_out_records = database.get_stock_out_records()
    return render_template('records.html',
                           stock_in_records=stock_in_records,
                           stock_out_records=stock_out_records)


@app.route('/monthly')
def monthly():
    """月度汇总"""
    summary = database.get_monthly_summary()
    return render_template('monthly.html', summary=summary)


@app.route('/api/item/<int:item_id>')
def get_item(item_id):
    """获取库存项详情API"""
    item = database.get_inventory_item(item_id)
    if item:
        return jsonify({
            'id': item['id'],
            'category': item['category'],
            'product_code': item['product_code'],
            'size': item['size'],
            'purchase_price': item['purchase_price'],
            'quantity': item['quantity']
        })
    return jsonify({'error': '未找到'}), 404


if __name__ == '__main__':
    database.init_db()
    app.run("0.0.0.0",debug=True, port=15000)
