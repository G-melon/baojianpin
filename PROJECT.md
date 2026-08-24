# 保健品家庭内购平台

这是一个保健品家庭内购平台框架，包含顾客端商品浏览/下单和商家后台商品/订单管理。

## 当前状态

- 已根据 `内购产品明细--20260820.xlsx` 的 `260820-实` sheet 导入 15 个商品。
- 已提取 Excel 内嵌产品图片到 `uploads/products/`。
- 商品详情字段包含：商品名称、功效、服用细节、内购价、产品图片、生产日期、库存。
- 已将业务分类调整为：营养补充、维生素矿物质、膳食纤维、益生菌调理、心脑血管、儿童成长、抗炎护肝、中老年营养、美容养护、运动营养、礼盒套装。

## 主要文件

- `index.html`：顾客端商城与下单页面。
- `admin.html`：商家后台，支持商品新增/编辑/删除、上传图片、查看订单、导出 XLSX。
- `server.py`：Flask 后端 API。
- `data/products.json`：商品数据。
- `uploads/`：后台上传的商品图片。
- `tools/import_products_from_xlsx.py`：从当前 Excel 内购明细重新导入商品和图片的脚本。

## 启动

```bash
pip install -r requirements.txt
python server.py
```

访问：

- 顾客端：`http://localhost:8899/`
- 后台：`http://localhost:8899/admin`

