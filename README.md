# 销售预测捆绑SKU拆解 Streamlit 小程序

将跨境电商「销售预测」工作簿中的捆绑SKU按BOM拆解为单品SKU并汇总成明细。

## 功能

- 读取运营表 + BOM捆绑关系 + 单品产品资料
- 按BOM关系递归拆解捆绑SKU为单品SKU
- 提供库存文件时额外生成「销售需求汇总」Sheet（含整套替代、按月从老到新分配）
- 输出3个工作表：销售需求汇总 + 未拆解整套版销售需求 + 异常提醒

## 安装

```bash
pip install -r requirements.txt
```

依赖：
- streamlit
- openpyxl
- pywin32（读取DRM加密xlsx需要Windows + Excel）

## 运行

```bash
# 方式1：命令行启动
streamlit run app.py

# 方式2：双击启动
run.bat
```

浏览器自动打开 `http://localhost:8501`。

## 使用

1. 上传**运营表** xlsx（支持DRM加密）
2. 上传**同款套装库存** CSV + **产品资料** CSV（成对必填）
3. 可选：上传**库存文件** xlsx → 额外生成Sheet1「销售需求汇总」
4. 点击「开始拆解」
5. 完成后点击「下载结果文件」

## 参数说明

侧边栏可配置：
- **飞书URL**：整套替代关系确认表（默认内置）
- **运营表Sheet名**：留空自动识别
- **CSV编码**：默认自动探测

## 技术细节

- DRM加密文件通过 pywin32 + Excel COM 读取
- CSV编码自动探测 gbk → utf-8-sig → utf-8
- 飞书映射通过 lark-cli 读取（需已授权）
- 表格美化：蓝色表头、斑马行、老套装黄色/新套装蓝色背景、冻结窗格
