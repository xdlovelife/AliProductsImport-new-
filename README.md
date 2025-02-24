# 阿里巴巴产品导入工具

一个用于自动化从阿里巴巴导入产品到Shopify商店的工具。

## 功能特点

- 自动化产品导入流程
- 支持批量处理多个产品
- 自动分类管理
- 自动处理产品变体
- 自动处理产品图片
- 用户友好的图形界面
- 实时导入进度显示
- 详细的操作日志

## 系统要求

- Windows 10/11
- Google Chrome浏览器
- 网络连接

## 安装说明

1. 下载最新版本的程序
2. 解压到任意目录
3. 运行 `AliProductsImport.exe`

## 使用方法

1. 准备Excel文件
   - 第一列：产品链接或搜索关键词
   - Sheet名称：对应的产品分类名称

2. 配置设置
   - 点击"设置"菜单
   - 配置ChromeDriver（可自动下载或手动指定）
   - 设置Chrome用户数据目录（可使用默认目录）

3. 导入产品
   - 点击"浏览"选择Excel文件
   - 检查数据预览是否正确
   - 点击"开始导入"
   - 等待导入完成

## 注意事项

1. 首次使用需要：
   - 安装Chrome浏览器
   - 安装IMportify插件并链接到对应的客户端

2. 导入过程中：
   - 请勿关闭浏览器
   - 请勿操作正在运行的浏览器窗口
   - 保持网络连接稳定

3. 常见问题：
   - 如果自动下载ChromeDriver失败，可以：
     1. 使用淘宝镜像手动下载：
        - 访问 `https://registry.npmmirror.com/binary.html?path=chromedriver/`
        - 选择与Chrome版本匹配的ChromeDriver版本
        - 下载对应的zip文件并解压
     2. 或使用官方地址（需要代理）：
        - 访问 `https://chromedriver.chromium.org/downloads`
        - 下载对应版本
   - 如果出现网络问题，建议使用代理或VPN
   - 如果出现登录问题，请检查用户数据目录设置

4. 使用中国镜像：
   - 程序默认使用淘宝镜像源下载ChromeDriver
   - 如果自动下载失败，可以手动从淘宝镜像下载
   - ChromeDriver淘宝镜像地址：`https://registry.npmmirror.com/binary.html?path=chromedriver/`
   - 下载后将chromedriver.exe放在程序目录下即可

## 更新日志

### v1.0.0 (2024-03)
- 首次发布
- 支持基本的产品导入功能
- 提供图形用户界面
- 支持自动下载ChromeDriver
- 集成淘宝镜像源支持

## 技术支持

如有问题或建议，请联系：
- Email: xdlovelife@gmail.com

## 许可证

Copyright (C) 2024 XD Love Life. 保留所有权利。

## 免责声明

本工具仅用于提高工作效率，请遵守相关平台的使用规则和条款。对于因使用本工具而产生的任何问题，开发者不承担责任。 