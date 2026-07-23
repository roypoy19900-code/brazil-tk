# GitHub Pages 部署指南

## 快速部署步骤

### 1. 准备文件

确保以下文件在同一目录中：
- index.html
- app.js
- README.md
- .gitignore

### 2. 创建GitHub仓库

1. 访问 https://github.com/new
2. 输入仓库名称，例如：`order-processor`
3. 选择 "Public"（公开）
4. 不要初始化README（我们已有自己的README）
5. 点击 "Create repository"

### 3. 上传文件

#### 方法A：通过网页界面（推荐新手）

1. 在仓库页面点击 "uploading an existing file"
2. 拖拽或选择所有文件（index.html, app.js, README.md, .gitignore）
3. 填写commit信息，例如："Initial commit - Order processor tool"
4. 点击 "Commit changes"

#### 方法B：使用Git命令

```bash
# 克隆仓库（替换为你的仓库地址）
git clone https://github.com/你的用户名/order-processor.git
cd order-processor

# 复制文件到此目录
cp /path/to/index.html .
cp /path/to/app.js .
cp /path/to/README.md .
cp /path/to/.gitignore .

# 提交并推送
git add .
git commit -m "Initial commit - Order processor tool"
git push origin main
```

### 4. 启用GitHub Pages

1. 进入仓库页面
2. 点击顶部导航栏的 "Settings"
3. 在左侧边栏找到 "Pages"（可能在 "Code and automation" 分类下）
4. 在 "Build and deployment" 部分：
   - Source: 选择 "Deploy from a branch"
   - Branch: 选择 "main" 
   - Folder: 选择 "/ (root)"
5. 点击 "Save"

### 5. 等待部署

- GitHub会自动构建和部署
- 通常需要1-3分钟
- 可以在 "Actions" 标签查看部署进度

### 6. 访问应用

部署完成后，你会看到类似这样的URL：
```
https://你的用户名.github.io/order-processor/
```

点击该链接即可访问应用！

## 验证部署

1. 打开上述URL
2. 应该能看到订单处理工具的界面
3. 尝试上传一个测试Excel文件
4. 确认功能正常工作

## 更新应用

如果需要修改代码：

```bash
# 修改文件后
git add .
git commit -m "Update: 描述你的修改"
git push origin main
```

GitHub Pages会在几分钟内自动重新部署。

## 自定义域名（可选）

如果想使用自己的域名：

1. 在仓库的 Settings > Pages 中找到 "Custom domain"
2. 输入你的域名，例如：`orders.yourdomain.com`
3. 在你的DNS服务商处添加CNAME记录指向 `你的用户名.github.io`
4. 等待DNS生效（可能需要几小时）

## 故障排除

**问题：页面显示404**
- 检查是否正确启用了GitHub Pages
- 确认文件名是 `index.html`（区分大小写）
- 等待几分钟让部署完成

**问题：SheetJS库加载失败**
- 检查网络连接
- CDN可能需要科学上网才能访问
- 可以考虑将SheetJS下载到本地

**问题：Excel文件无法处理**
- 确认文件格式是.xlsx或.xls
- 检查浏览器控制台是否有错误信息
- 确认Excel文件结构符合要求

## 安全提示

⚠️ **重要**：
- 本应用在浏览器本地运行，不会上传数据到服务器
- 但GitHub Pages是公开的，任何人都可以访问
- 如果处理敏感数据，建议：
  - 设置仓库为Private（私有）
  - 或使用其他部署方式（如Netlify、Vercel等支持私有部署的平台）
