# 📧 Microsoft Graph Outlook 邮件发送脚本

本项目通过 **Microsoft Graph API** 实现 Outlook / Hotmail 邮箱的 **纯文本邮件发送**，  
适用于需要向指定邮箱（例如自动化系统）发送完全纯文本信件的场景。  

---

## 🚀 功能简介

- 使用 Microsoft Graph API 进行安全 OAuth2 授权  
- 支持 **个人 Microsoft 帐户 (Outlook / Hotmail)**  
- 纯文本邮件（换行不会被编码为 `=0A`）  
- 自动替换行尾为 CRLF，兼容老旧系统

---

## 🧩 环境要求

- Python ≥ 3.10  
- 依赖：
```bash
pip install msal requests
```



⚙️ 配置 Microsoft Entra 应用

1️⃣ 打开 Microsoft Entra 管理中心(需注册azure,无论注册是否成功都行)

👉 https://entra.microsoft.com/

使用你的 Outlook / Hotmail 账户 登录。



2️⃣ 创建应用注册

在左侧菜单选择：

应用注册 (App registrations) → 新注册 (New registration)

填写如下内容：

项目	值
名称 (Name)	任意名称，如：PythonMailApp


受支持帐户类型 (Supported account types)	

选择 个人 Microsoft 帐户 (Personal accounts only)


重定向 URI (Redirect URI)	类型选 Web，值填入： http://localhost:8001


3️⃣ 获取关键信息

注册完成后，在应用概览页中可看到以下三项：

应用程序(客户端) ID|Application (client) ID  = `CLIENT_ID`

目录(租户) ID|Directory (tenant) ID         = `TENANT_ID`


客户端机密值|Client secret (5️⃣创建所得)      = `CLIENT_SECRET`


4️⃣ 配置权限

进入左侧菜单：

API 权限 (API permissions) → 添加权限 (Add a permission)

→ 选择 Microsoft Graph → 委托权限 (Delegated permissions)

勾选以下项目：

Mail.Send

Mail.ReadWrite

offline_access

点击 添加权限 (Add permissions) 




5️⃣ 创建客户端机密

在左侧菜单中点击：

证书和机密 (Certificates & secrets) → 新建客户端密码 (New client secret)

描述 (Description)：任意填写

过期时间：建议选“6个月”或“12个月”

保存后，系统会生成：

客户端机密(值) (Value) → 例如：Rzr8Q~xxxxxx~xxxxxxxxxxxx  = 3️⃣ `CLIENT_SECRET`

--------------------------------------------------------------------------------------
```bash
python out.py
```
