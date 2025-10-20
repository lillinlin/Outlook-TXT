import json
import requests
import os
from msal import ConfidentialClientApplication
import webbrowser

# =============================
# Microsoft Graph API 配置
# =============================
CLIENT_ID = "需填写"
CLIENT_SECRET = "需填写"
TENANT_ID = "需填写"
AUTHORITY = "https://login.microsoftonline.com/consumers"   # 个人账号
REDIRECT_URI = "http://localhost:8001"

SCOPES = [
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/Mail.ReadWrite",
    "https://graph.microsoft.com/User.Read"
]

# =============================
# 邮件内容配置
# =============================
TO_ADDRESS = "需填写收件人"
MAIL_SUBJECT = "需填写邮箱主题/标题"

MAIL_BODY = """需填写内容"""

# =============================
# 邮件发送函数
# =============================
def send_mail(access_token):
    url = "https://graph.microsoft.com/v1.0/me/sendMail"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # 💡 关键：替换为 CRLF，防止 Graph 自动编码成 Quoted-Printable (=0A)
    safe_body = MAIL_BODY.replace("\r\n", "\n").replace("\n", "\r\n")

    message = {
        "message": {
            "subject": MAIL_SUBJECT,
            "body": {"contentType": "Text", "content": safe_body},
            "toRecipients": [{"emailAddress": {"address": TO_ADDRESS}}],
        },
        "saveToSentItems": "true"
    }

    response = requests.post(url, headers=headers, data=json.dumps(message))
    if response.status_code == 202:
        print("✅ 邮件发送成功！")
    else:
        print(f"❌ 发送失败：{response.status_code} - {response.text}")

# =============================
# 授权获取 Token
# =============================
def acquire_token():
    app = ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=AUTHORITY
    )

    # Step 1：生成授权 URL
    auth_url = app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )

    print("\n🔗 请在浏览器中打开以下链接并登录授权：")
    print(auth_url)
    webbrowser.open(auth_url)

    # Step 2：输入授权码
    auth_code = input("\n👉 登录完成后，请复制浏览器地址栏中 code= 后的那串授权码并粘贴到这里：\n> ").strip()

    # Step 3：用授权码换取访问令牌
    result = app.acquire_token_by_authorization_code(
        auth_code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI
    )

    if "access_token" in result:
        print("✅ 获取访问令牌成功。")
        return result
    else:
        print("❌ 登录失败：", r
