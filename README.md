# Keditor

一个为 Kindle 与其他墨水屏设备设计的轻量 Markdown 网页编辑器。正文不写入浏览器本地存储；连接 OneDrive 后，修改会自动保存至应用专用文件夹。

## 本地运行

```bash
npm install
npm run dev
```

## 配置 OneDrive

1. 在 Microsoft Entra 管理中心注册一个应用。
2. 将“支持的账户类型”设为组织目录与个人 Microsoft 账户。
3. 添加“单页应用程序（SPA）”平台，并添加网站地址作为重定向 URI；本地开发可使用 `http://localhost:3000` 或实际预览端口。
4. 添加 Microsoft Graph 委托权限 `Files.ReadWrite.AppFolder` 与 `User.Read`。
5. 复制 `.env.example` 为 `.env.local`，填写公开的 Application (client) ID：

```env
NEXT_PUBLIC_MICROSOFT_CLIENT_ID=你的客户端ID
```

不需要、也不要把 Microsoft 客户端密钥放到前端。

## 保存行为

- 停止输入 1.2 秒后保存到 OneDrive。
- 未连接、正在保存或保存失败时，界面会明确提示。
- 关闭存在未同步修改的页面时，浏览器会发出离开警告。
- 浏览器中仅使用会话存储保存微软登录状态，不保存 Markdown 正文。
